Attribute VB_Name = "repBorradorDeta"
Option Explicit

'Global Const Version = "1.0"
'Global Const FechaModificacion = "26/12/2010"
'Global Const UltimaModificacion = "Diego Rosso "
'Global Const UltimaModificacion1 = ""
'Basado en la version 1.03 03/08/2009 -  'Encriptacion de string connection

Global Const Version = "1.1"
Global Const FechaModificacion = "19/10/2011"
Global Const UltimaModificacion = "Lisandro Moro "
Global Const UltimaModificacion1 = " Se agrego la validacion con sim_empleado."



'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
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
Global EmpErrores As Boolean

Global acum_01 As Integer
Global acum_02 As Integer
Global acum_03 As Integer
Global acum_04 As Integer

Global acum_Desc_01 As String
Global acum_Desc_02 As String
Global acum_Desc_03 As String
Global acum_Desc_04 As String

Global docum_tipo As Integer
Global docum_desc As String

Global tipoModelo As Integer

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String
Global empresa
Global emprNro
Global emprActiv
Global emprTer
Global emprDire
Global emprCuit

Global ii As Integer
Global acucontip(8) As String
Global acuconval(8) As Integer
Global acuconval2(8) As String
Global acuconetiq(8) As String
Global acuconmonto(8)



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
Dim proNro
Dim Ternro
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim acunroSueldo
Dim I
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim tituloReporte As String
Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden

'Dim listapronro2
'Dim historico1
'Dim historico2
'Dim ArrParametros

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
    
    Nombre_Arch = PathFLog & "ReporteBorradorDetaSIM" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    TiempoInicialProceso = GetTickCount
        
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
     Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
     Exit Sub
    End If
    
    HuboErrores = False
    
    On Error GoTo CE
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    'FGZ - 19/01/2007 - le cambié esto
    'cantRegistros = CInt(objRs!total)
    'por
    cantRegistros = CLng(objRs!total)
    
    totalEmpleados = cantRegistros
    
    objRs.Close
      
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "PID = " & PID
    Flog.writeline
    Flog.writeline "Inicio Proceso de Borrador Detallado Simulación: " & Now
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline

    Flog.writeline "Cambio el estado del proceso a Procesando"
    Flog.writeline
    
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
       listapronro = ArrParametros(0)
       
       'Obtengo el modelo a usar para obtener los datos
       tipoModelo = ArrParametros(1)
              
       'Obtengo los cortes de estructura
       tenro1 = CLng(ArrParametros(2))
       estrnro1 = CLng(ArrParametros(3))
       tenro2 = CLng(ArrParametros(4))
       estrnro2 = CLng(ArrParametros(5))
       tenro3 = CLng(ArrParametros(6))
       estrnro3 = CLng(ArrParametros(7))
       fecEstr = ArrParametros(8)
       
       'Obtengo el titulo del reporte
       tituloReporte = ArrParametros(9)
                      
'       If UBound(ArrParametros) > 10 Then
'            'Obtengo la lista de procesos historicos
'            listapronro2 = ArrParametros(10)
'
'            'Obtengo el valor del historico 1
'            historico1 = ArrParametros(11)
'
'            'Obtengo el valor del historico 2
'            historico2 = ArrParametros(12)
'       End If
       
       'EMPIEZA EL PROCESO
       
       'Busco en el confrep el numero de cuenta que se va a usar para
       ' buscar el valor del sueldo
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 60 "
       StrSql = StrSql & " AND ( "
       StrSql = StrSql & " ( confnrocol >=1 AND confnrocol <=4 ) OR"
       StrSql = StrSql & " ( confnrocol >=101 AND confnrocol <=104 )"
       StrSql = StrSql & " ) "
       
       OpenRecordset StrSql, objRs2
    
       acum_01 = 0
       acum_02 = 0
       acum_03 = 0
       acum_04 = 0
       
       acum_Desc_01 = ""
       acum_Desc_02 = ""
       acum_Desc_03 = ""
       acum_Desc_04 = ""
       
       For ii = 1 To 8
        acucontip(ii) = ""
        acuconval(ii) = 0
        acuconval2(ii) = ""
        acuconetiq(ii) = ""
        'acuconmonto(ii) = 0
       Next
       
        
       If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep para los AC"
          Exit Sub
       End If
       Flog.writeline "Obtengo los datos del confrep (AC)"
       Do Until objRs2.EOF
          Select Case objRs2!confnrocol
             Case 1
                acucontip(1) = objRs2!conftipo
                acuconval(1) = objRs2!confval
                If Not IsNull(objRs2!confval2) Then acuconval2(1) = objRs2!confval2
                acuconetiq(1) = objRs2!confetiq
             Case 2
                acucontip(2) = objRs2!conftipo
                acuconval(2) = objRs2!confval
                If Not IsNull(objRs2!confval2) Then acuconval2(2) = objRs2!confval2
                acuconetiq(2) = objRs2!confetiq
             Case 3
                acucontip(3) = objRs2!conftipo
                acuconval(3) = objRs2!confval
                If Not IsNull(objRs2!confval2) Then acuconval2(3) = objRs2!confval2
                acuconetiq(3) = objRs2!confetiq
             Case 4
                acucontip(4) = objRs2!conftipo
                acuconval(4) = objRs2!confval
                If Not IsNull(objRs2!confval2) Then acuconval2(4) = objRs2!confval2
                acuconetiq(4) = objRs2!confetiq
             Case 101
                acucontip(5) = objRs2!conftipo
                acuconval(5) = objRs2!confval
                If Not IsNull(objRs2!confval2) Then acuconval2(5) = objRs2!confval2
                acuconetiq(5) = objRs2!confetiq
             Case 102
                acucontip(6) = objRs2!conftipo
                acuconval(6) = objRs2!confval
                If Not IsNull(objRs2!confval2) Then acuconval2(6) = objRs2!confval2
                acuconetiq(6) = objRs2!confetiq
             Case 103
                acucontip(7) = objRs2!conftipo
                acuconval(7) = objRs2!confval
                If Not IsNull(objRs2!confval2) Then acuconval2(7) = objRs2!confval2
                acuconetiq(7) = objRs2!confetiq
             Case 104
                acucontip(8) = objRs2!conftipo
                acuconval(8) = objRs2!confval
                If Not IsNull(objRs2!confval2) Then acuconval2(8) = objRs2!confval2
                acuconetiq(8) = objRs2!confetiq
          End Select
          objRs2.MoveNext
       Loop
       objRs2.Close
     
        
        '----Busco el tipo de documento configurado en el confrep------
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " INNER JOIN tipodocu ON tipodocu.tidnro = confrep.confval "
       StrSql = StrSql & " WHERE repnro = 60 "
       StrSql = StrSql & "   AND conftipo = 'TD' "
       OpenRecordset StrSql, objRs2
       
       docum_tipo = 10
       docum_desc = "CUIL"
        
       If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep para TD"
          'Exit Sub
       End If
       
       Flog.writeline "Obtengo los datos del confrep (TD)"
       
       If Not objRs2.EOF Then
            docum_desc = objRs2!tidsigla
            docum_tipo = objRs2!confval
       End If
       objRs2.Close
       
        
       'Borro los datos del mismo proceso por si se corre mas de una vez
       StrSql = "DELETE FROM sim_rep_borradordeta WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
       StrSql = "DELETE FROM sim_rep_borrdeta_det WHERE bpronro = " & NroProceso
       objConn.Execute StrSql, , adExecuteNoRecords
       
       
       'Obtengo los empleados sobre los que tengo que generar los recibos
       CargarEmpleados NroProceso, rsEmpl
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       
       objConn.Execute StrSql, , adExecuteNoRecords
       
       arrpronro = Split(listapronro, ",")
       
       'Obtengo los datos de la empresa
       Call buscarDatosEmpresa(NroProceso, arrpronro(0))
        
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
             
             Call generarDatosEmpleado(proNro, Ternro, tituloReporte, orden)
             'Modificación para comparar historicos & simulacion
'             'Inicio
'             Select Case tipoModelo
'                Case 1
'                    Call generarDatosEmpleado(proNro, Ternro, tituloReporte, orden)
'                Case 2
'                    Call generarDatosComparacion(proNro, Ternro, tituloReporte, orden)
'             End Select
'            'Fin
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
              StrSql = StrSql & " AND ternro = " & Ternro
    
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
    
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    Flog.writeline "================================================================="
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline "================================================================="
End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function


'--------------------------------------------------------------------
' Se encarga de generar los datos para el Standar
'--------------------------------------------------------------------
Sub generarDatosEmpleado(proNro, Ternro, descripcion, orden)

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
Dim fecpago As String
Dim empfecalta
Dim fecalta As String
Dim fecbaja As String
Dim contrato As String
Dim categoria As String
Dim centrocosto As String
Dim Cuil As String
Dim estado
Dim ac01
Dim ac02
Dim ac03
Dim ac04
Dim accon05
Dim accon06
Dim accon07
Dim accon08
Dim profecpago
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim cliqnro
Dim pliqdesde As Date
Dim pliqhasta As Date

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
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empest, empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu,empfecbaja "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro
Flog.writeline "Buscando datos del empleado"
OpenRecordset StrSql, rsConsult
If rsConsult.EOF Then
    StrSql = " SELECT empest, empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu,empfecbaja "
    StrSql = StrSql & " FROM sim_empleado "
    StrSql = StrSql & " INNER JOIN tercero ON sim_empleado.ternro = tercero.ternro "
    StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
    OpenRecordset StrSql, rsConsult
End If

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
   
   'Fecha Alta
   If IsNull(rsConsult!empfaltagr) Then
       fecalta = ""
   Else
       fecalta = rsConsult!empfaltagr
   End If
    
   'Fecha Baja
   If IsNull(rsConsult!empfecbaja) Then
       fecbaja = ""
   Else
       fecbaja = rsConsult!empfecbaja
   End If
   
Else
   Flog.writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close



'------------------------------------------------------------------
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.*, sim_proceso.profecpago, sim_proceso.prodesc, sim_cabliq.cliqnro FROM periodo "
StrSql = StrSql & " INNER JOIN sim_proceso ON sim_proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND sim_proceso.pronro= " & proNro
StrSql = StrSql & " INNER JOIN sim_cabliq ON sim_proceso.pronro = sim_cabliq.pronro "
StrSql = StrSql & " AND sim_cabliq.empleado= " & Ternro

'---LOG---
Flog.writeline "Buscando datos del periodo"

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqmes = rsConsult!pliqmes
   pliqanio = rsConsult!pliqanio
   fecpago = rsConsult!profecpago
   profecpago = rsConsult!profecpago
   prodesc = rsConsult!prodesc
   pliqdesc = rsConsult!pliqdesc
   cliqnro = rsConsult!cliqnro
   pliqdesde = rsConsult!pliqdesde
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "Error al obtener los datos del periodo actual"
   GoTo Siguiente '27/09/2006 - Si habia dos proceso y el empleado pertenecia a solo uno, daba error
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos del centro de costo"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From sim_his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=sim_his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND sim_his_estructura.tenro=5 AND sim_his_estructura.ternro=" & Ternro
       
OpenRecordset StrSql, rsConsult

centrocosto = ""

If Not rsConsult.EOF Then
   centrocosto = rsConsult!estrdabr
Else
    Flog.writeline "No se encontro el centro de costo del empleado"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del contrato
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos del contrato"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From sim_his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=sim_his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND sim_his_estructura.tenro=18 AND sim_his_estructura.ternro=" & Ternro
       
OpenRecordset StrSql, rsConsult

contrato = ""

If Not rsConsult.EOF Then
   contrato = rsConsult!estrdabr
Else
   Flog.writeline "No se encontro el contrato del empleado"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la categoria
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos de la categoria"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From sim_his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=sim_his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND sim_his_estructura.tenro=3 AND sim_his_estructura.ternro=" & Ternro
       
OpenRecordset StrSql, rsConsult

categoria = ""

If Not rsConsult.EOF Then
   categoria = rsConsult!estrdabr
Else
   Flog.writeline "No se encontro la categoria del empleado"
'   GoTo MError
End If

rsConsult.Close
   
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos estructura 1"

If tenro1 <> 0 Then
    
    StrSql = " SELECT estrdabr, tedabr "
    StrSql = StrSql & " FROM sim_his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = sim_his_estructura.estrnro "
    StrSql = StrSql & "    AND sim_his_estructura.ternro = " & Ternro & " AND sim_his_estructura.tenro = " & tenro1
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro1 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro1
    End If
    
    StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = sim_his_estructura.tenro "
           
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
Flog.writeline "Buscando datos estructura 2"

If tenro2 <> 0 Then
    
    StrSql = " SELECT estrdabr, tedabr "
    StrSql = StrSql & " FROM sim_his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = sim_his_estructura.estrnro "
    StrSql = StrSql & "    AND sim_his_estructura.ternro = " & Ternro & " AND sim_his_estructura.tenro = " & tenro2
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro2 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro2
    End If
    
    StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = sim_his_estructura.tenro "
           
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
Flog.writeline "Buscando datos estructura 3"

If tenro3 <> 0 Then
    
    StrSql = " SELECT estrdabr, tedabr "
    StrSql = StrSql & " FROM sim_his_estructura "
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = sim_his_estructura.estrnro "
    StrSql = StrSql & "    AND sim_his_estructura.ternro = " & Ternro & " AND sim_his_estructura.tenro = " & tenro3
    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
    
    If estrnro3 <> 0 Then
        StrSql = StrSql & " AND estructura.estrnro = " & estrnro3
    End If
    
    StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = sim_his_estructura.tenro "
           
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       estrnomb3 = rsConsult!estrdabr
       tenomb3 = rsConsult!tedabr
    End If
End If
    
'------------------------------------------------------------------
'Busco el valor del fecha nac, cuil, doc y tipo doc
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos de los documentos"

Cuil = ""
sql = " SELECT ter_doc.nrodoc, tipodocu.tidsigla "
sql = sql & " FROM tercero "
sql = sql & " INNER JOIN ter_doc ON (tercero.ternro = ter_doc.ternro AND ter_doc.tidnro = " & docum_tipo & ") "
sql = sql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
sql = sql & " WHERE tercero.ternro= " & Ternro
OpenRecordset sql, rsConsult
If Not rsConsult.EOF Then
   Cuil = rsConsult!nrodoc
   docum_desc = rsConsult!tidsigla
   rsConsult.Close
Else
    rsConsult.Close
    sql = " SELECT ter_doc.nrodoc, tipodocu.tidsigla "
    sql = sql & " FROM tercero "
    sql = sql & " INNER JOIN ter_doc ON (tercero.ternro = ter_doc.ternro AND ter_doc.tidnro = 10) "
    sql = sql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
    sql = sql & " WHERE tercero.ternro= " & Ternro
    OpenRecordset sql, rsConsult
    If Not rsConsult.EOF Then
        Cuil = rsConsult!nrodoc
        docum_desc = rsConsult!tidsigla
    End If
    rsConsult.Close
End If


'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------
For ii = 1 To 8
    If acucontip(ii) = "AC" Then
        If acuconval(ii) = 0 Then
            acuconmonto(ii) = "NULL"
        Else
            sql = " SELECT almonto,acunro "
            sql = sql & " FROM sim_acu_liq"
            sql = sql & " WHERE acunro = " & acuconval(ii)
            sql = sql & " AND cliqnro =  " & cliqnro
            '---LOG---
            Flog.writeline "Buscando datos del acumulador 01"
            OpenRecordset sql, rsConsult
            If rsConsult.EOF Then
               acuconmonto(ii) = 0
            Else
               acuconmonto(ii) = rsConsult!almonto
            End If
            rsConsult.Close
        End If
    End If
    If acucontip(ii) = "CO" Then
        If acuconval2(ii) = "" Then
            acuconmonto(ii) = "NULL"
        Else
            sql = " SELECT sim_detliq.dlimonto "
            sql = sql & " FROM sim_detliq "
            sql = sql & " INNER JOIN concepto ON concepto.conccod = " & acuconval2(ii)
            sql = sql & " AND concepto.concnro = sim_detliq.concnro "
            sql = sql & " WHERE sim_detliq.cliqnro = " & cliqnro
            '---LOG---
            Flog.writeline "Buscando datos del concepto 01"
            OpenRecordset sql, rsConsult
            If rsConsult.EOF Then
               acuconmonto(ii) = 0
            Else
               acuconmonto(ii) = rsConsult!dlimonto
            End If
            rsConsult.Close
        End If
    End If
Next
'Acumulador 01
'If acum_01 = 0 Then
'    ac01 = "NULL"
'Else
'    sql = " SELECT almonto,acunro "
'    sql = sql & " FROM acu_liq"
'    sql = sql & " WHERE acunro = " & acum_01
'    sql = sql & " AND cliqnro =  " & cliqnro
'    '---LOG---
'    Flog.writeline "Buscando datos del acumulador 01"
'    OpenRecordset sql, rsConsult
'    If rsConsult.EOF Then
'       ac01 = 0
'    Else
'       ac01 = rsConsult!almonto
'    End If
'    rsConsult.Close
'End If


'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

StrSql = " INSERT INTO sim_rep_borradordeta "
StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
StrSql = StrSql & " empresa , emprnro, pliqnro, "
StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
StrSql = StrSql & " centrocosto , documento,"
StrSql = StrSql & " acumval1, acumdesc1, "
StrSql = StrSql & " acumval2, acumdesc2, "
StrSql = StrSql & " acumval3, acumdesc3, "
StrSql = StrSql & " acumval4, acumdesc4, "
StrSql = StrSql & " acumval5, acumdesc5, "
StrSql = StrSql & " acumval6, acumdesc6, "
StrSql = StrSql & " acumval7, acumdesc7, "
StrSql = StrSql & " acumval8, acumdesc8, "
StrSql = StrSql & " prodesc , descripcion, pliqdesc,  "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,tidsigla) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & Ternro
StrSql = StrSql & "," & proNro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & empresa & "'"
StrSql = StrSql & "," & emprNro
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & fecalta & "'"
StrSql = StrSql & ",'" & fecbaja & "'"
StrSql = StrSql & ",'" & contrato & "'"
StrSql = StrSql & ",'" & categoria & "'"
StrSql = StrSql & ",'" & centrocosto & "'"
StrSql = StrSql & ",'" & Cuil & "'"
For ii = 1 To 8
    StrSql = StrSql & "," & numberForSQL(acuconmonto(ii))
    StrSql = StrSql & ",'" & acuconetiq(ii) & "'"
Next
StrSql = StrSql & ",'" & prodesc & "'"
StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & docum_desc & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Busco el detalle de la liquidacion
'------------------------------------------------------------------

StrSql = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, sim_detliq.dlicant, sim_detliq.dlimonto " & _
    " FROM sim_cabliq " & _
    " INNER JOIN sim_proceso  ON sim_proceso.pronro = sim_cabliq.pronro AND sim_proceso.pronro = " & proNro & _
    " INNER JOIN periodo  ON sim_proceso.pliqnro = periodo.pliqnro " & _
    " INNER JOIN sim_detliq   ON sim_cabliq.cliqnro = sim_detliq.cliqnro  AND sim_cabliq.empleado = " & Ternro & _
    " INNER JOIN concepto ON concepto.concnro = sim_detliq.concnro AND concepto.concimp = -1 " & _
    " ORDER BY concepto.conccod "

'---LOG---
Flog.writeline "Buscando datos del detalle de liquidacion"

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF

  StrSql = " INSERT INTO sim_rep_borrdeta_det "
  StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
  StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
  StrSql = StrSql & " dlimonto) "
  StrSql = StrSql & " VALUES "
  StrSql = StrSql & "(" & NroProceso
  StrSql = StrSql & "," & Ternro
  StrSql = StrSql & "," & proNro
  StrSql = StrSql & ",'" & rsConsult!concabr & "'"
  StrSql = StrSql & ",'" & rsConsult!Conccod & "'"
  StrSql = StrSql & "," & rsConsult!concnro
  StrSql = StrSql & "," & rsConsult!concimp
  StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
  StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
  StrSql = StrSql & ")"
  
  objConn.Execute StrSql, , adExecuteNoRecords

  rsConsult.MoveNext
  
Loop

rsConsult.Close

Siguiente:

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
    StrEmpl = StrEmpl & " ORDER BY estado " '28/08/2006 - Lisandro Moro = Se agrego el orden
    OpenRecordset StrEmpl, rsEmpl
End Sub

Function numberForSQL(Str)
   
    If IsEmpty(Str) Or IsNull(Str) Then
        numberForSQL = "null"
    Else
        numberForSQL = Replace(Str, ",", ".")
    End If

End Function


Function strForSQL(Str)
   
  If IsNull(Str) Then
     strForSQL = "NULL"
  Else
     strForSQL = Str
  End If

End Function

Public Function Calcular_Edad(ByVal Fecha As Date) As Integer
'...........................................................................
' Archivo       : edad.i                              fecha ini. : 20/01/92
' Nombre progr. :
' tipo programa : FGZ
' Descripcion   :
'...........................................................................
Dim años  As Integer

    años = Year(Date) - Year(Fecha)
    If Month(Date) < Month(Fecha) Then
       años = años - 1
    Else
        If Month(Date) = Month(Fecha) Then
            If Day(Date) < Day(Fecha) Then
                años = años - 1
            End If
        End If
    End If
    Calcular_Edad = años
End Function


Sub buscarDatosEmpresa(NroProc, proNro)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim Ternro
Dim profecpago

    empresa = ""
    emprNro = 0
    emprActiv = ""
    emprTer = 0
    emprCuit = ""
    emprDire = ""

    ' -------------------------------------------------------------------------
    'Busco a un empleado para saber a que empresa pertenece
    ' -------------------------------------------------------------------------
    
    StrSql = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
       Ternro = 0
       Flog.writeline "Error: Buscando datos de la empresa: al obtener el empleado"
    Else
       Ternro = rsConsult!Ternro
    End If
    
    rsConsult.Close
    
    '------------------------------------------------------------------
    'Busco los datos del proceso
    '------------------------------------------------------------------
    StrSql = " SELECT * FROM sim_proceso "
    StrSql = StrSql & " WHERE sim_proceso.pronro= " & proNro
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       profecpago = rsConsult!profecpago
    Else
       Flog.writeline "Error: Buscando datos de la empresa: al obtener los datos del proceso"
    End If
    
    rsConsult.Close

    ' -------------------------------------------------------------------------
    ' Busco los datos de la empresa
    '--------------------------------------------------------------------------
    
    StrSql = "SELECT sim_his_estructura.estrnro, empresa.ternro, empresa.empnom,empresa.empactiv " & _
        " From sim_his_estructura " & _
        " INNER JOIN empresa ON empresa.estrnro = sim_his_estructura.estrnro " & _
        " WHERE sim_his_estructura.htetdesde <= " & ConvFecha(profecpago) & " AND " & _
        " (sim_his_estructura.htethasta >= " & ConvFecha(profecpago) & " OR sim_his_estructura.htethasta IS NULL)" & _
        " AND sim_his_estructura.ternro = " & Ternro & _
        " AND sim_his_estructura.tenro  = 10"
    
    '---LOG---
    Flog.writeline "Buscando datos de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    emprNro = 0
    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró la empresa"
        Exit Sub
    Else
        empresa = rsConsult!empnom
        emprNro = rsConsult!Estrnro
        emprActiv = rsConsult!empactiv
        emprTer = rsConsult!Ternro
    End If
    
    rsConsult.Close
    
    'Consulta para obtener la direccion de la empresa
    StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc From cabdom " & _
        " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & emprTer & _
        " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
    
    '---LOG---
    Flog.writeline "Buscando datos de la direccion de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el domicilio de la empresa"
        'Exit Sub
        emprDire = "   "
    Else
        emprDire = rsConsult!calle & " " & rsConsult!Nro & " - " & rsConsult!locdesc
    End If
   
    rsConsult.Close
    
    'Consulta para obtener el cuit de la empresa
    StrSql = "SELECT cuit.nrodoc FROM tercero " & _
             " INNER JOIN ter_doc cuit ON (tercero.ternro = cuit.ternro and cuit.tidnro = 6)" & _
             " Where tercero.ternro =" & emprTer
    
    '---LOG---
    Flog.writeline "Buscando datos del cuit de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el CUIT de la Empresa"
        'Exit Sub
        emprCuit = "  "
    Else
        emprCuit = rsConsult!nrodoc
    End If
    
    rsConsult.Close

End Sub

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

'Sub generarDatosEmpleadoHis(proNro, Ternro, descripcion, orden)
'
'Dim StrSql As String
'Dim rsConsult As New ADODB.Recordset
'
''Variables donde se guardan los datos del INSERT final
'
'Dim Legajo As Long
'Dim apellido As String
'Dim apellido2 As String
'Dim nombre As String
'Dim nombre2 As String
'Dim pliqnro
'Dim pliqmes
'Dim pliqanio
'Dim fecpago As String
'Dim empfecalta
'Dim fecalta As String
'Dim fecbaja As String
'Dim contrato As String
'Dim categoria As String
'Dim centrocosto As String
'Dim Cuil As String
'Dim estado
'Dim ac01
'Dim ac02
'Dim ac03
'Dim ac04
'Dim accon05
'Dim accon06
'Dim accon07
'Dim accon08
'Dim profecpago
'Dim prodesc
'Dim pliqdesc
'Dim pliqfecdep
'Dim pliqbco
'Dim cliqnro
'Dim pliqdesde As Date
'Dim pliqhasta As Date
'
'Dim EmpTernro
'
'Dim estrnomb1
'Dim estrnomb2
'Dim estrnomb3
'Dim tenomb1
'Dim tenomb2
'Dim tenomb3
'
'Dim sql As String
'
'On Error GoTo MError
'
'estrnomb1 = ""
'estrnomb2 = ""
'estrnomb3 = ""
'tenomb1 = ""
'tenomb2 = ""
'tenomb3 = ""
'
''------------------------------------------------------------------
''Busco los datos del empleado
''------------------------------------------------------------------
'StrSql = " SELECT empest, empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu,empfecbaja "
'StrSql = StrSql & " FROM empleado "
'StrSql = StrSql & " WHERE ternro= " & Ternro
'Flog.writeline "Buscando datos del empleado"
'OpenRecordset StrSql, rsConsult
'If rsConsult.EOF Then
'    StrSql = " SELECT empest, empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu,empfecbaja "
'    StrSql = StrSql & " FROM sim_his_empleado "
'    StrSql = StrSql & " INNER JOIN tercero ON sim_his_empleado.ternro = tercero.ternro "
'    StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
'    OpenRecordset StrSql, rsConsult
'End If
'
'If Not rsConsult.EOF Then
'   nombre = rsConsult!ternom
'   If IsNull(rsConsult!ternom2) Then
'      nombre2 = ""
'   Else
'      nombre2 = rsConsult!ternom2
'   End If
'   apellido = rsConsult!terape
'   If IsNull(rsConsult!terape2) Then
'      apellido2 = ""
'   Else
'      apellido2 = rsConsult!terape2
'   End If
'
'   Legajo = rsConsult!empleg
'
'   'Fecha Alta
'   If IsNull(rsConsult!empfaltagr) Then
'       fecalta = ""
'   Else
'       fecalta = rsConsult!empfaltagr
'   End If
'
'   'Fecha Baja
'   If IsNull(rsConsult!empfecbaja) Then
'       fecbaja = ""
'   Else
'       fecbaja = rsConsult!empfecbaja
'   End If
'
'Else
'   Flog.writeline "Error al obtener los datos del empleado"
''   GoTo MError
'End If
'
'rsConsult.Close
'
'
'
''------------------------------------------------------------------
''Busco los datos del periodo actual
''------------------------------------------------------------------
'StrSql = " SELECT periodo.*, sim_his_proceso.profecpago, sim_his_proceso.prodesc, sim_his_cabliq.cliqnro FROM periodo "
'StrSql = StrSql & " INNER JOIN sim_his_proceso ON sim_his_proceso.pliqnro = periodo.pliqnro "
'StrSql = StrSql & " AND sim_his_proceso.pronro= " & proNro
'StrSql = StrSql & " INNER JOIN sim_his_cabliq ON sim_his_proceso.pronro = sim_his_cabliq.pronro "
'StrSql = StrSql & " AND sim_his_cabliq.empleado= " & Ternro
'
''---LOG---
'Flog.writeline "Buscando datos del periodo"
'
'OpenRecordset StrSql, rsConsult
'
'If Not rsConsult.EOF Then
'   pliqnro = rsConsult!pliqnro
'   pliqmes = rsConsult!pliqmes
'   pliqanio = rsConsult!pliqanio
'   fecpago = rsConsult!profecpago
'   profecpago = rsConsult!profecpago
'   prodesc = rsConsult!prodesc
'   pliqdesc = rsConsult!pliqdesc
'   cliqnro = rsConsult!cliqnro
'   pliqdesde = rsConsult!pliqdesde
'   pliqhasta = rsConsult!pliqhasta
'Else
'   Flog.writeline "Error al obtener los datos del periodo actual"
'   GoTo Siguiente '27/09/2006 - Si habia dos proceso y el empleado pertenecia a solo uno, daba error
'End If
'
'rsConsult.Close
'
''------------------------------------------------------------------
''Busco el valor del centro de costo
''------------------------------------------------------------------
'
''---LOG---
'Flog.writeline "Buscando datos del centro de costo"
'
'StrSql = " SELECT estrdabr "
'StrSql = StrSql & " From sim_his_his_estructura"
'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=sim_his_his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND sim_his_his_estructura.tenro=5 AND sim_his_his_estructura.ternro=" & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'centrocosto = ""
'
'If Not rsConsult.EOF Then
'   centrocosto = rsConsult!estrdabr
'Else
'    Flog.writeline "No se encontro el centro de costo del empleado"
''   GoTo MError
'End If
'
'rsConsult.Close
'
''------------------------------------------------------------------
''Busco el valor del contrato
''------------------------------------------------------------------
'
''---LOG---
'Flog.writeline "Buscando datos del contrato"
'
'StrSql = " SELECT estrdabr "
'StrSql = StrSql & " From sim_his_his_estructura"
'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=sim_his_his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND sim_his_his_estructura.tenro=18 AND sim_his_his_estructura.ternro=" & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'contrato = ""
'
'If Not rsConsult.EOF Then
'   contrato = rsConsult!estrdabr
'Else
'   Flog.writeline "No se encontro el contrato del empleado"
''   GoTo MError
'End If
'
'rsConsult.Close
'
''------------------------------------------------------------------
''Busco el valor de la categoria
''------------------------------------------------------------------
'
''---LOG---
'Flog.writeline "Buscando datos de la categoria"
'
'StrSql = " SELECT estrdabr "
'StrSql = StrSql & " From sim_his_his_estructura"
'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=sim_his_his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND sim_his_estructura.tenro=3 AND sim_his_estructura.ternro=" & Ternro
'
'OpenRecordset StrSql, rsConsult
'
'categoria = ""
'
'If Not rsConsult.EOF Then
'   categoria = rsConsult!estrdabr
'Else
'   Flog.writeline "No se encontro la categoria del empleado"
''   GoTo MError
'End If
'
'rsConsult.Close
'
''------------------------------------------------------------------
''Busco los datos del tipos de estructura 1
''------------------------------------------------------------------
'
''---LOG---
'Flog.writeline "Buscando datos estructura 1"
'
'If tenro1 <> 0 Then
'
'    StrSql = " SELECT estrdabr, tedabr "
'    StrSql = StrSql & " FROM sim_his_estructura "
'    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = sim_his_estructura.estrnro "
'    StrSql = StrSql & "    AND sim_his_estructura.ternro = " & Ternro & " AND sim_his_estructura.tenro = " & tenro1
'    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
'
'    If estrnro1 <> 0 Then
'        StrSql = StrSql & " AND estructura.estrnro = " & estrnro1
'    End If
'
'    StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = sim_his_estructura.tenro "
'
'    OpenRecordset StrSql, rsConsult
'
'    If Not rsConsult.EOF Then
'       estrnomb1 = rsConsult!estrdabr
'       tenomb1 = rsConsult!tedabr
'    End If
'End If
'
'
''------------------------------------------------------------------
''Busco los datos del tipos de estructura 2
''------------------------------------------------------------------
'
''---LOG---
'Flog.writeline "Buscando datos estructura 2"
'
'If tenro2 <> 0 Then
'
'    StrSql = " SELECT estrdabr, tedabr "
'    StrSql = StrSql & " FROM sim_his_estructura "
'    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = sim_his_estructura.estrnro "
'    StrSql = StrSql & "    AND sim_his_estructura.ternro = " & Ternro & " AND sim_his_estructura.tenro = " & tenro2
'    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
'
'    If estrnro2 <> 0 Then
'        StrSql = StrSql & " AND estructura.estrnro = " & estrnro2
'    End If
'
'    StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = sim_his_estructura.tenro "
'
'    OpenRecordset StrSql, rsConsult
'
'    If Not rsConsult.EOF Then
'       estrnomb2 = rsConsult!estrdabr
'       tenomb2 = rsConsult!tedabr
'    End If
'End If
'
''------------------------------------------------------------------
''Busco los datos del tipos de estructura 3
''------------------------------------------------------------------
'
''---LOG---
'Flog.writeline "Buscando datos estructura 3"
'
'If tenro3 <> 0 Then
'
'    StrSql = " SELECT estrdabr, tedabr "
'    StrSql = StrSql & " FROM sim_his_estructura "
'    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = sim_his_estructura.estrnro "
'    StrSql = StrSql & "    AND sim_his_estructura.ternro = " & Ternro & " AND sim_his_estructura.tenro = " & tenro3
'    StrSql = StrSql & "    AND (htetdesde<=" & ConvFecha(fecEstr) & " AND (htethasta is null or htethasta>=" & ConvFecha(fecEstr) & "))"
'
'    If estrnro3 <> 0 Then
'        StrSql = StrSql & " AND estructura.estrnro = " & estrnro3
'    End If
'
'    StrSql = StrSql & "  INNER JOIN tipoestructura ON tipoestructura.tenro = sim_his_estructura.tenro "
'
'    OpenRecordset StrSql, rsConsult
'
'    If Not rsConsult.EOF Then
'       estrnomb3 = rsConsult!estrdabr
'       tenomb3 = rsConsult!tedabr
'    End If
'End If
'
''------------------------------------------------------------------
''Busco el valor del fecha nac, cuil, doc y tipo doc
''------------------------------------------------------------------
'
''---LOG---
'Flog.writeline "Buscando datos de los documentos"
'
'Cuil = ""
'sql = " SELECT ter_doc.nrodoc, tipodocu.tidsigla "
'sql = sql & " FROM tercero "
'sql = sql & " INNER JOIN ter_doc ON (tercero.ternro = ter_doc.ternro AND ter_doc.tidnro = " & docum_tipo & ") "
'sql = sql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
'sql = sql & " WHERE tercero.ternro= " & Ternro
'OpenRecordset sql, rsConsult
'If Not rsConsult.EOF Then
'   Cuil = rsConsult!nrodoc
'   docum_desc = rsConsult!tidsigla
'   rsConsult.Close
'Else
'    rsConsult.Close
'    sql = " SELECT ter_doc.nrodoc, tipodocu.tidsigla "
'    sql = sql & " FROM tercero "
'    sql = sql & " INNER JOIN ter_doc ON (tercero.ternro = ter_doc.ternro AND ter_doc.tidnro = 10) "
'    sql = sql & " INNER JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
'    sql = sql & " WHERE tercero.ternro= " & Ternro
'    OpenRecordset sql, rsConsult
'    If Not rsConsult.EOF Then
'        Cuil = rsConsult!nrodoc
'        docum_desc = rsConsult!tidsigla
'    End If
'    rsConsult.Close
'End If
'
'
''------------------------------------------------------------------
''Busco los datos de los acumuladores
''------------------------------------------------------------------
'For ii = 1 To 8
'    If acucontip(ii) = "AC" Then
'        If acuconval(ii) = 0 Then
'            acuconmonto(ii) = "NULL"
'        Else
'            sql = " SELECT almonto,acunro "
'            sql = sql & " FROM sim_acu_liq"
'            sql = sql & " WHERE acunro = " & acuconval(ii)
'            sql = sql & " AND cliqnro =  " & cliqnro
'            '---LOG---
'            Flog.writeline "Buscando datos del acumulador 01"
'            OpenRecordset sql, rsConsult
'            If rsConsult.EOF Then
'               acuconmonto(ii) = 0
'            Else
'               acuconmonto(ii) = rsConsult!almonto
'            End If
'            rsConsult.Close
'        End If
'    End If
'    If acucontip(ii) = "CO" Then
'        If acuconval2(ii) = "" Then
'            acuconmonto(ii) = "NULL"
'        Else
'            sql = " SELECT sim_detliq.dlimonto "
'            sql = sql & " FROM sim_detliq "
'            sql = sql & " INNER JOIN concepto ON concepto.conccod = " & acuconval2(ii)
'            sql = sql & " AND concepto.concnro = sim_detliq.concnro "
'            sql = sql & " WHERE sim_detliq.cliqnro = " & cliqnro
'            '---LOG---
'            Flog.writeline "Buscando datos del concepto 01"
'            OpenRecordset sql, rsConsult
'            If rsConsult.EOF Then
'               acuconmonto(ii) = 0
'            Else
'               acuconmonto(ii) = rsConsult!dlimonto
'            End If
'            rsConsult.Close
'        End If
'    End If
'Next
''Acumulador 01
''If acum_01 = 0 Then
''    ac01 = "NULL"
''Else
''    sql = " SELECT almonto,acunro "
''    sql = sql & " FROM acu_liq"
''    sql = sql & " WHERE acunro = " & acum_01
''    sql = sql & " AND cliqnro =  " & cliqnro
''    '---LOG---
''    Flog.writeline "Buscando datos del acumulador 01"
''    OpenRecordset sql, rsConsult
''    If rsConsult.EOF Then
''       ac01 = 0
''    Else
''       ac01 = rsConsult!almonto
''    End If
''    rsConsult.Close
''End If
'
'
''------------------------------------------------------------------
''Armo la SQL para guardar los datos
''------------------------------------------------------------------
'
'StrSql = " INSERT INTO sim_rep_borradordeta "
'StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
'StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
'StrSql = StrSql & " empresa , emprnro, pliqnro, "
'StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
'StrSql = StrSql & " centrocosto , documento,"
'StrSql = StrSql & " acumval1, acumdesc1, "
'StrSql = StrSql & " acumval2, acumdesc2, "
'StrSql = StrSql & " acumval3, acumdesc3, "
'StrSql = StrSql & " acumval4, acumdesc4, "
'StrSql = StrSql & " acumval5, acumdesc5, "
'StrSql = StrSql & " acumval6, acumdesc6, "
'StrSql = StrSql & " acumval7, acumdesc7, "
'StrSql = StrSql & " acumval8, acumdesc8, "
'StrSql = StrSql & " prodesc , descripcion, pliqdesc,  "
'StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,tidsigla) "
'StrSql = StrSql & " VALUES "
'StrSql = StrSql & "(" & NroProceso
'StrSql = StrSql & "," & Legajo
'StrSql = StrSql & "," & Ternro
'StrSql = StrSql & "," & proNro
'StrSql = StrSql & ",'" & apellido & "'"
'StrSql = StrSql & ",'" & apellido2 & "'"
'StrSql = StrSql & ",'" & nombre & "'"
'StrSql = StrSql & ",'" & nombre2 & "'"
'StrSql = StrSql & ",'" & empresa & "'"
'StrSql = StrSql & "," & emprNro
'StrSql = StrSql & "," & pliqnro
'StrSql = StrSql & ",'" & fecalta & "'"
'StrSql = StrSql & ",'" & fecbaja & "'"
'StrSql = StrSql & ",'" & contrato & "'"
'StrSql = StrSql & ",'" & categoria & "'"
'StrSql = StrSql & ",'" & centrocosto & "'"
'StrSql = StrSql & ",'" & Cuil & "'"
'For ii = 1 To 8
'    StrSql = StrSql & "," & numberForSQL(acuconmonto(ii))
'    StrSql = StrSql & ",'" & acuconetiq(ii) & "'"
'Next
'StrSql = StrSql & ",'" & prodesc & "'"
'StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
'StrSql = StrSql & ",'" & pliqdesc & "'"
'StrSql = StrSql & "," & controlNull(estrnomb1)
'StrSql = StrSql & "," & controlNull(estrnomb2)
'StrSql = StrSql & "," & controlNull(estrnomb3)
'StrSql = StrSql & "," & controlNull(tenomb1)
'StrSql = StrSql & "," & controlNull(tenomb2)
'StrSql = StrSql & "," & controlNull(tenomb3)
'StrSql = StrSql & "," & orden
'StrSql = StrSql & ",'" & docum_desc & "'"
'StrSql = StrSql & ")"
'
''------------------------------------------------------------------
''Guardo los datos en la BD
''------------------------------------------------------------------
'
'objConn.Execute StrSql, , adExecuteNoRecords
'
''------------------------------------------------------------------
''Busco el detalle de la liquidacion
''------------------------------------------------------------------
'
'StrSql = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, sim_detliq.dlicant, sim_detliq.dlimonto " & _
'    " FROM sim_cabliq " & _
'    " INNER JOIN sim_proceso  ON sim_proceso.pronro = sim_cabliq.pronro AND sim_proceso.pronro = " & proNro & _
'    " INNER JOIN periodo  ON sim_proceso.pliqnro = periodo.pliqnro " & _
'    " INNER JOIN sim_detliq   ON sim_cabliq.cliqnro = sim_detliq.cliqnro  AND sim_cabliq.empleado = " & Ternro & _
'    " INNER JOIN concepto ON concepto.concnro = sim_detliq.concnro AND concepto.concimp = -1 " & _
'    " ORDER BY concepto.conccod "
'
''---LOG---
'Flog.writeline "Buscando datos del detalle de liquidacion"
'
'OpenRecordset StrSql, rsConsult
'
'Do Until rsConsult.EOF
'
'  StrSql = " INSERT INTO sim_rep_borrdeta_det "
'  StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
'  StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
'  StrSql = StrSql & " dlimonto) "
'  StrSql = StrSql & " VALUES "
'  StrSql = StrSql & "(" & NroProceso
'  StrSql = StrSql & "," & Ternro
'  StrSql = StrSql & "," & proNro
'  StrSql = StrSql & ",'" & rsConsult!concabr & "'"
'  StrSql = StrSql & ",'" & rsConsult!Conccod & "'"
'  StrSql = StrSql & "," & rsConsult!concnro
'  StrSql = StrSql & "," & rsConsult!concimp
'  StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
'  StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
'  StrSql = StrSql & ")"
'
'  objConn.Execute StrSql, , adExecuteNoRecords
'
'  rsConsult.MoveNext
'
'Loop
'
'rsConsult.Close
'
'Siguiente:
'
'Exit Sub
'
'MError:
'    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
'    HuboErrores = True
'    EmpErrores = True
'    Exit Sub
'End Sub




''--------------------------------------------------------------------
'' Se encarga de generar los datos para comparar
''--------------------------------------------------------------------
'Sub generarDatosComparacion(proNro, Ternro, descripcion, orden)
'
'If historico1 = 0 And historico2 = 0 Then
'    'solo para simulacion
'    Call generarDatosEmpleado(proNro, Ternro, descripcion, orden)
'End If
'
'If historico1 > 0 And historico2 > 0 Then
''solo para historico
'End If
'
'If historico1 = 0 And historico2 > 0 Then
''para simulacion & historico
'End If
'
'If historico1 > 0 And historico2 = 0 Then
''para historico & simulacion
'End If
'
'
'End Sub

