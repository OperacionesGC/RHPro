Attribute VB_Name = "repLibroLey"
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

Global acum_msr As Integer
Global acum_neto As Integer
Global acum_basico As Integer
Global acum_asi_flia As Integer
Global acum_Dtos As Integer
Global acum_bruto As Integer
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
Global TipDocNroInscr As Long
Global empresa
Global emprNro
Global emprActiv
Global emprTer
Global emprDire
Global emprCuit
Global zonaDomicilio
Global concFamiliar01
Global concFamiliar02
Global param_empresa
Global listapronro
Global l_orden
Global filtro
Global totalEmpleados
Global cantRegistros


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
Dim ord

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
    
    Nombre_Arch = PathFLog & "ReporteLibroLey" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.Writeline "Inicio Proceso de Libro Ley : " & Now
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
       
       'Obtengo el modelo a usar para obtener los datos
       tipoModelo = ArrParametros(1)
       
       'Obtengo el nro de la ultima pagina impresa
       Pagina = CLng(ArrParametros(2))
       
       'Obtengo el tipo de familiares
       tipoFamiliares = ArrParametros(3)
       
       'Obtengo los cortes de estructura
       tenro1 = CInt(ArrParametros(4))
       estrnro1 = CInt(ArrParametros(5))
       tenro2 = CInt(ArrParametros(6))
       estrnro2 = CInt(ArrParametros(7))
       tenro3 = CInt(ArrParametros(8))
       estrnro3 = CInt(ArrParametros(9))
       fecEstr = ArrParametros(10)
       
       'Obtengo el titulo del reporte
       tituloReporte = ArrParametros(11)
       
       'Obtengo el nro. de Empresa
       ' Si param_empresa = "", entonces se ejecuta el proceso normalmente (se obtiene la empresa con buscarDatosEmpresa original)
       ' Si param_empresa = 0, entonces se ejecuta el proceso para todas las empresas
       ' Si param_empresa = 1169 (un estrnro cualquiera), entonces se ejecuta el proceso para esa empresa en particular
       param_empresa = ArrParametros(12)
       
       'Obtengo el filtro del reporte
       filtro = ArrParametros(13)
       
       'Obtengo el orden de la busqueda del reporte
       l_orden = ArrParametros(14)
       
       
       'EMPIEZA EL PROCESO
       
       'Busco en el confrep el numero de cuenta que se va a usar para
       ' buscar el valor del sueldo
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 61 "
      
       OpenRecordset StrSql, objRs2
       
       acum_msr = 0
       acum_neto = 0
       acum_basico = 0
       acum_asi_flia = 0
       acum_Dtos = 0
       acum_bruto = 0
       
       If objRs2.EOF Then
          Flog.Writeline "No esta configurado el ConfRep"
          Exit Sub
       End If
       
       Flog.Writeline "Obtengo los datos del confrep"
       
       TipDocNroInscr = 0
       NroBusqProg = 0
       NroBusqProg2 = 0
       zonaDomicilio = 3
       concFamiliar01 = 0
       concFamiliar02 = 0
       
       cantAcumGrupo1 = 0
       cantAcumGrupo2 = 0
       cantAcumGrupo3 = 0
       
       Do Until objRs2.EOF
       
          Select Case objRs2!confnrocol
             Case 1
                  acum_msr = objRs2!confval
             Case 2
                  acum_neto = objRs2!confval
             Case 3
                  acum_basico = objRs2!confval
             Case 4
                  acum_asi_flia = objRs2!confval
             Case 5
                  acum_Dtos = objRs2!confval
             Case 6
                  acum_bruto = objRs2!confval
             Case 10
                  NroBusqProg = objRs2!confval
             Case 11
                  TipDocNroInscr = objRs2!confval
             Case 12
                  NroBusqProg2 = objRs2!confval
             Case 15
                  zonaDomicilio = objRs2!confval
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
                  
             Case 50
                'acumGrupo 1
                cantAcumGrupo1 = cantAcumGrupo1 + 1
                acumGrupo1(cantAcumGrupo1) = objRs2!confval
             Case 51
                'acumGrupo 2
                cantAcumGrupo2 = cantAcumGrupo2 + 1
                acumGrupo2(cantAcumGrupo2) = objRs2!confval
             Case 52
                'acumGrupo 3
                cantAcumGrupo3 = cantAcumGrupo3 + 1
                acumGrupo3(cantAcumGrupo3) = objRs2!confval
                  
          End Select
       
          objRs2.MoveNext
       Loop

       'Busco el tipo de cada concepto
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 61 "

       OpenRecordset StrSql, objRs2
       
       Do Until objRs2.EOF
           
           Select Case objRs2!conftipo
              'Remunerativo
              Case "RE"
                 arrTipoConc(objRs2!confval) = 1
              'No Remunerativo
              Case "NR"
                 arrTipoConc(objRs2!confval) = 2
              'Descuento
              Case "DS"
                 arrTipoConc(objRs2!confval) = 3
           End Select
        
           objRs2.MoveNext
       Loop
        
       objRs2.Close
        
       'Obtengo los empleados sobre los que tengo que generar los recibos
       If param_empresa = "" Then
           CargarEmpleados NroProceso, rsEmpl, 0
            StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                        ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                        ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
            
            objConn.Execute StrSql, , adExecuteNoRecords
       End If
       
       
       'Obtengo la lista de procesos
       arrpronro = Split(listapronro, ",")
       
       'Chequeo el parametro empresa y ejecuto de acuerdo a su valor
       If param_empresa = "" Then
       ' Ejecuto el procedimiento normalmente (sin considerar empresa)
          
               'Obtengo los datos de la empresa
                Call buscarDatosEmpresa(NroProceso, arrpronro(0))
       
               'Genero por cada empleado un registro
               Do Until rsEmpl.EOF
                  arrpronro = Split(listapronro, ",")
                  EmpErrores = False
                  ternro = rsEmpl!ternro
                  orden = rsEmpl!estado
                  
                  'Genero una entrada para el empleado por cada proceso
                  For i = 0 To UBound(arrpronro)
                     proNro = arrpronro(i)
                     Flog.Writeline "Generando datos empleado " & ternro & " para el proceso " & proNro
                     
                     'De acuerdo al modelo es la forma de calcular los datos
                     Select Case tipoModelo
                        'Modelo usado para estantar y Deloitte
                        Case 1
                           Call generarDatosEmpleado01(proNro, ternro, tituloReporte, orden)
                        
                        'Modelo usado para Temaiken
                        Case 2
                           Call generarDatosEmpleado02(proNro, ternro, tituloReporte, orden)
                        
                        'Modelo usado para Accor
                        Case 3
                           Call generarDatosEmpleado03(proNro, ternro, tituloReporte, orden)
                        
                        'Modelo usado para Citrusvil
                        Case 4
                           Call generarDatosEmpleado04(proNro, ternro, tituloReporte, orden)
                     
                        'Modelo usado para Jugos
                        Case 5
                           Call generarDatosEmpleado05(proNro, ternro, tituloReporte, orden)
                     
                        'Modelo usado para Roche
                        Case 6
                           Call generarDatosEmpleado06(proNro, ternro, tituloReporte, orden)
                     
                        'Modelo usado para Teleperformance
                        Case 7
                           Call generarDatosEmpleado07(proNro, ternro, tituloReporte, orden)
                     
                        'Modelo usado para Promofilm
                        Case 8
                           Call generarDatosEmpleado08(proNro, ternro, tituloReporte, orden)
                     End Select
                     
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
       ' param_empresa <> ""
               
               StrSql = "SELECT * FROM empresa "
               If CLng(param_empresa) > 0 Then
                    StrSql = StrSql & "WHERE estrnro = " & CLng(param_empresa)
               End If
               StrSql = StrSql & "ORDER BY estrnro "
            
               OpenRecordset StrSql, rsEmpresas
       
                ' Actualizo el estado del proceso
               If CLng(param_empresa) > 0 Then
                      StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
            
                   objConn.Execute StrSql, , adExecuteNoRecords
               Else
               
                      cantRegistros = rsEmpresas.RecordCount
                      totalEmpleados = cantRegistros
                    
                      StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
            
                   objConn.Execute StrSql, , adExecuteNoRecords
 
               End If
               
               Do Until rsEmpresas.EOF
       
                    'Obtengo los datos de la empresa
                    Call buscarDatosEmpresaDada(rsEmpresas!estrnro, rsEmpresas!ternro, rsEmpresas!empnom, rsEmpresas!empactiv)
                    
                    ' Cargo los empleados para la empresa
                    CargarEmpleados 0, rsEmpl, CLng(rsEmpresas!estrnro)
                    
                    ord = 0
        
                    'Genero por cada empleado un registro
                    Do Until rsEmpl.EOF
                       arrpronro = Split(listapronro, ",")
                       EmpErrores = False
                       ternro = rsEmpl!ternro
                       orden = ord
                       
                       'Genero una entrada para el empleado por cada proceso
                       For i = 0 To UBound(arrpronro)
                          proNro = arrpronro(i)
                          Flog.Writeline "Generando datos empleado " & ternro & " para el proceso " & proNro
                          
                          'De acuerdo al modelo es la forma de calcular los datos
                          Select Case tipoModelo
                             'Modelo usado para estantar y Deloitte
                             Case 1
                                Call generarDatosEmpleado01(proNro, ternro, tituloReporte, orden)
                             
                             'Modelo usado para Temaiken
                             Case 2
                                Call generarDatosEmpleado02(proNro, ternro, tituloReporte, orden)
                             
                             'Modelo usado para Accor
                             Case 3
                                Call generarDatosEmpleado03(proNro, ternro, tituloReporte, orden)
                             
                             'Modelo usado para Citrusvil
                             Case 4
                                Call generarDatosEmpleado04(proNro, ternro, tituloReporte, orden)
                          
                             'Modelo usado para Jugos
                             Case 5
                                Call generarDatosEmpleado05(proNro, ternro, tituloReporte, orden)
                          
                             'Modelo usado para Roche
                             Case 6
                                Call generarDatosEmpleado06(proNro, ternro, tituloReporte, orden)
                          
                             'Modelo usado para Teleperformance
                             Case 7
                                Call generarDatosEmpleado07(proNro, ternro, tituloReporte, orden)
                          
                             'Modelo usado para Promofilm
                             Case 8
                                Call generarDatosEmpleado08(proNro, ternro, tituloReporte, orden)
                          End Select
                          
                       Next
                       
                       'Actualizo el estado del proceso
                       TiempoAcumulado = GetTickCount
                       
                       cantRegistros = cantRegistros - 1
                       
                       StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                                ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                                ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                          
                       objConn.Execute StrSql, , adExecuteNoRecords
                       
                       ord = ord + 1
                       rsEmpl.MoveNext
                    Loop
                    
                    rsEmpresas.MoveNext
               
               Loop
        
               rsEmpresas.Close
                       
       End If ' (param_empresa = "")
    
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
Dim fecpago As String
Dim empfecalta
Dim fecalta As String
Dim fecbaja As String
Dim contrato As String
Dim categoria As String
Dim direccion As String
Dim puesto As String
Dim documento  As String
Dim fecha_nac As String
Dim est_civil As String
Dim Cuil As String
Dim estado
Dim reg_prev As String
Dim lug_trab  As String
Dim basico As String
Dim neto As String
Dim msr As String
Dim asi_flia As String
Dim dtos As String
Dim bruto As String
Dim profecpago
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim terfecnac
Dim famfecvto
Dim famDGIdesde
Dim famDGIhasta
Dim cliqnro
Dim GuardarFam As Boolean
Dim pliqdesde As Date
Dim pliqhasta As Date

Dim centroCosto
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
   empfecalta = rsConsult!empfaltagr
Else
   Flog.Writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close

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
   fecpago = rsConsult!profecpago
   profecpago = rsConsult!profecpago
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
' Obtengo los datos de las estructuras
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de las estructuras"

 sql = " SELECT empfaltagr, empleado.empest, empfecbaja,estr1.estrdabr AS categoria, estr2.estrdabr AS contrato, estr3.estrdabr AS puesto, estr4.estrdabr AS regprev, estr5.estrdabr AS sucursal "
 sql = sql & " FROM empleado "
 'Busco la categoria
 sql = sql & " LEFT JOIN his_estructura his1 ON his1.ternro = empleado.ternro AND his1.tenro = 3 AND empleado.ternro= " & ternro & " AND his1.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his1.htethasta IS NULL OR his1.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr1 ON estr1.estrnro = his1.estrnro "
 'Busco el contrato
 sql = sql & " LEFT JOIN his_estructura his2 ON his2.ternro = empleado.ternro AND his2.tenro = 18 AND empleado.ternro= " & ternro & " AND his2.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his2.htethasta IS NULL OR his2.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr2 ON estr2.estrnro = his2.estrnro "
 'Busco el puesto
 sql = sql & " LEFT JOIN his_estructura his3 ON his3.ternro = empleado.ternro AND his3.tenro = 4 AND empleado.ternro= " & ternro & " AND his3.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his3.htethasta IS NULL OR his3.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr3 ON estr3.estrnro = his3.estrnro "
 'Busco el regimen previsional
 sql = sql & " LEFT JOIN his_estructura his4 ON his4.ternro = empleado.ternro AND his4.tenro = 15 AND empleado.ternro= " & ternro & " AND his4.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his4.htethasta IS NULL OR his4.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr4 ON estr4.estrnro = his4.estrnro "
 'Busco la sucursal
 sql = sql & " LEFT JOIN his_estructura his5 ON his5.ternro = empleado.ternro AND his5.tenro = 1 AND empleado.ternro= " & ternro & " AND his5.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his5.htethasta IS NULL OR his5.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr5 ON estr5.estrnro = his5.estrnro "
    
 sql = sql & " WHERE empleado.ternro = " & ternro

 OpenRecordset sql, rsConsult

 'Fecha Alta
 If IsNull(rsConsult!empfaltagr) Then
    fecalta = ""
 Else
    fecalta = rsConsult!empfaltagr
 End If
 
 'Contrato
 If IsNull(rsConsult!contrato) Then
   contrato = ""
 Else
   contrato = rsConsult!contrato
 End If
 
 'Categoria
 If IsNull(rsConsult!categoria) Then
   categoria = ""
 Else
   categoria = rsConsult!categoria
 End If
 
 'Puesto
 If IsNull(rsConsult!puesto) Then
   puesto = ""
 Else
   puesto = rsConsult!puesto
 End If
 
 'Reg. Prev.
 If IsNull(rsConsult!regprev) Then
   reg_prev = ""
 Else
   reg_prev = rsConsult!regprev
 End If
 
 'Sucursal
 If IsNull(rsConsult!sucursal) Then
   lug_trab = ""
 Else
   lug_trab = rsConsult!sucursal
 End If

 estado = rsConsult!empest
 
 rsConsult.Close
 
 
 ' -------------------------------------------------------------------------
' Busco la fecha de baja
'--------------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de fecha de baja"

 StrSql = " SELECT altfec, bajfec "
 StrSql = StrSql & " FROM fases "
 StrSql = StrSql & " WHERE real = -1 AND fases.empleado=" & ternro & " order by altfec DESC"

 OpenRecordset StrSql, rsConsult

 If rsConsult.EOF Then
      fecbaja = ""
  Else
    If Not IsNull(rsConsult!bajfec) Then
       fecbaja = rsConsult!bajfec
    End If
 End If
 
 rsConsult.Close
 
'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos del centro de costo"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult

centroCosto = ""

If Not rsConsult.EOF Then
   centroCosto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
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
'Busco el valor del fecha nac, cuil, doc y tipo doc
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de los documentos"

sql = " SELECT tercero.terfecnac, cuil.nrodoc AS nrocuil, docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc, estcivil.estcivdesabr"
sql = sql & " FROM tercero "
sql = sql & " LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
sql = sql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
sql = sql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
sql = sql & " LEFT JOIN estcivil ON estcivil.estcivnro= tercero.estcivnro "
sql = sql & " WHERE tercero.ternro= " & ternro

OpenRecordset sql, rsConsult

Cuil = ""
documento = ""
fecha_nac = ""
est_civil = 0

If Not rsConsult.EOF Then
   Cuil = rsConsult!nrocuil
   documento = rsConsult!sigladoc & "-" & rsConsult!nrodoc
   fecha_nac = rsConsult!terfecnac
   est_civil = rsConsult!estcivdesabr
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------

'Basico
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_basico
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del basico"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   basico = 0
Else
   basico = rsConsult!almonto
End If

rsConsult.Close

'Neto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_neto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del neto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   neto = 0
Else
   neto = rsConsult!almonto
End If

rsConsult.Close

'MSR
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_msr
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del msr"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   msr = 0
Else
   msr = rsConsult!almonto
End If

rsConsult.Close

'Asig. Familia
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_asi_flia
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de la asig. familia"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   asi_flia = 0
Else
   asi_flia = rsConsult!almonto
End If

rsConsult.Close

'Descuentos
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_Dtos
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de los descuentos"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   dtos = 0
Else
   dtos = rsConsult!almonto
End If

rsConsult.Close

'Bruto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_bruto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del bruto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   bruto = 0
Else
   bruto = rsConsult!almonto
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la direccion
'------------------------------------------------------------------

direccion = ""

Select Case zonaDomicilio
   'Direccion de la sucursal
   Case 1

        sql = " SELECT detdom.calle,detdom.nro,localidad.locdesc"
        sql = sql & " From his_estructura"
        sql = sql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND tenro=1 AND his_estructura.ternro=" & ternro
        sql = sql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
        sql = sql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        sql = sql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
    
        '---LOG---
        Flog.Writeline "Buscando datos de la direccion de la sucursal"
        
        OpenRecordset sql, rsConsult
        
        If Not rsConsult.EOF Then
           direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
        End If
        
        rsConsult.Close
    
  'Direccion de la empresa
  Case 2

        sql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
            " From his_estructura" & _
            " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
            " WHERE his_estructura.htetdesde <=" & ConvFecha(fecEstr) & " AND " & _
            " (his_estructura.htethasta >= " & ConvFecha(fecEstr) & " OR his_estructura.htethasta IS NULL)" & _
            " AND his_estructura.ternro = " & ternro & _
            " AND his_estructura.tenro  = 10"
        
        OpenRecordset sql, rsConsult
        
        EmpTernro = 0
        
        If Not rsConsult.EOF Then
            EmpTernro = rsConsult!ternro
        End If
        
        rsConsult.Close
        
        'Consulta para obtener la direccion de la empresa
        sql = "SELECT detdom.calle,detdom.nro,localidad.locdesc From cabdom " & _
            " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & EmpTernro & _
            " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
        
        '---LOG---
        Flog.Writeline "Buscando datos de la direccion de la empresa"
        
        OpenRecordset sql, rsConsult
        
        If Not rsConsult.EOF Then
           direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
        End If
        
        rsConsult.Close
    
  'Direccion del empleado
  Case 3

        sql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,detdom.piso,detdom.oficdepto,detdom.barrio"
        sql = sql & " FROM  cabdom "
        sql = sql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        sql = sql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
        sql = sql & " WHERE  cabdom.domdefault = -1 AND cabdom.ternro = " & ternro
       
        '---LOG---
        Flog.Writeline "Buscando datos de la direccion del empleado"
        
        OpenRecordset sql, rsConsult
        
        If Not rsConsult.EOF Then
           direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
        End If
        
        rsConsult.Close
    
End Select


'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & emprNro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!banco
Else
   pliqfecdep = ""
   pliqbco = ""
   Flog.Writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

If IsNull(pliqfecdep) Then
   pliqfecdep = "NULL"
Else
   If Trim(pliqfecdep) = "" Then
      pliqfecdep = "NULL"
   Else
      pliqfecdep = ConvFecha(pliqfecdep)
   End If
End If

StrSql = " INSERT INTO rep_libroley "
StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
StrSql = StrSql & " empresa , emprnro, pliqnro, fecpago,"
StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
StrSql = StrSql & " direccion , puesto, documento, fecha_nac,"
StrSql = StrSql & " est_civil , cuil, estado, reg_prev,"
StrSql = StrSql & " lug_trab , basico, neto, msr,"
StrSql = StrSql & " asi_flia , dtos, bruto, "
StrSql = StrSql & " prodesc , descripcion, pliqdesc, pliqmes, "
StrSql = StrSql & " pliqanio , profecpago, pliqfecdep, pliqbco,ultima_pag_impr, auxchar1, auxchar2, auxchar3, "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,auxchar4) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & proNro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & empresa & "'"
StrSql = StrSql & "," & emprNro
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & fecpago & "'"
StrSql = StrSql & ",'" & fecalta & "'"
StrSql = StrSql & ",'" & fecbaja & "'"
StrSql = StrSql & ",'" & contrato & "'"
StrSql = StrSql & ",'" & categoria & "'"
StrSql = StrSql & ",'" & Mid(direccion, 1, 50) & "'"
StrSql = StrSql & ",'" & puesto & "'"
StrSql = StrSql & ",'" & documento & "'"
StrSql = StrSql & ",'" & fecha_nac & "'"
StrSql = StrSql & ",'" & est_civil & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & "," & estado
StrSql = StrSql & ",'" & reg_prev & "'"
StrSql = StrSql & ",'" & lug_trab & "'"
StrSql = StrSql & "," & numberForSQL(basico)
StrSql = StrSql & "," & numberForSQL(neto)
StrSql = StrSql & "," & numberForSQL(msr)
StrSql = StrSql & "," & numberForSQL(asi_flia)
StrSql = StrSql & "," & numberForSQL(dtos)
StrSql = StrSql & "," & numberForSQL(bruto)
StrSql = StrSql & ",'" & prodesc & "'"
StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & "," & ConvFecha(profecpago)
StrSql = StrSql & "," & pliqfecdep
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & "," & Pagina
StrSql = StrSql & ",'" & emprCuit & "'"
StrSql = StrSql & ",'" & emprDire & "'"
StrSql = StrSql & ",'" & emprActiv & "'"
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & centroCosto & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Busco el detalle de la liquidacion
'------------------------------------------------------------------

StrSql = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, detliq.dlicant, detliq.dlimonto " & _
    " FROM cabliq " & _
    " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND proceso.pronro = " & proNro & _
    " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro " & _
    " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & _
    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 " & _
    " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "

'---LOG---
Flog.Writeline "Buscando datos del detalle de liquidacion"

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF

  StrSql = " INSERT INTO rep_libroley_det "
  StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
  StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
  StrSql = StrSql & " dlimonto) "
  StrSql = StrSql & " VALUES "
  StrSql = StrSql & "(" & NroProceso
  StrSql = StrSql & "," & ternro
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

'------------------------------------------------------------------
'Busco los datos de los familiares
'------------------------------------------------------------------

If (tipoFamiliares = "1") Or (tipoFamiliares = "2") Then

    StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc,tercero.ternro,tercero.terape, tercero.ternom, tercero.terfecnac, tercero.tersex, famest,famtrab,faminc,famDGIdesde " & _
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
    
    Do Until rsConsult.EOF
    
      GuardarFam = True
    
      If (tipoFamiliares = "2") Then
          GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg, cliqnro, concFamiliar01)
          If Not GuardarFam Then
             GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg2, cliqnro, concFamiliar02)
          End If
      End If
    
      If GuardarFam Then
      
          If IsNull(rsConsult!terfecnac) Then
             terfecnac = "NULL"
          Else
             terfecnac = ConvFecha(rsConsult!terfecnac)
          End If
        
          If IsNull(rsConsult!famfecvto) Then
             famfecvto = "NULL"
          Else
             famfecvto = ConvFecha(rsConsult!famfecvto)
          End If
        
          If IsNull(rsConsult!famDGIdesde) Then
             famDGIdesde = "NULL"
          Else
             famDGIdesde = ConvFecha(rsConsult!famDGIdesde)
          End If
          
          If IsNull(rsConsult!famDGIhasta) Then
             famDGIhasta = "NULL"
          Else
             famDGIhasta = ConvFecha(rsConsult!famDGIhasta)
          End If
          
        
          StrSql = " INSERT INTO rep_libroley_fam "
          StrSql = StrSql & " (bpronro, ternro , pronro, nrodoc,"
          StrSql = StrSql & " sigladoc, ternrofam, terape, ternom,"
          StrSql = StrSql & " terfecnac , tersex, famest, famtrab,"
          StrSql = StrSql & " faminc, famsalario, famfecvto, famCargaDGI,"
          StrSql = StrSql & " famDGIdesde , famDGIhasta, famemergencia, paredesc"
          StrSql = StrSql & " ) "
          StrSql = StrSql & " VALUES "
          StrSql = StrSql & "(" & NroProceso
          StrSql = StrSql & "," & ternro
          StrSql = StrSql & "," & proNro
          StrSql = StrSql & ",'" & rsConsult!nrodoc & "'"
          StrSql = StrSql & ",'" & rsConsult!sigladoc & "'"
          StrSql = StrSql & "," & rsConsult!ternro
          StrSql = StrSql & ",'" & rsConsult!terape & "'"
          StrSql = StrSql & ",'" & rsConsult!ternom & "'"
          StrSql = StrSql & "," & terfecnac
          StrSql = StrSql & "," & strForSQL(rsConsult!tersex)
          StrSql = StrSql & "," & strForSQL(rsConsult!famest)
          StrSql = StrSql & "," & strForSQL(rsConsult!famtrab)
          StrSql = StrSql & "," & strForSQL(rsConsult!faminc)
          StrSql = StrSql & "," & strForSQL(rsConsult!famsalario)
          StrSql = StrSql & "," & famfecvto
          StrSql = StrSql & "," & strForSQL(rsConsult!famCargaDGI)
          StrSql = StrSql & "," & famDGIdesde
          StrSql = StrSql & "," & famDGIhasta
          StrSql = StrSql & "," & strForSQL(rsConsult!famemergencia)
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


'--------------------------------------------------------------------
' Se encarga de generar los datos para Temaiken
'--------------------------------------------------------------------
Sub generarDatosEmpleado02(proNro, ternro, descripcion, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim pliqnro As Integer
Dim pliqmes
Dim pliqanio
Dim fecpago As String
Dim empfecalta
Dim fecalta As String
Dim fecbaja As String
Dim contrato As String
Dim categoria As String
Dim direccion As String
Dim puesto As String
Dim documento  As String
Dim fecha_nac As String
Dim est_civil As String
Dim Cuil As String
Dim estado As Integer
Dim reg_prev As String
Dim lug_trab  As String
Dim basico As String
Dim neto As String
Dim msr As String
Dim asi_flia As String
Dim dtos As String
Dim bruto As String
Dim profecpago
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim terfecnac
Dim famfecvto
Dim famDGIdesde
Dim famDGIhasta
Dim sexo
Dim Sueldo
Dim cliqnro
Dim acuSueldo
Dim GuardarFam As Boolean
Dim pliqdesde As Date
Dim pliqhasta As Date

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
   fecpago = rsConsult!profecpago
   profecpago = rsConsult!profecpago
   prodesc = rsConsult!prodesc
   pliqdesc = rsConsult!pliqdesc
   cliqnro = rsConsult!cliqnro
   pliqdesde = rsConsult!pliqdesde
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.Writeline "El empleado no esta en el proceso"
   Exit Sub
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2,empleado.empfaltagr,empleado.empremu,tercero.tersex "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro AND empleado.ternro= " & ternro
StrSql = StrSql & " WHERE empleado.ternro= " & ternro
       
'---LOG---
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
   empfecalta = rsConsult!empfaltagr
   Sueldo = rsConsult!empremu
   If CInt(rsConsult!tersex) = -1 Then
     sexo = "M"
   Else
     sexo = "F"
   End If
Else
   Flog.Writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close


'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & emprNro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!banco
Else
   pliqfecdep = ""
   pliqbco = ""
   Flog.Writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
' Obtengo los datos de las estructuras
'------------------------------------------------------------------

 sql = " SELECT empfaltagr, empleado.empest, empfecbaja,estr1.estrdabr AS categoria, estr2.estrdabr AS contrato, estr3.estrdabr AS puesto, estr4.estrdabr AS regprev, estr5.estrdabr AS sucursal "
 sql = sql & " FROM empleado "
 'Busco la categoria
 sql = sql & " LEFT JOIN his_estructura his1 ON his1.ternro = empleado.ternro AND his1.tenro = 3 AND empleado.ternro= " & ternro & " AND his1.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his1.htethasta IS NULL OR his1.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr1 ON estr1.estrnro = his1.estrnro "
 'Busco el contrato
 sql = sql & " LEFT JOIN his_estructura his2 ON his2.ternro = empleado.ternro AND his2.tenro = 18 AND empleado.ternro= " & ternro & " AND his2.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his2.htethasta IS NULL OR his2.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr2 ON estr2.estrnro = his2.estrnro "
 'Busco el puesto
 sql = sql & " LEFT JOIN his_estructura his3 ON his3.ternro = empleado.ternro AND his3.tenro = 4 AND empleado.ternro= " & ternro & " AND his3.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his3.htethasta IS NULL OR his3.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr3 ON estr3.estrnro = his3.estrnro "
 'Busco el regimen previsional
 sql = sql & " LEFT JOIN his_estructura his4 ON his4.ternro = empleado.ternro AND his4.tenro = 15 AND empleado.ternro= " & ternro & " AND his4.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his4.htethasta IS NULL OR his4.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr4 ON estr4.estrnro = his4.estrnro "
 'Busco la sucursal
 sql = sql & " LEFT JOIN his_estructura his5 ON his5.ternro = empleado.ternro AND his5.tenro = 1 AND empleado.ternro= " & ternro & " AND his5.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his5.htethasta IS NULL OR his5.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr5 ON estr5.estrnro = his5.estrnro "
    
 sql = sql & " WHERE empleado.ternro = " & ternro
 
'---LOG---
Flog.Writeline "Buscando datos de las estructuras"
 

 OpenRecordset sql, rsConsult

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

 'Contrato
 If IsNull(rsConsult!contrato) Then
   contrato = ""
 Else
   contrato = rsConsult!contrato
 End If
 
 'Categoria
 If IsNull(rsConsult!categoria) Then
   categoria = ""
 Else
   categoria = rsConsult!categoria
 End If
 
 'Puesto
 If IsNull(rsConsult!puesto) Then
   puesto = ""
 Else
   puesto = rsConsult!puesto
 End If
 
 'Reg. Prev.
 If IsNull(rsConsult!regprev) Then
   reg_prev = ""
 Else
   reg_prev = rsConsult!regprev
 End If
 
 'Sucursal
 If IsNull(rsConsult!sucursal) Then
   lug_trab = ""
 Else
   lug_trab = rsConsult!sucursal
 End If

 estado = rsConsult!empest
 
 rsConsult.Close
 
 
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos del estructura 1"

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
Flog.Writeline "Buscando datos del estructura 2"

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
Flog.Writeline "Buscando datos del estructura 3"

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
'Busco el valor de la direccion
'------------------------------------------------------------------


sql = "SELECT calle,nro,cabdom.domdefault,locdesc "
sql = sql & " FROM detdom INNER JOIN cabdom ON detdom.domnro=cabdom.domnro "
sql = sql & " INNER JOIN tipodomi ON cabdom.tidonro=tipodomi.tidonro AND tipodomi.tidonro=2 "
sql = sql & " LEFT JOIN localidad on localidad.locnro = detdom.locnro"
sql = sql & " WHERE cabdom.domdefault = -1 AND cabdom.ternro=" & ternro
    
'---LOG---
Flog.Writeline "Buscando datos de la direccion"
    
OpenRecordset sql, rsConsult

direccion = ""
If Not rsConsult.EOF Then
       direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
End If

rsConsult.Close
    
'------------------------------------------------------------------
'Busco el valor del fecha nac, cuil, doc y tipo doc
'------------------------------------------------------------------

sql = " SELECT tercero.terfecnac, cuil.nrodoc AS nrocuil, docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc, estcivil.estcivdesabr"
sql = sql & " FROM tercero "
sql = sql & " LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
sql = sql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
sql = sql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
sql = sql & " LEFT JOIN estcivil ON estcivil.estcivnro= tercero.estcivnro "
sql = sql & " WHERE tercero.ternro= " & ternro

'---LOG---
Flog.Writeline "Buscando datos de los documentos"

OpenRecordset sql, rsConsult

Cuil = ""
documento = ""
fecha_nac = ""
est_civil = 0

If Not rsConsult.EOF Then
   Cuil = rsConsult!nrocuil
   documento = rsConsult!sigladoc & "-" & rsConsult!nrodoc
   fecha_nac = rsConsult!terfecnac
   est_civil = rsConsult!estcivdesabr
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------

'Basico
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_basico
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del basico"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   basico = 0
Else
   basico = rsConsult!almonto
End If

rsConsult.Close

'Neto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_neto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del neto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   neto = 0
Else
   neto = rsConsult!almonto
End If

rsConsult.Close

'MSR
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_msr
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del MSR"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   msr = 0
Else
   msr = rsConsult!almonto
End If

rsConsult.Close

'Asig. Familia
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_asi_flia
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de la asig. fam."

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   asi_flia = 0
Else
   asi_flia = rsConsult!almonto
End If

rsConsult.Close

'Descuentos
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_Dtos
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de descuentos"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   dtos = 0
Else
   dtos = rsConsult!almonto
End If

rsConsult.Close

'Bruto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_bruto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del bruto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   bruto = 0
Else
   bruto = rsConsult!almonto
End If

rsConsult.Close

' -------------------------------------------------------------------------
' Busco los datos del sueldo
'--------------------------------------------------------------------------

If IsNull(Sueldo) Then
    sql = " SELECT * FROM confrep "
    sql = sql & " WHERE repnro = 61 "
    sql = sql & " AND confnrocol = 9 "
    
    OpenRecordset sql, rsConsult

    If rsConsult.EOF Then
       acuSueldo = 0
    Else
       acuSueldo = rsConsult!confval
    End If

    rsConsult.Close
       
    'Sueldo
    sql = " SELECT almonto,acunro "
    sql = sql & " FROM acu_liq"
    sql = sql & " WHERE acunro = " & acuSueldo
    sql = sql & " AND cliqnro =  " & cliqnro
    
    '---LOG---
    Flog.Writeline "Buscando datos del sueldo"
    
    OpenRecordset sql, rsConsult
    
    If rsConsult.EOF Then
       Sueldo = 0
    Else
       Sueldo = rsConsult!almonto
    End If
    
    rsConsult.Close
   
End If

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

If IsNull(pliqfecdep) Then
   pliqfecdep = "NULL"
Else
   pliqfecdep = ConvFecha(pliqfecdep)
End If

StrSql = " INSERT INTO rep_libroley "
StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
StrSql = StrSql & " empresa , emprnro, pliqnro, fecpago,"
StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
StrSql = StrSql & " direccion , puesto, documento, fecha_nac,"
StrSql = StrSql & " est_civil , cuil, estado, reg_prev,"
StrSql = StrSql & " lug_trab , basico, neto, msr,"
StrSql = StrSql & " asi_flia , dtos, bruto, "
StrSql = StrSql & " prodesc , descripcion, pliqdesc, pliqmes, "
StrSql = StrSql & " pliqanio , profecpago, pliqfecdep, pliqbco,ultima_pag_impr,auxdeci1,auxchar1, "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,tipofam) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & proNro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & empresa & "'"
StrSql = StrSql & "," & emprNro
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & fecpago & "'"
StrSql = StrSql & ",'" & fecalta & "'"
StrSql = StrSql & ",'" & fecbaja & "'"
StrSql = StrSql & ",'" & contrato & "'"
StrSql = StrSql & ",'" & categoria & "'"
StrSql = StrSql & ",'" & direccion & "'"
StrSql = StrSql & ",'" & puesto & "'"
StrSql = StrSql & ",'" & documento & "'"
StrSql = StrSql & ",'" & fecha_nac & "'"
StrSql = StrSql & ",'" & est_civil & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & "," & estado
StrSql = StrSql & ",'" & reg_prev & "'"
StrSql = StrSql & ",'" & lug_trab & "'"
StrSql = StrSql & "," & numberForSQL(basico)
StrSql = StrSql & "," & numberForSQL(neto)
StrSql = StrSql & "," & numberForSQL(msr)
StrSql = StrSql & "," & numberForSQL(asi_flia)
StrSql = StrSql & "," & numberForSQL(dtos)
StrSql = StrSql & "," & numberForSQL(bruto)
StrSql = StrSql & ",'" & prodesc & "'"
StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & "," & ConvFecha(profecpago)
StrSql = StrSql & "," & pliqfecdep
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & "," & Pagina
StrSql = StrSql & "," & numberForSQL(Sueldo)
StrSql = StrSql & ",'" & sexo & "'"
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & "," & orden
StrSql = StrSql & "," & tipoFamiliares
StrSql = StrSql & ")"


'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Busco el detalle de la liquidacion
'------------------------------------------------------------------

StrSql = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, detliq.dlicant, detliq.dlimonto " & _
    " FROM cabliq " & _
    " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND proceso.pronro = " & proNro & _
    " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro " & _
    " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & _
    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 " & _
    " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "

'---LOG---
Flog.Writeline "Buscando datos del detalle de liquidacion"

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF

  StrSql = " INSERT INTO rep_libroley_det "
  StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
  StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
  StrSql = StrSql & " dlimonto) "
  StrSql = StrSql & " VALUES "
  StrSql = StrSql & "(" & NroProceso
  StrSql = StrSql & "," & ternro
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

'------------------------------------------------------------------
'Busco los datos de los familiares
'------------------------------------------------------------------

If (tipoFamiliares = "1") Or (tipoFamiliares = "2") Then

    StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc,tercero.ternro,tercero.terape, tercero.ternom, tercero.terfecnac, tercero.tersex, famest,famtrab,faminc,famDGIdesde " & _
          " famsalario, famfecvto, famCargaDGI, famDGIdesde, famDGIhasta, famemergencia, " & _
          " paredesc " & _
          " FROM  tercero INNER JOIN familiar ON tercero.ternro=familiar.ternro " & _
          " LEFT JOIN parentesco ON familiar.parenro=parentesco.parenro " & _
          " LEFT JOIN ter_doc docu ON (docu.ternro= familiar.ternro and docu.tidnro>0 and docu.tidnro<5) " & _
          " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro " & _
          " WHERE familiar.famest = -1 AND familiar.parenro=2 AND familiar.empleado = " & ternro
          
    '---LOG---
    Flog.Writeline "Buscando datos de los familiares"
    
    OpenRecordset StrSql, rsConsult
    
    Do Until rsConsult.EOF
    
      GuardarFam = True
    
      If (tipoFamiliares = "2") Then
          GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg, cliqnro, concFamiliar01)
          If Not GuardarFam Then
             GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg2, cliqnro, concFamiliar02)
          End If
          
      End If
    
      If GuardarFam Then
    
          If IsNull(rsConsult!terfecnac) Then
             terfecnac = "NULL"
          Else
             terfecnac = ConvFecha(rsConsult!terfecnac)
          End If
        
          If IsNull(rsConsult!famfecvto) Then
             famfecvto = "NULL"
          Else
             famfecvto = ConvFecha(rsConsult!famfecvto)
          End If
        
          If IsNull(rsConsult!famDGIdesde) Then
             famDGIdesde = "NULL"
          Else
             famDGIdesde = ConvFecha(rsConsult!famDGIdesde)
          End If
          
          If IsNull(rsConsult!famDGIhasta) Then
             famDGIhasta = "NULL"
          Else
             famDGIhasta = ConvFecha(rsConsult!famDGIhasta)
          End If
          
        
          StrSql = " INSERT INTO rep_libroley_fam "
          StrSql = StrSql & " (bpronro, ternro , pronro, nrodoc,"
          StrSql = StrSql & " sigladoc, ternrofam, terape, ternom,"
          StrSql = StrSql & " terfecnac , tersex, famest, famtrab,"
          StrSql = StrSql & " faminc, famsalario, famfecvto, famCargaDGI,"
          StrSql = StrSql & " famDGIdesde , famDGIhasta, famemergencia, paredesc"
          StrSql = StrSql & " ) "
          StrSql = StrSql & " VALUES "
          StrSql = StrSql & "(" & NroProceso
          StrSql = StrSql & "," & ternro
          StrSql = StrSql & "," & proNro
          StrSql = StrSql & ",'" & rsConsult!nrodoc & "'"
          StrSql = StrSql & ",'" & rsConsult!sigladoc & "'"
          StrSql = StrSql & "," & rsConsult!ternro
          StrSql = StrSql & ",'" & rsConsult!terape & "'"
          StrSql = StrSql & ",'" & rsConsult!ternom & "'"
          StrSql = StrSql & "," & terfecnac
          StrSql = StrSql & "," & strForSQL(rsConsult!tersex)
          StrSql = StrSql & "," & strForSQL(rsConsult!famest)
          StrSql = StrSql & "," & strForSQL(rsConsult!famtrab)
          StrSql = StrSql & "," & strForSQL(rsConsult!faminc)
          StrSql = StrSql & "," & strForSQL(rsConsult!famsalario)
          StrSql = StrSql & "," & famfecvto
          StrSql = StrSql & "," & strForSQL(rsConsult!famCargaDGI)
          StrSql = StrSql & "," & famDGIdesde
          StrSql = StrSql & "," & famDGIhasta
          StrSql = StrSql & "," & strForSQL(rsConsult!famemergencia)
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


'--------------------------------------------------------------------
' Se encarga de generar los datos para Accor
'--------------------------------------------------------------------
Sub generarDatosEmpleado03(proNro, ternro, descripcion, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim pliqnro As Integer
Dim pliqmes
Dim pliqanio
Dim fecpago As String
Dim empfecalta
Dim fecalta As String
Dim fecbaja As String
Dim contrato As String
Dim categoria As String
Dim direccion As String
Dim puesto As String
Dim documento  As String
Dim fecha_nac As String
Dim est_civil As String
Dim Cuil As String
Dim estado As Integer
Dim reg_prev As String
Dim lug_trab  As String
Dim basico As String
Dim neto As String
Dim msr As String
Dim asi_flia As String
Dim dtos As String
Dim bruto As String
Dim profecpago
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim terfecnac
Dim famfecvto
Dim famDGIdesde
Dim famDGIhasta
Dim sexo
Dim Sueldo
Dim nacionalidad
Dim causaEgreso
Dim cliqnro
Dim GuardarFam As Boolean
Dim pliqdesde As Date
Dim pliqhasta As Date

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
   fecpago = rsConsult!profecpago
   profecpago = rsConsult!profecpago
   prodesc = rsConsult!prodesc
   pliqdesc = rsConsult!pliqdesc
   cliqnro = rsConsult!cliqnro
   pliqdesde = rsConsult!pliqdesde
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.Writeline "El empleado no se encuentra en el proceso."
   Exit Sub
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2,empleado.empfaltagr,empleado.empremu,tercero.tersex "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro AND empleado.ternro= " & ternro
StrSql = StrSql & " WHERE empleado.ternro= " & ternro
       
'---LOG---
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
   empfecalta = rsConsult!empfaltagr
   Sueldo = rsConsult!empremu
   If CInt(rsConsult!tersex) = -1 Then
     sexo = "M"
   Else
     sexo = "F"
   End If
Else
   Flog.Writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close


'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & emprNro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!banco
Else
   pliqfecdep = ""
   pliqbco = ""
   Flog.Writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
' Obtengo los datos de las estructuras
'------------------------------------------------------------------

 sql = " SELECT empfaltagr, empleado.empest, empfecbaja,estr1.estrdabr AS categoria, estr2.estrdabr AS contrato, estr3.estrdabr AS puesto, estr4.estrdabr AS regprev, estr5.estrdabr AS sucursal "
 sql = sql & " FROM empleado "
 'Busco la categoria
 sql = sql & " LEFT JOIN his_estructura his1 ON his1.ternro = empleado.ternro AND his1.tenro = 3 AND empleado.ternro= " & ternro & " AND his1.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his1.htethasta IS NULL OR his1.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr1 ON estr1.estrnro = his1.estrnro "
 'Busco el contrato
 sql = sql & " LEFT JOIN his_estructura his2 ON his2.ternro = empleado.ternro AND his2.tenro = 18 AND empleado.ternro= " & ternro & " AND his2.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his2.htethasta IS NULL OR his2.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr2 ON estr2.estrnro = his2.estrnro "
 'Busco el puesto
 sql = sql & " LEFT JOIN his_estructura his3 ON his3.ternro = empleado.ternro AND his3.tenro = 4 AND empleado.ternro= " & ternro & " AND his3.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his3.htethasta IS NULL OR his3.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr3 ON estr3.estrnro = his3.estrnro "
 'Busco el regimen previsional
 sql = sql & " LEFT JOIN his_estructura his4 ON his4.ternro = empleado.ternro AND his4.tenro = 15 AND empleado.ternro= " & ternro & " AND his4.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his4.htethasta IS NULL OR his4.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr4 ON estr4.estrnro = his4.estrnro "
 'Busco la sucursal
 sql = sql & " LEFT JOIN his_estructura his5 ON his5.ternro = empleado.ternro AND his5.tenro = 1 AND empleado.ternro= " & ternro & " AND his5.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his5.htethasta IS NULL OR his5.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr5 ON estr5.estrnro = his5.estrnro "
    
 sql = sql & " WHERE empleado.ternro = " & ternro
 
'---LOG---
Flog.Writeline "Buscando datos de las estructuras"

 OpenRecordset sql, rsConsult

 'Fecha Alta
 If IsNull(rsConsult!empfaltagr) Then
    fecalta = ""
 Else
    fecalta = rsConsult!empfaltagr
 End If
 
 'Fecha Baja
' If IsNull(rsConsult!empfecbaja) Then
'    fecbaja = ""
' Else
'    fecbaja = rsConsult!empfecbaja
' End If

 'Contrato
 If IsNull(rsConsult!contrato) Then
   contrato = ""
 Else
   contrato = rsConsult!contrato
 End If
 
 'Categoria
 If IsNull(rsConsult!categoria) Then
   categoria = ""
 Else
   categoria = rsConsult!categoria
 End If
 
 'Puesto
 If IsNull(rsConsult!puesto) Then
   puesto = ""
 Else
   puesto = rsConsult!puesto
 End If
 
 'Reg. Prev.
 If IsNull(rsConsult!regprev) Then
   reg_prev = ""
 Else
   reg_prev = rsConsult!regprev
 End If
 
 'Sucursal
 If IsNull(rsConsult!sucursal) Then
   lug_trab = ""
 Else
   lug_trab = rsConsult!sucursal
 End If

 estado = rsConsult!empest
 
 rsConsult.Close
 
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos del estructura 1"

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
Flog.Writeline "Buscando datos del estructura 2"

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
Flog.Writeline "Buscando datos del estructura 3"

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
'Busco el valor de la direccion
'------------------------------------------------------------------

sql = " SELECT detdom.calle,detdom.nro,localidad.locdesc"
sql = sql & " From his_estructura"
sql = sql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND tenro=1 AND his_estructura.ternro=" & ternro
sql = sql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
sql = sql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
sql = sql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
    
'---LOG---
Flog.Writeline "Buscando datos de la direccion"
    
OpenRecordset sql, rsConsult

direccion = ""
If Not rsConsult.EOF Then
       direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
End If

rsConsult.Close
    
'------------------------------------------------------------------
'Busco el valor del fecha nac, cuil, doc y tipo doc
'------------------------------------------------------------------

sql = " SELECT tercero.terfecnac, cuil.nrodoc AS nrocuil, docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc, estcivil.estcivdesabr"
sql = sql & " FROM tercero "
sql = sql & " LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
sql = sql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
sql = sql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
sql = sql & " LEFT JOIN estcivil ON estcivil.estcivnro= tercero.estcivnro "
sql = sql & " WHERE tercero.ternro= " & ternro

'---LOG---
Flog.Writeline "Buscando datos del documento"

OpenRecordset sql, rsConsult

Cuil = ""
documento = ""
fecha_nac = ""
est_civil = 0

If Not rsConsult.EOF Then
   Cuil = rsConsult!nrocuil
   documento = rsConsult!sigladoc & "-" & rsConsult!nrodoc
   fecha_nac = rsConsult!terfecnac
   est_civil = rsConsult!estcivdesabr
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------

'Basico
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_basico
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del basico"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   basico = 0
Else
   basico = rsConsult!almonto
End If

rsConsult.Close

'Neto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_neto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del neto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   neto = 0
Else
   neto = rsConsult!almonto
End If

rsConsult.Close

'MSR
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_msr
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del MSR"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   msr = 0
Else
   msr = rsConsult!almonto
End If

rsConsult.Close

'Asig. Familia
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_asi_flia
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de la asig. fam."

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   asi_flia = 0
Else
   asi_flia = rsConsult!almonto
End If

rsConsult.Close

'Descuentos
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_Dtos
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos descuentos"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   dtos = 0
Else
   dtos = rsConsult!almonto
End If

rsConsult.Close

'Bruto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_bruto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del bruto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   bruto = 0
Else
   bruto = rsConsult!almonto
End If

rsConsult.Close

If IsNull(Sueldo) Then
   Sueldo = bruto
End If

' -------------------------------------------------------------------------
' Busco la nacionalidad
'--------------------------------------------------------------------------

StrSql = "SELECT nacionaldes " & _
    " FROM tercero " & _
    " INNER JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro AND tercero.ternro = " & ternro

'---LOG---
Flog.Writeline "Buscando datos de la nacionalidad"

OpenRecordset StrSql, rsConsult

nacionalidad = ""

If rsConsult.EOF Then
    Flog.Writeline "No se encontró la nacionalidad"
'    Exit Sub
Else
    nacionalidad = rsConsult!nacionaldes
End If

rsConsult.Close

' -------------------------------------------------------------------------
' Busco la fecha de baja y la causa de egreso
'--------------------------------------------------------------------------

causaEgreso = ""
fecbaja = ""

'---LOG---
Flog.Writeline "Buscando datos de la causa de egreso"

 StrSql = " SELECT altfec, bajfec, caudes, estado, empatareas "
 StrSql = StrSql & " FROM fases "
 StrSql = StrSql & " LEFT JOIN empant ON fases.empantnro=empant.empantnro "
 StrSql = StrSql & " LEFT JOIN causa ON fases.caunro=causa.caunro "
 StrSql = StrSql & " WHERE fases.empleado=" & ternro & " order by altfec DESC"

 OpenRecordset StrSql, rsConsult

 If rsConsult.EOF Then
     Flog.Writeline "No se encontró la causa de egreso"
 '    Exit Sub
 Else
    If Not CBool(rsConsult!estado) Then
       fecbaja = rsConsult!bajfec
       causaEgreso = rsConsult!caudes
    End If
 End If
 
 rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

If IsNull(pliqfecdep) Or (pliqfecdep = "") Then
   pliqfecdep = "NULL"
Else
   pliqfecdep = ConvFecha(pliqfecdep)
End If

StrSql = " INSERT INTO rep_libroley "
StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
StrSql = StrSql & " empresa , emprnro, pliqnro, fecpago,"
StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
StrSql = StrSql & " direccion , puesto, documento, fecha_nac,"
StrSql = StrSql & " est_civil , cuil, estado, reg_prev,"
StrSql = StrSql & " lug_trab , basico, neto, msr,"
StrSql = StrSql & " asi_flia , dtos, bruto, "
StrSql = StrSql & " prodesc , descripcion, pliqdesc, pliqmes, "
StrSql = StrSql & " pliqanio , profecpago, pliqfecdep, pliqbco,ultima_pag_impr,auxchar1,auxchar2,auxchar3,auxchar4, "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,auxchar5) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & proNro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & empresa & "'"
StrSql = StrSql & "," & emprNro
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & fecpago & "'"
StrSql = StrSql & ",'" & fecalta & "'"
StrSql = StrSql & ",'" & fecbaja & "'"
StrSql = StrSql & ",'" & contrato & "'"
StrSql = StrSql & ",'" & categoria & "'"
StrSql = StrSql & ",'" & direccion & "'"
StrSql = StrSql & ",'" & puesto & "'"
StrSql = StrSql & ",'" & documento & "'"
StrSql = StrSql & ",'" & fecha_nac & "'"
StrSql = StrSql & ",'" & est_civil & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & "," & estado
StrSql = StrSql & ",'" & reg_prev & "'"
StrSql = StrSql & ",'" & lug_trab & "'"
StrSql = StrSql & "," & numberForSQL(basico)
StrSql = StrSql & "," & numberForSQL(neto)
StrSql = StrSql & "," & numberForSQL(msr)
StrSql = StrSql & "," & numberForSQL(asi_flia)
StrSql = StrSql & "," & numberForSQL(dtos)
StrSql = StrSql & "," & numberForSQL(bruto)
StrSql = StrSql & ",'" & prodesc & "'"
StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & "," & ConvFecha(profecpago)
StrSql = StrSql & "," & pliqfecdep
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & "," & Pagina
StrSql = StrSql & ",'" & nacionalidad & "'"
StrSql = StrSql & ",'" & causaEgreso & "'"
StrSql = StrSql & ",'" & emprCuit & "'"
StrSql = StrSql & ",'" & emprDire & "'"
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & "," & orden
StrSql = StrSql & "," & controlNull(emprActiv)
StrSql = StrSql & ")"


'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Busco el detalle de la liquidacion
'------------------------------------------------------------------

StrSql = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, detliq.dlicant, detliq.dlimonto " & _
    " FROM cabliq " & _
    " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND proceso.pronro = " & proNro & _
    " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro " & _
    " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & _
    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 " & _
    " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "

'---LOG---
Flog.Writeline "Buscando datos del detalle de liquidacion"

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF

  StrSql = " INSERT INTO rep_libroley_det "
  StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
  StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
  StrSql = StrSql & " dlimonto) "
  StrSql = StrSql & " VALUES "
  StrSql = StrSql & "(" & NroProceso
  StrSql = StrSql & "," & ternro
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

'------------------------------------------------------------------
'Busco los datos de los familiares
'------------------------------------------------------------------

If (tipoFamiliares = "1") Or (tipoFamiliares = "2") Then

    StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc,tercero.ternro,tercero.terape, tercero.ternom, tercero.terfecnac, tercero.tersex, famest,famtrab,faminc,famDGIdesde " & _
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
    
    Do Until rsConsult.EOF
    
      GuardarFam = True
    
      If (tipoFamiliares = "2") Then
          GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg, cliqnro, concFamiliar01)
          If Not GuardarFam Then
             GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg2, cliqnro, concFamiliar02)
          End If
      End If
    
      If GuardarFam Then
   
          If IsNull(rsConsult!terfecnac) Then
             terfecnac = "NULL"
          Else
             terfecnac = ConvFecha(rsConsult!terfecnac)
          End If
        
          If IsNull(rsConsult!famfecvto) Then
             famfecvto = "NULL"
          Else
             famfecvto = ConvFecha(rsConsult!famfecvto)
          End If
        
          If IsNull(rsConsult!famDGIdesde) Then
             famDGIdesde = "NULL"
          Else
             famDGIdesde = ConvFecha(rsConsult!famDGIdesde)
          End If
          
          If IsNull(rsConsult!famDGIhasta) Then
             famDGIhasta = "NULL"
          Else
             famDGIhasta = ConvFecha(rsConsult!famDGIhasta)
          End If
          
        
          StrSql = " INSERT INTO rep_libroley_fam "
          StrSql = StrSql & " (bpronro, ternro , pronro, nrodoc,"
          StrSql = StrSql & " sigladoc, ternrofam, terape, ternom,"
          StrSql = StrSql & " terfecnac , tersex, famest, famtrab,"
          StrSql = StrSql & " faminc, famsalario, famfecvto, famCargaDGI,"
          StrSql = StrSql & " famDGIdesde , famDGIhasta, famemergencia, paredesc"
          StrSql = StrSql & " ) "
          StrSql = StrSql & " VALUES "
          StrSql = StrSql & "(" & NroProceso
          StrSql = StrSql & "," & ternro
          StrSql = StrSql & "," & proNro
          StrSql = StrSql & ",'" & rsConsult!nrodoc & "'"
          StrSql = StrSql & ",'" & rsConsult!sigladoc & "'"
          StrSql = StrSql & "," & rsConsult!ternro
          StrSql = StrSql & ",'" & rsConsult!terape & "'"
          StrSql = StrSql & ",'" & rsConsult!ternom & "'"
          StrSql = StrSql & "," & terfecnac
          StrSql = StrSql & "," & strForSQL(rsConsult!tersex)
          StrSql = StrSql & "," & strForSQL(rsConsult!famest)
          StrSql = StrSql & "," & strForSQL(rsConsult!famtrab)
          StrSql = StrSql & "," & strForSQL(rsConsult!faminc)
          StrSql = StrSql & "," & strForSQL(rsConsult!famsalario)
          StrSql = StrSql & "," & famfecvto
          StrSql = StrSql & "," & strForSQL(rsConsult!famCargaDGI)
          StrSql = StrSql & "," & famDGIdesde
          StrSql = StrSql & "," & famDGIhasta
          StrSql = StrSql & "," & strForSQL(rsConsult!famemergencia)
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

'--------------------------------------------------------------------
' Se encarga de generar los datos para Citrusvil
'--------------------------------------------------------------------
Sub generarDatosEmpleado04(proNro, ternro, descripcion, orden)

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
Dim direccion As String
Dim puesto As String
Dim documento  As String
Dim fecha_nac As String
Dim est_civil As String
Dim Cuil As String
Dim estado
Dim reg_prev As String
Dim lug_trab  As String
Dim basico As String
Dim neto As String
Dim msr As String
Dim asi_flia As String
Dim dtos As String
Dim bruto As String
Dim profecpago
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim terfecnac
Dim famfecvto
Dim famDGIdesde
Dim famDGIhasta
Dim cliqnro
Dim GuardarFam As Boolean
Dim pliqdesde As Date
Dim pliqhasta As Date
Dim nroInscripcion

Dim centroCosto

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
   fecpago = rsConsult!profecpago
   profecpago = rsConsult!profecpago
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
   empfecalta = rsConsult!empfaltagr
Else
   Flog.Writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & emprNro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!banco
Else
   pliqfecdep = ""
   pliqbco = ""
   Flog.Writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
' Obtengo los datos de las estructuras
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de las estructuras"

 sql = " SELECT empfaltagr, empleado.empest, empfecbaja,estr1.estrdabr AS categoria, estr2.estrdabr AS contrato, estr3.estrdabr AS puesto, estr4.estrdabr AS regprev, estr5.estrdabr AS sucursal "
 sql = sql & " FROM empleado "
 'Busco la categoria
 sql = sql & " LEFT JOIN his_estructura his1 ON his1.ternro = empleado.ternro AND his1.tenro = 3 AND empleado.ternro= " & ternro & " AND his1.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his1.htethasta IS NULL OR his1.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr1 ON estr1.estrnro = his1.estrnro "
 'Busco el contrato
 sql = sql & " LEFT JOIN his_estructura his2 ON his2.ternro = empleado.ternro AND his2.tenro = 18 AND empleado.ternro= " & ternro & " AND his2.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his2.htethasta IS NULL OR his2.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr2 ON estr2.estrnro = his2.estrnro "
 'Busco el puesto
 sql = sql & " LEFT JOIN his_estructura his3 ON his3.ternro = empleado.ternro AND his3.tenro = 4 AND empleado.ternro= " & ternro & " AND his3.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his3.htethasta IS NULL OR his3.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr3 ON estr3.estrnro = his3.estrnro "
 'Busco el regimen previsional
 sql = sql & " LEFT JOIN his_estructura his4 ON his4.ternro = empleado.ternro AND his4.tenro = 15 AND empleado.ternro= " & ternro & " AND his4.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his4.htethasta IS NULL OR his4.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr4 ON estr4.estrnro = his4.estrnro "
 'Busco la sucursal
 sql = sql & " LEFT JOIN his_estructura his5 ON his5.ternro = empleado.ternro AND his5.tenro = 1 AND empleado.ternro= " & ternro & " AND his5.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his5.htethasta IS NULL OR his5.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr5 ON estr5.estrnro = his5.estrnro "
    
 sql = sql & " WHERE empleado.ternro = " & ternro

 OpenRecordset sql, rsConsult

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

 'Contrato
 If IsNull(rsConsult!contrato) Then
   contrato = ""
 Else
   contrato = rsConsult!contrato
 End If
 
 'Categoria
 If IsNull(rsConsult!categoria) Then
   categoria = ""
 Else
   categoria = rsConsult!categoria
 End If
 
 'Puesto
 If IsNull(rsConsult!puesto) Then
   puesto = ""
 Else
   puesto = rsConsult!puesto
 End If
 
 'Reg. Prev.
 If IsNull(rsConsult!regprev) Then
   reg_prev = ""
 Else
   reg_prev = rsConsult!regprev
 End If
 
 'Sucursal
 If IsNull(rsConsult!sucursal) Then
   lug_trab = ""
 Else
   lug_trab = rsConsult!sucursal
 End If

 estado = rsConsult!empest
 
 rsConsult.Close
 
'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos del centro de costo"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult

centroCosto = ""

If Not rsConsult.EOF Then
   centroCosto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
'   GoTo MError
End If

rsConsult.Close
   
'------------------------------------------------------------------
'Busco el valor de la condicion
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de la condicion"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=36 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult

contrato = ""

If Not rsConsult.EOF Then
   contrato = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la condicion"
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
'Busco el valor de la direccion
'------------------------------------------------------------------

sql = " SELECT detdom.calle,detdom.nro,localidad.locdesc"
sql = sql & " From his_estructura"
sql = sql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND tenro=1 AND his_estructura.ternro=" & ternro
sql = sql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
sql = sql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
sql = sql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
    
'---LOG---
Flog.Writeline "Buscando datos de la direccion"

OpenRecordset sql, rsConsult

direccion = ""
If Not rsConsult.EOF Then
       direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
End If

rsConsult.Close
    
'------------------------------------------------------------------
'Busco el valor del fecha nac, cuil, doc y tipo doc
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de los documentos"

sql = " SELECT tercero.terfecnac, cuil.nrodoc AS nrocuil, docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc, estcivil.estcivdesabr"
sql = sql & " FROM tercero "
sql = sql & " LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
sql = sql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
sql = sql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
sql = sql & " LEFT JOIN estcivil ON estcivil.estcivnro= tercero.estcivnro "
sql = sql & " WHERE tercero.ternro= " & ternro

OpenRecordset sql, rsConsult

Cuil = ""
documento = ""
fecha_nac = ""
est_civil = 0

If Not rsConsult.EOF Then
   Cuil = rsConsult!nrocuil
   documento = rsConsult!sigladoc & "-" & rsConsult!nrodoc
   fecha_nac = rsConsult!terfecnac
   est_civil = rsConsult!estcivdesabr
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------

'Basico
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_basico
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del basico"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   basico = 0
Else
   basico = rsConsult!almonto
End If

rsConsult.Close

'Neto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_neto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del neto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   neto = 0
Else
   neto = rsConsult!almonto
End If

rsConsult.Close

'MSR
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_msr
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del msr"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   msr = 0
Else
   msr = rsConsult!almonto
End If

rsConsult.Close

'Asig. Familia
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_asi_flia
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de la asig. familia"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   asi_flia = 0
Else
   asi_flia = rsConsult!almonto
End If

rsConsult.Close

'Descuentos
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_Dtos
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de los descuentos"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   dtos = 0
Else
   dtos = rsConsult!almonto
End If

rsConsult.Close

'Bruto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_bruto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del bruto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   bruto = 0
Else
   bruto = rsConsult!almonto
End If

rsConsult.Close

'Consulta para obtener el nro. de inscripcion de la empresa
StrSql = "SELECT nrodoc FROM tercero " & _
         " INNER JOIN ter_doc ON (tercero.ternro = ter_doc.ternro and ter_doc.tidnro = " & TipDocNroInscr & ")" & _
         " Where tercero.ternro =" & emprTer

'---LOG---
Flog.Writeline "Buscando datos del nro. de inscripcion de la empresa"

OpenRecordset StrSql, rsConsult

If rsConsult.EOF Then
    Flog.Writeline "No se encontró el nro. de inscripcion de la Empresa"
    'Exit Sub
    nroInscripcion = "  "
Else
    nroInscripcion = rsConsult!nrodoc
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

If IsNull(pliqfecdep) Then
   pliqfecdep = "NULL"
Else
   pliqfecdep = ConvFecha(pliqfecdep)
End If

StrSql = " INSERT INTO rep_libroley "
StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
StrSql = StrSql & " empresa , emprnro, pliqnro, fecpago,"
StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
StrSql = StrSql & " direccion , puesto, documento, fecha_nac,"
StrSql = StrSql & " est_civil , cuil, estado, reg_prev,"
StrSql = StrSql & " lug_trab , basico, neto, msr,"
StrSql = StrSql & " asi_flia , dtos, bruto, "
StrSql = StrSql & " prodesc , descripcion, pliqdesc, pliqmes, "
StrSql = StrSql & " pliqanio , profecpago, pliqfecdep, pliqbco,ultima_pag_impr, auxchar1, auxchar2, auxchar3, "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,auxchar4,auxchar5) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & proNro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & empresa & "'"
StrSql = StrSql & "," & emprNro
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & fecpago & "'"
StrSql = StrSql & ",'" & fecalta & "'"
StrSql = StrSql & ",'" & fecbaja & "'"
StrSql = StrSql & ",'" & contrato & "'"
StrSql = StrSql & ",'" & categoria & "'"
StrSql = StrSql & ",'" & direccion & "'"
StrSql = StrSql & ",'" & puesto & "'"
StrSql = StrSql & ",'" & documento & "'"
StrSql = StrSql & ",'" & fecha_nac & "'"
StrSql = StrSql & ",'" & est_civil & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & "," & estado
StrSql = StrSql & ",'" & reg_prev & "'"
StrSql = StrSql & ",'" & lug_trab & "'"
StrSql = StrSql & "," & numberForSQL(basico)
StrSql = StrSql & "," & numberForSQL(neto)
StrSql = StrSql & "," & numberForSQL(msr)
StrSql = StrSql & "," & numberForSQL(asi_flia)
StrSql = StrSql & "," & numberForSQL(dtos)
StrSql = StrSql & "," & numberForSQL(bruto)
StrSql = StrSql & ",'" & prodesc & "'"
StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & "," & ConvFecha(profecpago)
StrSql = StrSql & "," & pliqfecdep
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & "," & Pagina
StrSql = StrSql & ",'" & emprCuit & "'"
StrSql = StrSql & ",'" & emprDire & "'"
StrSql = StrSql & ",'" & emprActiv & "'"
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & centroCosto & "'"
StrSql = StrSql & ",'" & nroInscripcion & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Busco el detalle de la liquidacion
'------------------------------------------------------------------

StrSql = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, detliq.dlicant, detliq.dlimonto " & _
    " FROM cabliq " & _
    " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND proceso.pronro = " & proNro & _
    " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro " & _
    " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & _
    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 " & _
    " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "

'---LOG---
Flog.Writeline "Buscando datos del detalle de liquidacion"

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF

  StrSql = " INSERT INTO rep_libroley_det "
  StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
  StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
  StrSql = StrSql & " dlimonto) "
  StrSql = StrSql & " VALUES "
  StrSql = StrSql & "(" & NroProceso
  StrSql = StrSql & "," & ternro
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

'------------------------------------------------------------------
'Busco los datos de los familiares
'------------------------------------------------------------------

If (tipoFamiliares = "1") Or (tipoFamiliares = "2") Then

    StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc,tercero.ternro,tercero.terape, tercero.ternom, tercero.terfecnac, tercero.tersex, famest,famtrab,faminc,famDGIdesde " & _
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
    
    Do Until rsConsult.EOF
    
      GuardarFam = True
    
      If (tipoFamiliares = "2") Then
          GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg, cliqnro, concFamiliar01)
          If Not GuardarFam Then
             GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg2, cliqnro, concFamiliar02)
          End If
      End If
    
      If GuardarFam Then
    
          If IsNull(rsConsult!terfecnac) Then
             terfecnac = "NULL"
          Else
             terfecnac = ConvFecha(rsConsult!terfecnac)
          End If
        
          If IsNull(rsConsult!famfecvto) Then
             famfecvto = "NULL"
          Else
             famfecvto = ConvFecha(rsConsult!famfecvto)
          End If
        
          If IsNull(rsConsult!famDGIdesde) Then
             famDGIdesde = "NULL"
          Else
             famDGIdesde = ConvFecha(rsConsult!famDGIdesde)
          End If
          
          If IsNull(rsConsult!famDGIhasta) Then
             famDGIhasta = "NULL"
          Else
             famDGIhasta = ConvFecha(rsConsult!famDGIhasta)
          End If
          
        
          StrSql = " INSERT INTO rep_libroley_fam "
          StrSql = StrSql & " (bpronro, ternro , pronro, nrodoc,"
          StrSql = StrSql & " sigladoc, ternrofam, terape, ternom,"
          StrSql = StrSql & " terfecnac , tersex, famest, famtrab,"
          StrSql = StrSql & " faminc, famsalario, famfecvto, famCargaDGI,"
          StrSql = StrSql & " famDGIdesde , famDGIhasta, famemergencia, paredesc"
          StrSql = StrSql & " ) "
          StrSql = StrSql & " VALUES "
          StrSql = StrSql & "(" & NroProceso
          StrSql = StrSql & "," & ternro
          StrSql = StrSql & "," & proNro
          StrSql = StrSql & ",'" & rsConsult!nrodoc & "'"
          StrSql = StrSql & ",'" & rsConsult!sigladoc & "'"
          StrSql = StrSql & "," & rsConsult!ternro
          StrSql = StrSql & ",'" & rsConsult!terape & "'"
          StrSql = StrSql & ",'" & rsConsult!ternom & "'"
          StrSql = StrSql & "," & terfecnac
          StrSql = StrSql & "," & strForSQL(rsConsult!tersex)
          StrSql = StrSql & "," & strForSQL(rsConsult!famest)
          StrSql = StrSql & "," & strForSQL(rsConsult!famtrab)
          StrSql = StrSql & "," & strForSQL(rsConsult!faminc)
          StrSql = StrSql & "," & strForSQL(rsConsult!famsalario)
          StrSql = StrSql & "," & famfecvto
          StrSql = StrSql & "," & strForSQL(rsConsult!famCargaDGI)
          StrSql = StrSql & "," & famDGIdesde
          StrSql = StrSql & "," & famDGIhasta
          StrSql = StrSql & "," & strForSQL(rsConsult!famemergencia)
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

'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal empresa As Long)

Dim StrEmpl As String

    If NroProc > 0 Then
        
        StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc
    
    Else

        If tenro3 <> 0 Then ' esto ocurre solo cuando se seleccionan los tres niveles
                StrEmpl = " SELECT DISTINCT v_empleado.ternro, empleg, terape, estact1.tenro AS tenro1, estact1.estrnro AS estrnro1, " & _
                " estact2.tenro AS tenro2, estact2.estrnro AS estrnro2, estact3.tenro AS tenro3, estact3.estrnro AS estrnro3, nestact1.estrdabr AS nestr1, nestact2.estrdabr AS nestr2, nestact3.estrdabr AS nestr3  "
                StrEmpl = StrEmpl & " FROM cabliq "
                If listapronro = "" Or listapronro = "0" Then
                    StrEmpl = StrEmpl & " INNER JOIN v_empleado ON v_empleado.ternro = cabliq.empleado "
                Else
                    StrEmpl = StrEmpl & " INNER JOIN v_empleado ON v_empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
                End If
                 
                 StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = v_empleado.ternro AND his_estructura.tenro = 10 "
                 StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If empresa > 0 Then
                 StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
                End If
                
                 StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON v_empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                 StrEmpl = StrEmpl & " AND (estact1.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If estrnro1 <> 0 Then   'cuando se le asigna un valor al nivel 1
                    StrEmpl = StrEmpl & " AND estact1.estrnro =" & estrnro1
                End If
                 
                 StrEmpl = StrEmpl & " INNER JOIN his_estructura estact2 ON v_empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
                 StrEmpl = StrEmpl & " AND (estact2.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If estrnro2 <> 0 Then   'cuando se le asigna un valor al nivel 2
                    StrEmpl = StrEmpl & " AND estact2.estrnro =" & estrnro2
                End If
                 
                 StrEmpl = StrEmpl & " INNER JOIN his_estructura estact3 ON v_empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3 & _
                 " AND (estact3.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                If estrnro3 <> 0 Then   'cuando se le asigna un valor al nivel 3
                    StrEmpl = StrEmpl & " AND estact3.estrnro =" & estrnro3
                End If
                
                If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                    StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = v_empleado.ternro AND ter_doc.tidnro=10"
                End If
                 
                 StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro "
                 StrEmpl = StrEmpl & " INNER JOIN estructura nestact2 ON nestact2.estrnro = estact2.estrnro "
                 StrEmpl = StrEmpl & " INNER JOIN estructura nestact3 ON nestact3.estrnro = estact3.estrnro "
                 StrEmpl = StrEmpl & " WHERE " & filtro
                 StrEmpl = StrEmpl & " ORDER BY nestr1,nestr2,nestr3," & l_orden
        
        Else
                If tenro2 <> 0 Then   ' ocurre cuando se selecciono hasta el segundo nivel
                    
                    StrEmpl = "SELECT DISTINCT v_empleado.ternro, empleg, terape, estact1.tenro AS tenro1, estact1.estrnro AS estrnro1, " & _
                        " estact2.tenro AS tenro2, estact2.estrnro AS estrnro2, nestact1.estrdabr AS nestr1, nestact2.estrdabr AS nestr2 "
                    StrEmpl = StrEmpl & " FROM cabliq "
                    
                    If listapronro = "" Or listapronro = "0" Then
                        StrEmpl = StrEmpl & " INNER JOIN v_empleado ON v_empleado.ternro = cabliq.empleado "
                    Else
                        StrEmpl = StrEmpl & " INNER JOIN v_empleado ON v_empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
                    End If
                    
                    StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = v_empleado.ternro AND his_estructura.tenro = 10 "
                    StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                    If empresa > 0 Then
                     StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
                    End If
                
                    StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON v_empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                    StrEmpl = StrEmpl & " AND (estact1.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
                    
                    If estrnro1 <> 0 Then
                         StrEmpl = StrEmpl & " AND estact1.estrnro =" & estrnro1
                    End If
                    
                    StrEmpl = StrEmpl & " INNER JOIN his_estructura estact2 ON v_empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & _
                    " AND (estact2.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(fecEstr) & "))"
                    
                    If estrnro2 <> 0 Then
                        StrEmpl = StrEmpl & " AND estact2.estrnro =" & estrnro2
                    End If
                    
                    If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                        StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = v_empleado.ternro AND ter_doc.tidnro=10"
                    End If
                    
                    StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro "
                    StrEmpl = StrEmpl & " INNER JOIN estructura nestact2 ON nestact2.estrnro = estact2.estrnro "
                    StrEmpl = StrEmpl & " WHERE " & filtro
                    StrEmpl = StrEmpl & " ORDER BY nestr1,nestr2," & l_orden
        
                Else
                    If tenro1 <> 0 Then   ' Cuando solo selecionamos el primer nivel
                        StrEmpl = "SELECT DISTINCT v_empleado.ternro, empleg, terape, estact1.tenro AS tenro1, estact1.estrnro AS estrnro1, nestact1.estrdabr AS nestr1"
                        StrEmpl = StrEmpl & " FROM cabliq "
                        
                        If listapronro = "" Or listapronro = "0" Then
                            StrEmpl = StrEmpl & " INNER JOIN v_empleado ON v_empleado.ternro = cabliq.empleado "
                        Else
                            StrEmpl = StrEmpl & " INNER JOIN v_empleado ON v_empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
                        End If
                        
                        StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = v_empleado.ternro AND his_estructura.tenro = 10 "
                        StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                        If empresa > 0 Then
                         StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
                        End If
                
                        StrEmpl = StrEmpl & " INNER JOIN his_estructura estact1 ON v_empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
                        StrEmpl = StrEmpl & " AND (estact1.htetdesde<=" & ConvFecha(fecEstr) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(fecEstr) & "))"
                        
                        If estrnro1 <> 0 Then
                            StrEmpl = StrEmpl & " AND estact1.estrnro =" & estrnro1
                        End If
                        
                        If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                            StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = v_empleado.ternro AND ter_doc.tidnro=10"
                        End If
                        
                        StrEmpl = StrEmpl & " INNER JOIN estructura nestact1 ON nestact1.estrnro = estact1.estrnro "
                        StrEmpl = StrEmpl & " WHERE " & filtro
                        StrEmpl = StrEmpl & " ORDER BY nestr1," & l_orden
        
                    Else ' cuando no hay nivel de estructura seleccionado
                        StrEmpl = " SELECT DISTINCT v_empleado.ternro, empleg, terape FROM cabliq "
                        
                        If listapronro = "" Or listapronro = "0" Then
                            StrEmpl = StrEmpl & " INNER JOIN v_empleado ON v_empleado.ternro = cabliq.empleado "
                        Else
                            StrEmpl = StrEmpl & " INNER JOIN v_empleado ON v_empleado.ternro = cabliq.empleado AND cabliq.pronro IN (" & listapronro & ") "
                        End If
                        
                        StrEmpl = StrEmpl & " INNER JOIN his_estructura ON his_estructura.ternro = v_empleado.ternro AND his_estructura.tenro = 10 "
                        StrEmpl = StrEmpl & " AND (his_estructura.htetdesde<=" & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null OR his_estructura.htethasta>=" & ConvFecha(fecEstr) & "))"
                
                        If empresa > 0 Then
                         StrEmpl = StrEmpl & " AND his_estructura.estrnro = " & empresa
                        End If
                                                
                        If Trim(l_orden) = "ter_doc.nrodoc" Then  'EL RESTO DE LOS FILTROS (1)EMPRESA
                            StrEmpl = StrEmpl & " INNER JOIN ter_doc ON ter_doc.ternro = v_empleado.ternro AND ter_doc.tidnro=10"
                        End If
                        
                        StrEmpl = StrEmpl & " WHERE " & filtro
                        StrEmpl = StrEmpl & " ORDER BY " & l_orden
                    End If
                
                End If
        
        End If
    
    End If
   
    OpenRecordset StrEmpl, rsEmpl
    
    cantRegistros = rsEmpl.RecordCount
    totalEmpleados = cantRegistros
    
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
Dim sexo As Integer                 '1(Masc) / 2 (Fem) / 3 (Todos)
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
        sexo = Param_cur!Auxint5
        Estudia = Param_cur!Auxchar2
        Ayuda_Escolar = Param_cur!Auxchar3
        Suma_FliaNumerosa = True
        Paga_FliaNumerosa = True
        Trabaja_Conyuge = Param_cur!Auxchar5
        Retroactivo_Prenatal = Param_cur!Auxint2
        Nivel_Estudio = IIf(EsNulo(Param_cur!Auxchar1), 0, Param_cur!Auxchar1)
        Periodo_Escolar = Param_cur!Auxint3
        Parentesco = Param_cur!Auxchar4
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
             edad_f = Calcular_Edad(rs_Familiar!terfecnac)
             
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
                ((sexo = 1 And CBool(rs_Familiar!tersex)) Or (sexo = 2 And Not CBool(rs_Familiar!tersex)) Or sexo = 3) Then
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
Dim ternro
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
       ternro = 0
       Flog.Writeline "Error: Buscando datos de la empresa: al obtener el empleado"
    Else
       ternro = rsConsult!ternro
    End If
    
    rsConsult.Close
    
    '------------------------------------------------------------------
    'Busco los datos del proceso
    '------------------------------------------------------------------
    StrSql = " SELECT * FROM proceso "
    StrSql = StrSql & " WHERE proceso.pronro= " & proNro
    
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       profecpago = rsConsult!profecpago
    Else
       Flog.Writeline "Error: Buscando datos de la empresa: al obtener los datos del proceso"
    End If
    
    rsConsult.Close

    ' -------------------------------------------------------------------------
    ' Busco los datos de la empresa
    '--------------------------------------------------------------------------
    
    StrSql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom,empresa.empactiv " & _
        " From his_estructura " & _
        " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro " & _
        " WHERE his_estructura.htetdesde <= " & ConvFecha(profecpago) & " AND " & _
        " (his_estructura.htethasta >= " & ConvFecha(profecpago) & " OR his_estructura.htethasta IS NULL)" & _
        " AND his_estructura.ternro = " & ternro & _
        " AND his_estructura.tenro  = 10"
    
    '---LOG---
    Flog.Writeline "Buscando datos de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    emprNro = 0
    
    If rsConsult.EOF Then
        Flog.Writeline "No se encontró la empresa"
        Exit Sub
    Else
        empresa = rsConsult!empnom
        emprNro = rsConsult!estrnro
        emprActiv = rsConsult!empactiv
        emprTer = rsConsult!ternro
    End If
    
    rsConsult.Close
    
    'Consulta para obtener la direccion de la empresa
    StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc From cabdom " & _
        " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & emprTer & _
        " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
    
    '---LOG---
    Flog.Writeline "Buscando datos de la direccion de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.Writeline "No se encontró el domicilio de la empresa"
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
    Flog.Writeline "Buscando datos del cuit de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.Writeline "No se encontró el CUIT de la Empresa"
        'Exit Sub
        emprCuit = "  "
    Else
        emprCuit = rsConsult!nrodoc
    End If
    
    rsConsult.Close

End Sub

'--------------------------------------------------------------------
' Se encarga de generar los datos para Jugos
'--------------------------------------------------------------------
Sub generarDatosEmpleado05(proNro, ternro, descripcion, orden)

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
Dim direccion As String
Dim puesto As String
Dim documento  As String
Dim fecha_nac As String
Dim est_civil As String
Dim Cuil As String
Dim estado
Dim reg_prev As String
Dim lug_trab  As String
Dim basico As String
Dim neto As String
Dim msr As String
Dim asi_flia As String
Dim dtos As String
Dim bruto As String
Dim profecpago
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim terfecnac
Dim famfecvto
Dim famDGIdesde
Dim famDGIhasta
Dim cliqnro
Dim GuardarFam As Boolean
Dim pliqdesde As Date
Dim pliqhasta As Date

Dim centroCosto
Dim regimenHorario
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
   fecpago = rsConsult!profecpago
   profecpago = rsConsult!profecpago
   prodesc = rsConsult!prodesc
   pliqdesc = rsConsult!pliqdesc
   cliqnro = rsConsult!cliqnro
   pliqdesde = rsConsult!pliqdesde
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.Writeline "El empleado no se encuetra en el proceso"
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
   empfecalta = rsConsult!empfaltagr
Else
   Flog.Writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
' Obtengo los datos de las estructuras
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de las estructuras"

 sql = " SELECT empfaltagr, empleado.empest, empfecbaja,estr1.estrdabr AS categoria, estr2.estrdabr AS contrato, estr3.estrdabr AS puesto, estr4.estrdabr AS regprev, estr5.estrdabr AS sucursal "
 sql = sql & " FROM empleado "
 'Busco la categoria
 sql = sql & " LEFT JOIN his_estructura his1 ON his1.ternro = empleado.ternro AND his1.tenro = 3 AND empleado.ternro= " & ternro & " AND his1.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his1.htethasta IS NULL OR his1.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr1 ON estr1.estrnro = his1.estrnro "
 'Busco el contrato
 sql = sql & " LEFT JOIN his_estructura his2 ON his2.ternro = empleado.ternro AND his2.tenro = 18 AND empleado.ternro= " & ternro & " AND his2.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his2.htethasta IS NULL OR his2.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr2 ON estr2.estrnro = his2.estrnro "
 'Busco el puesto
 sql = sql & " LEFT JOIN his_estructura his3 ON his3.ternro = empleado.ternro AND his3.tenro = 4 AND empleado.ternro= " & ternro & " AND his3.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his3.htethasta IS NULL OR his3.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr3 ON estr3.estrnro = his3.estrnro "
 'Busco el regimen previsional
 sql = sql & " LEFT JOIN his_estructura his4 ON his4.ternro = empleado.ternro AND his4.tenro = 15 AND empleado.ternro= " & ternro & " AND his4.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his4.htethasta IS NULL OR his4.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr4 ON estr4.estrnro = his4.estrnro "
 'Busco la sucursal
 sql = sql & " LEFT JOIN his_estructura his5 ON his5.ternro = empleado.ternro AND his5.tenro = 1 AND empleado.ternro= " & ternro & " AND his5.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his5.htethasta IS NULL OR his5.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr5 ON estr5.estrnro = his5.estrnro "
    
 sql = sql & " WHERE empleado.ternro = " & ternro

 OpenRecordset sql, rsConsult

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

 'Contrato
 If IsNull(rsConsult!contrato) Then
   contrato = ""
 Else
   contrato = rsConsult!contrato
 End If
 
 'Categoria
 If IsNull(rsConsult!categoria) Then
   categoria = ""
 Else
   categoria = rsConsult!categoria
 End If
 
 'Puesto
 If IsNull(rsConsult!puesto) Then
   puesto = ""
 Else
   puesto = rsConsult!puesto
 End If
 
 'Reg. Prev.
 If IsNull(rsConsult!regprev) Then
   reg_prev = ""
 Else
   reg_prev = rsConsult!regprev
 End If
 
 'Sucursal
 If IsNull(rsConsult!sucursal) Then
   lug_trab = ""
 Else
   lug_trab = rsConsult!sucursal
 End If

 estado = rsConsult!empest
 
 rsConsult.Close
 
'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos del centro de costo"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult

centroCosto = ""

If Not rsConsult.EOF Then
   centroCosto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor del regimen horario
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos del regimen horario"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=21 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult

regimenHorario = ""

If Not rsConsult.EOF Then
   regimenHorario = rsConsult!estrdabr
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
'Busco el valor del fecha nac, cuil, doc y tipo doc
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de los documentos"

sql = " SELECT tercero.terfecnac, cuil.nrodoc AS nrocuil, docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc, estcivil.estcivdesabr"
sql = sql & " FROM tercero "
sql = sql & " LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
sql = sql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
sql = sql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
sql = sql & " LEFT JOIN estcivil ON estcivil.estcivnro= tercero.estcivnro "
sql = sql & " WHERE tercero.ternro= " & ternro

OpenRecordset sql, rsConsult

Cuil = ""
documento = ""
fecha_nac = ""
est_civil = 0

If Not rsConsult.EOF Then
   Cuil = rsConsult!nrocuil
   documento = rsConsult!sigladoc & "-" & rsConsult!nrodoc
   fecha_nac = rsConsult!terfecnac
   est_civil = rsConsult!estcivdesabr
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------

'Basico
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_basico
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del basico"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   basico = 0
Else
   basico = rsConsult!almonto
End If

rsConsult.Close

'Neto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_neto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del neto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   neto = 0
Else
   neto = rsConsult!almonto
End If

rsConsult.Close

'MSR
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_msr
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del msr"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   msr = 0
Else
   msr = rsConsult!almonto
End If

rsConsult.Close

'Asig. Familia
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_asi_flia
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de la asig. familia"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   asi_flia = 0
Else
   asi_flia = rsConsult!almonto
End If

rsConsult.Close

'Descuentos
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_Dtos
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de los descuentos"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   dtos = 0
Else
   dtos = rsConsult!almonto
End If

rsConsult.Close

'Bruto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_bruto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del bruto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   bruto = 0
Else
   bruto = rsConsult!almonto
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la direccion y localidad del empleado
'------------------------------------------------------------------

StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,detdom.piso,detdom.oficdepto,detdom.barrio"
StrSql = StrSql & " FROM  cabdom "
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
StrSql = StrSql & " WHERE  cabdom.domdefault = -1 AND cabdom.ternro = " & ternro
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   direccion = IIf(sinDatos(rsConsult!calle), "", rsConsult!calle)
   direccion = direccion & " " & IIf(sinDatos(rsConsult!Nro), "", rsConsult!Nro)
   direccion = direccion & " " & IIf(sinDatos(rsConsult!piso), "", rsConsult!piso & "P")
   direccion = direccion & " " & IIf(sinDatos(rsConsult!oficdepto), "", """" & rsConsult!oficdepto & """")
   direccion = direccion & " " & IIf(sinDatos(rsConsult!barrio), "", rsConsult!barrio)
   direccion = direccion & "," & IIf(sinDatos(rsConsult!locdesc), "", rsConsult!locdesc)
   
Else
   direccion = ""
'   Flog.writeline "Error al obtener los datos de la localidad"
'   GoTo MError
End If

'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & emprNro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!banco
Else
   pliqfecdep = ""
   pliqbco = ""
   Flog.Writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

If IsNull(pliqfecdep) Then
   pliqfecdep = "NULL"
Else
   If Trim(pliqfecdep) = "" Then
      pliqfecdep = "NULL"
   Else
      pliqfecdep = ConvFecha(pliqfecdep)
   End If
End If

StrSql = " INSERT INTO rep_libroley "
StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
StrSql = StrSql & " empresa , emprnro, pliqnro, fecpago,"
StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
StrSql = StrSql & " direccion , puesto, documento, fecha_nac,"
StrSql = StrSql & " est_civil , cuil, estado, reg_prev,"
StrSql = StrSql & " lug_trab , basico, neto, msr,"
StrSql = StrSql & " asi_flia , dtos, bruto, "
StrSql = StrSql & " prodesc , descripcion, pliqdesc, pliqmes, "
StrSql = StrSql & " pliqanio , profecpago, pliqfecdep, pliqbco,ultima_pag_impr, auxchar1, auxchar2, auxchar3, "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,auxchar4,auxchar5) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & proNro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & empresa & "'"
StrSql = StrSql & "," & emprNro
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & fecpago & "'"
StrSql = StrSql & ",'" & fecalta & "'"
StrSql = StrSql & ",'" & fecbaja & "'"
StrSql = StrSql & ",'" & contrato & "'"
StrSql = StrSql & ",'" & categoria & "'"
StrSql = StrSql & ",'" & direccion & "'"
StrSql = StrSql & ",'" & puesto & "'"
StrSql = StrSql & ",'" & documento & "'"
StrSql = StrSql & ",'" & fecha_nac & "'"
StrSql = StrSql & ",'" & est_civil & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & "," & estado
StrSql = StrSql & ",'" & reg_prev & "'"
StrSql = StrSql & ",'" & lug_trab & "'"
StrSql = StrSql & "," & numberForSQL(basico)
StrSql = StrSql & "," & numberForSQL(neto)
StrSql = StrSql & "," & numberForSQL(msr)
StrSql = StrSql & "," & numberForSQL(asi_flia)
StrSql = StrSql & "," & numberForSQL(dtos)
StrSql = StrSql & "," & numberForSQL(bruto)
StrSql = StrSql & ",'" & prodesc & "'"
StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & "," & ConvFecha(profecpago)
StrSql = StrSql & "," & pliqfecdep
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & "," & Pagina
StrSql = StrSql & ",'" & emprCuit & "'"
StrSql = StrSql & ",'" & emprDire & "'"
StrSql = StrSql & ",'" & emprActiv & "'"
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & centroCosto & "'"
StrSql = StrSql & ",'" & regimenHorario & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Busco el detalle de la liquidacion
'------------------------------------------------------------------

StrSql = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, detliq.dlicant, detliq.dlimonto " & _
    " FROM cabliq " & _
    " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND proceso.pronro = " & proNro & _
    " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro " & _
    " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & _
    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 " & _
    " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "

'---LOG---
Flog.Writeline "Buscando datos del detalle de liquidacion"

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF

  StrSql = " INSERT INTO rep_libroley_det "
  StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
  StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
  StrSql = StrSql & " dlimonto) "
  StrSql = StrSql & " VALUES "
  StrSql = StrSql & "(" & NroProceso
  StrSql = StrSql & "," & ternro
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

'------------------------------------------------------------------
'Busco los datos de los familiares
'------------------------------------------------------------------

If (tipoFamiliares = "1") Or (tipoFamiliares = "2") Then

    StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc,tercero.ternro,tercero.terape, tercero.ternom, tercero.terfecnac, tercero.tersex, famest,famtrab,faminc,famDGIdesde " & _
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
    
    Do Until rsConsult.EOF
    
      GuardarFam = True
    
      If (tipoFamiliares = "2") Then
          GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg, cliqnro, concFamiliar01)
          If Not GuardarFam Then
             GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg2, cliqnro, concFamiliar02)
          End If
      End If
    
      If GuardarFam Then
      
          If IsNull(rsConsult!terfecnac) Then
             terfecnac = "NULL"
          Else
             terfecnac = ConvFecha(rsConsult!terfecnac)
          End If
        
          If IsNull(rsConsult!famfecvto) Then
             famfecvto = "NULL"
          Else
             famfecvto = ConvFecha(rsConsult!famfecvto)
          End If
        
          If IsNull(rsConsult!famDGIdesde) Then
             famDGIdesde = "NULL"
          Else
             famDGIdesde = ConvFecha(rsConsult!famDGIdesde)
          End If
          
          If IsNull(rsConsult!famDGIhasta) Then
             famDGIhasta = "NULL"
          Else
             famDGIhasta = ConvFecha(rsConsult!famDGIhasta)
          End If
          
        
          StrSql = " INSERT INTO rep_libroley_fam "
          StrSql = StrSql & " (bpronro, ternro , pronro, nrodoc,"
          StrSql = StrSql & " sigladoc, ternrofam, terape, ternom,"
          StrSql = StrSql & " terfecnac , tersex, famest, famtrab,"
          StrSql = StrSql & " faminc, famsalario, famfecvto, famCargaDGI,"
          StrSql = StrSql & " famDGIdesde , famDGIhasta, famemergencia, paredesc"
          StrSql = StrSql & " ) "
          StrSql = StrSql & " VALUES "
          StrSql = StrSql & "(" & NroProceso
          StrSql = StrSql & "," & ternro
          StrSql = StrSql & "," & proNro
          StrSql = StrSql & ",'" & rsConsult!nrodoc & "'"
          StrSql = StrSql & ",'" & rsConsult!sigladoc & "'"
          StrSql = StrSql & "," & rsConsult!ternro
          StrSql = StrSql & ",'" & rsConsult!terape & "'"
          StrSql = StrSql & ",'" & rsConsult!ternom & "'"
          StrSql = StrSql & "," & terfecnac
          StrSql = StrSql & "," & strForSQL(rsConsult!tersex)
          StrSql = StrSql & "," & strForSQL(rsConsult!famest)
          StrSql = StrSql & "," & strForSQL(rsConsult!famtrab)
          StrSql = StrSql & "," & strForSQL(rsConsult!faminc)
          StrSql = StrSql & "," & strForSQL(rsConsult!famsalario)
          StrSql = StrSql & "," & famfecvto
          StrSql = StrSql & "," & strForSQL(rsConsult!famCargaDGI)
          StrSql = StrSql & "," & famDGIdesde
          StrSql = StrSql & "," & famDGIhasta
          StrSql = StrSql & "," & strForSQL(rsConsult!famemergencia)
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


'--------------------------------------------------------------------
' Se encarga de generar los datos para Roche
'--------------------------------------------------------------------
Sub generarDatosEmpleado06(proNro, ternro, descripcion, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim Cont As Long

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim pliqnro As Integer
Dim pliqmes
Dim pliqanio
Dim fecpago As String
Dim empfecalta
Dim fecalta As String
Dim fecbaja As String
Dim contrato As String
Dim categoria As String
Dim direccion As String
Dim puesto As String
Dim documento  As String
Dim fecha_nac As String
Dim est_civil As String
Dim Cuil As String
Dim estado As Integer
Dim reg_prev As String
Dim lug_trab  As String
Dim basico As String
Dim neto As String
Dim msr As String
Dim asi_flia As String
Dim dtos As String
Dim bruto As String
Dim profecpago
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim terfecnac
Dim famfecvto
Dim famDGIdesde
Dim famDGIhasta
Dim sexo
Dim Sueldo
Dim nacionalidad
Dim causaEgreso
Dim cliqnro
Dim GuardarFam As Boolean
Dim pliqdesde As Date
Dim pliqhasta As Date

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
   fecpago = rsConsult!profecpago
   profecpago = rsConsult!profecpago
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
StrSql = " SELECT empleado.empleg,empleado.terape,empleado.terape2,empleado.ternom,empleado.ternom2,empleado.empfaltagr,empleado.empremu,tercero.tersex "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro AND empleado.ternro= " & ternro
StrSql = StrSql & " WHERE empleado.ternro= " & ternro
       
'---LOG---
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
   empfecalta = rsConsult!empfaltagr
   Sueldo = rsConsult!empremu
   If CInt(rsConsult!tersex) = -1 Then
     sexo = "M"
   Else
     sexo = "F"
   End If
Else
   Flog.Writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close


'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & emprNro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!banco
Else
   pliqfecdep = ""
   pliqbco = ""
   Flog.Writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
' Obtengo los datos de las estructuras
'------------------------------------------------------------------

 sql = " SELECT empfaltagr, empleado.empest, empfecbaja,estr1.estrdabr AS categoria, estr2.estrdabr AS contrato, estr3.estrdabr AS puesto, estr4.estrdabr AS regprev, estr5.estrdabr AS sucursal "
 sql = sql & " FROM empleado "
 'Busco la categoria
 sql = sql & " LEFT JOIN his_estructura his1 ON his1.ternro = empleado.ternro AND his1.tenro = 3 AND empleado.ternro= " & ternro & " AND his1.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his1.htethasta IS NULL OR his1.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr1 ON estr1.estrnro = his1.estrnro "
 'Busco el contrato
 sql = sql & " LEFT JOIN his_estructura his2 ON his2.ternro = empleado.ternro AND his2.tenro = 18 AND empleado.ternro= " & ternro & " AND his2.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his2.htethasta IS NULL OR his2.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr2 ON estr2.estrnro = his2.estrnro "
 'Busco el puesto
 sql = sql & " LEFT JOIN his_estructura his3 ON his3.ternro = empleado.ternro AND his3.tenro = 4 AND empleado.ternro= " & ternro & " AND his3.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his3.htethasta IS NULL OR his3.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr3 ON estr3.estrnro = his3.estrnro "
 'Busco el regimen previsional
 sql = sql & " LEFT JOIN his_estructura his4 ON his4.ternro = empleado.ternro AND his4.tenro = 15 AND empleado.ternro= " & ternro & " AND his4.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his4.htethasta IS NULL OR his4.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr4 ON estr4.estrnro = his4.estrnro "
 'Busco la sucursal
 sql = sql & " LEFT JOIN his_estructura his5 ON his5.ternro = empleado.ternro AND his5.tenro = 1 AND empleado.ternro= " & ternro & " AND his5.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his5.htethasta IS NULL OR his5.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr5 ON estr5.estrnro = his5.estrnro "
    
 sql = sql & " WHERE empleado.ternro = " & ternro
 
'---LOG---
Flog.Writeline "Buscando datos de las estructuras"

 OpenRecordset sql, rsConsult

 'Fecha Alta
 If IsNull(rsConsult!empfaltagr) Then
    fecalta = ""
 Else
    fecalta = rsConsult!empfaltagr
 End If
 
 'Fecha Baja
' If IsNull(rsConsult!empfecbaja) Then
'    fecbaja = ""
' Else
'    fecbaja = rsConsult!empfecbaja
' End If

 'Contrato
 If IsNull(rsConsult!contrato) Then
   contrato = ""
 Else
   contrato = rsConsult!contrato
 End If
 
 'Categoria
 If IsNull(rsConsult!categoria) Then
   categoria = ""
 Else
   categoria = rsConsult!categoria
 End If
 
 'Puesto
 If IsNull(rsConsult!puesto) Then
   puesto = ""
 Else
   puesto = rsConsult!puesto
 End If
 
 'Reg. Prev.
 If IsNull(rsConsult!regprev) Then
   reg_prev = ""
 Else
   reg_prev = rsConsult!regprev
 End If
 
 'Sucursal
 If IsNull(rsConsult!sucursal) Then
   lug_trab = ""
 Else
   lug_trab = rsConsult!sucursal
 End If

 estado = rsConsult!empest
 
 rsConsult.Close
 
'------------------------------------------------------------------
'Busco los datos del tipos de estructura 1
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos del estructura 1"

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
Flog.Writeline "Buscando datos del estructura 2"

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
Flog.Writeline "Buscando datos del estructura 3"

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
'Busco el valor de la direccion
'------------------------------------------------------------------

StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,detdom.piso,detdom.oficdepto,detdom.barrio"
StrSql = StrSql & " FROM  cabdom "
StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
StrSql = StrSql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
StrSql = StrSql & " WHERE  cabdom.domdefault = -1 AND cabdom.ternro = " & ternro

'---LOG---
Flog.Writeline "Buscando datos de la direccion del empleado"

OpenRecordset StrSql, rsConsult

direccion = " "

If Not rsConsult.EOF Then
   direccion = rsConsult!calle & " " & rsConsult!Nro & " " & rsConsult!piso & " " & rsConsult!oficdepto & ", " & rsConsult!locdesc
End If

rsConsult.Close
    
'------------------------------------------------------------------
'Busco el valor del fecha nac, cuil, doc y tipo doc
'------------------------------------------------------------------

sql = " SELECT tercero.terfecnac, cuil.nrodoc AS nrocuil, docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc, estcivil.estcivdesabr"
sql = sql & " FROM tercero "
sql = sql & " LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
sql = sql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
sql = sql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
sql = sql & " LEFT JOIN estcivil ON estcivil.estcivnro= tercero.estcivnro "
sql = sql & " WHERE tercero.ternro= " & ternro

'---LOG---
Flog.Writeline "Buscando datos del documento"

OpenRecordset sql, rsConsult

Cuil = ""
documento = ""
fecha_nac = ""
est_civil = 0

If Not rsConsult.EOF Then
   Cuil = rsConsult!nrocuil
   documento = rsConsult!sigladoc & "-" & rsConsult!nrodoc
   fecha_nac = rsConsult!terfecnac
   est_civil = rsConsult!estcivdesabr
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------

'Basico
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_basico
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del basico"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   basico = 0
Else
   basico = rsConsult!almonto
End If

rsConsult.Close

'Neto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_neto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del neto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   neto = 0
Else
   neto = rsConsult!almonto
End If

rsConsult.Close

'MSR
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_msr
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del MSR"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   msr = 0
Else
   msr = rsConsult!almonto
End If

rsConsult.Close

'Asig. Familia
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_asi_flia
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de la asig. fam."

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   asi_flia = 0
Else
   asi_flia = rsConsult!almonto
End If

rsConsult.Close

'Descuentos
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_Dtos
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos descuentos"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   dtos = 0
Else
   dtos = rsConsult!almonto
End If

rsConsult.Close

'Bruto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_bruto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del bruto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   bruto = 0
Else
   bruto = rsConsult!almonto
End If

rsConsult.Close

If IsNull(Sueldo) Then
   Sueldo = bruto
End If

' -------------------------------------------------------------------------
' Busco la nacionalidad
'--------------------------------------------------------------------------

StrSql = "SELECT nacionaldes " & _
    " FROM tercero " & _
    " INNER JOIN nacionalidad ON nacionalidad.nacionalnro = tercero.nacionalnro AND tercero.ternro = " & ternro

'---LOG---
Flog.Writeline "Buscando datos de la nacionalidad"

OpenRecordset StrSql, rsConsult

nacionalidad = ""

If rsConsult.EOF Then
    Flog.Writeline "No se encontró la nacionalidad"
'    Exit Sub
Else
    nacionalidad = rsConsult!nacionaldes
End If

rsConsult.Close

' -------------------------------------------------------------------------
' Busco la fecha de baja y la causa de egreso
'--------------------------------------------------------------------------

causaEgreso = ""
fecbaja = ""

'---LOG---
Flog.Writeline "Buscando datos de la causa de egreso"

 StrSql = " SELECT altfec, bajfec, caudes, estado, empatareas "
 StrSql = StrSql & " FROM fases "
 StrSql = StrSql & " LEFT JOIN empant ON fases.empantnro=empant.empantnro "
 StrSql = StrSql & " LEFT JOIN causa ON fases.caunro=causa.caunro "
 StrSql = StrSql & " WHERE fases.empleado=" & ternro & " order by altfec DESC"

 OpenRecordset StrSql, rsConsult

 If rsConsult.EOF Then
     Flog.Writeline "No se encontró la causa de egreso"
 '    Exit Sub
 Else
    If Not CBool(rsConsult!estado) Then
       fecbaja = rsConsult!bajfec
       causaEgreso = rsConsult!caudes
    End If
 End If
 
 rsConsult.Close

'------------------------------------------------------------------
'Busco la direccion de la empresa
'------------------------------------------------------------------
 
 'Consulta para obtener la direccion de la empresa
StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc,codigopostal,barrio From cabdom " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & emprTer & _
    " INNER JOIN localidad ON detdom.locnro = localidad.locnro "

OpenRecordset StrSql, rsConsult

If rsConsult.EOF Then
    Flog.Writeline "No se encontró el domicilio de la empresa"
    'Exit Sub
    emprDire = "   "
Else
    emprDire = rsConsult!calle & " " & rsConsult!Nro & "<br>" & rsConsult!codigopostal & " " & rsConsult!barrio & " - " & rsConsult!locdesc
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

If IsNull(pliqfecdep) Or (pliqfecdep = "") Then
   pliqfecdep = "NULL"
Else
   pliqfecdep = ConvFecha(pliqfecdep)
End If

StrSql = " INSERT INTO rep_libroley "
StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
StrSql = StrSql & " empresa , emprnro, pliqnro, fecpago,"
StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
StrSql = StrSql & " direccion , puesto, documento, fecha_nac,"
StrSql = StrSql & " est_civil , cuil, estado, reg_prev,"
StrSql = StrSql & " lug_trab , basico, neto, msr,"
StrSql = StrSql & " asi_flia , dtos, bruto, "
StrSql = StrSql & " prodesc , descripcion, pliqdesc, pliqmes, "
StrSql = StrSql & " pliqanio , profecpago, pliqfecdep, pliqbco,ultima_pag_impr,auxchar1,auxchar2,auxchar3,auxchar4, "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,auxchar5) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & proNro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & empresa & "'"
StrSql = StrSql & "," & emprNro
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & fecpago & "'"
StrSql = StrSql & ",'" & fecalta & "'"
StrSql = StrSql & ",'" & fecbaja & "'"
StrSql = StrSql & ",'" & contrato & "'"
StrSql = StrSql & ",'" & categoria & "'"
StrSql = StrSql & ",'" & direccion & "'"
StrSql = StrSql & ",'" & puesto & "'"
StrSql = StrSql & ",'" & documento & "'"
StrSql = StrSql & ",'" & fecha_nac & "'"
StrSql = StrSql & ",'" & est_civil & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & "," & estado
StrSql = StrSql & ",'" & reg_prev & "'"
StrSql = StrSql & ",'" & lug_trab & "'"
StrSql = StrSql & "," & numberForSQL(basico)
StrSql = StrSql & "," & numberForSQL(neto)
StrSql = StrSql & "," & numberForSQL(msr)
StrSql = StrSql & "," & numberForSQL(asi_flia)
StrSql = StrSql & "," & numberForSQL(dtos)
StrSql = StrSql & "," & numberForSQL(bruto)
StrSql = StrSql & ",'" & prodesc & "'"
StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & "," & ConvFecha(profecpago)
StrSql = StrSql & "," & pliqfecdep
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & "," & Pagina
StrSql = StrSql & ",'" & nacionalidad & "'"
StrSql = StrSql & ",'" & causaEgreso & "'"
StrSql = StrSql & ",'" & emprCuit & "'"
StrSql = StrSql & ",'" & emprDire & "'"
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & "," & orden
StrSql = StrSql & "," & controlNull(emprActiv)
StrSql = StrSql & ")"


'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Busco el detalle de la liquidacion
'------------------------------------------------------------------

'---------------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado, de la columna 1
'---------------------------------------------------------------------------

For Cont = 1 To cantAcumGrupo1

    StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & proNro
    StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
    StrSql = StrSql & " INNER JOIN con_acum ON concepto.concnro = con_acum.concnro AND con_acum.acunro = " & acumGrupo1(Cont)
    StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
        
    OpenRecordset StrSql, rsConsult
    
    Do Until rsConsult.EOF
      
        StrSql = " SELECT ternro FROM rep_libroley_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND ternro = " & ternro
        StrSql = StrSql & " AND pronro = " & proNro
        StrSql = StrSql & " AND concnro = " & rsConsult!concnro
        
        OpenRecordset StrSql, rsConsult2
        
        If rsConsult2.EOF Then
        
              StrSql = " INSERT INTO rep_libroley_det "
              StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
              StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
              StrSql = StrSql & " dlimonto, conctipo) "
              StrSql = StrSql & " VALUES "
              StrSql = StrSql & "(" & NroProceso
              StrSql = StrSql & "," & ternro
              StrSql = StrSql & "," & proNro
              StrSql = StrSql & ",'" & rsConsult!concabr & "'"
              StrSql = StrSql & ",'" & rsConsult!Conccod & "'"
              StrSql = StrSql & "," & rsConsult!concnro
              StrSql = StrSql & "," & acumGrupo1(Cont)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
              StrSql = StrSql & ",'1')"
              
              objConn.Execute StrSql, , adExecuteNoRecords
              
        End If
        
        rsConsult2.Close
        
        rsConsult.MoveNext
    
    Loop
    
    rsConsult.Close

Next

'---------------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado, de la columna 2
'---------------------------------------------------------------------------

For Cont = 1 To cantAcumGrupo2
    
    StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & proNro
    StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
    StrSql = StrSql & " INNER JOIN con_acum ON concepto.concnro = con_acum.concnro AND con_acum.acunro = " & acumGrupo2(Cont)
    StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
        
    OpenRecordset StrSql, rsConsult
    
    Do Until rsConsult.EOF
      
        StrSql = " SELECT ternro FROM rep_libroley_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND ternro = " & ternro
        StrSql = StrSql & " AND pronro = " & proNro
        StrSql = StrSql & " AND concnro = " & rsConsult!concnro
        
        OpenRecordset StrSql, rsConsult2
        
        If rsConsult2.EOF Then
            
              StrSql = " INSERT INTO rep_libroley_det "
              StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
              StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
              StrSql = StrSql & " dlimonto, conctipo) "
              StrSql = StrSql & " VALUES "
              StrSql = StrSql & "(" & NroProceso
              StrSql = StrSql & "," & ternro
              StrSql = StrSql & "," & proNro
              StrSql = StrSql & ",'" & rsConsult!concabr & "'"
              StrSql = StrSql & ",'" & rsConsult!Conccod & "'"
              StrSql = StrSql & "," & rsConsult!concnro
              StrSql = StrSql & "," & acumGrupo1(Cont)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
              StrSql = StrSql & ",'2')"
              
              objConn.Execute StrSql, , adExecuteNoRecords
              
        End If
        
        rsConsult2.Close
        
        rsConsult.MoveNext
    
    Loop
    
    rsConsult.Close

Next


'---------------------------------------------------------------------------
'Obtengo los datos del los conceptos del empleado, de la columna 3
'---------------------------------------------------------------------------

For Cont = 1 To cantAcumGrupo3
    
    StrSql = " SELECT cabliq.cliqnro, concepto.concabr, concepto.conccod, concepto.concnro, concepto.tconnro, concepto.concimp, detliq.dlicant, detliq.dlimonto,cabliq.pronro,proceso.prodesc, periodo.pliqdesc, periodo.pliqnro,periodo.pliqmes,periodo.pliqanio "
    StrSql = StrSql & " FROM cabliq "
    StrSql = StrSql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & proNro
    StrSql = StrSql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
    StrSql = StrSql & " INNER JOIN con_acum ON concepto.concnro = con_acum.concnro AND con_acum.acunro = " & acumGrupo3(Cont)
    StrSql = StrSql & " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "
        
    OpenRecordset StrSql, rsConsult
    
    Do Until rsConsult.EOF
      
        StrSql = " SELECT ternro FROM rep_libroley_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND ternro = " & ternro
        StrSql = StrSql & " AND pronro = " & proNro
        StrSql = StrSql & " AND concnro = " & rsConsult!concnro
        
        OpenRecordset StrSql, rsConsult2
        
        If rsConsult2.EOF Then
            
              StrSql = " INSERT INTO rep_libroley_det "
              StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
              StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
              StrSql = StrSql & " dlimonto, conctipo) "
              StrSql = StrSql & " VALUES "
              StrSql = StrSql & "(" & NroProceso
              StrSql = StrSql & "," & ternro
              StrSql = StrSql & "," & proNro
              StrSql = StrSql & ",'" & rsConsult!concabr & "'"
              StrSql = StrSql & ",'" & rsConsult!Conccod & "'"
              StrSql = StrSql & "," & rsConsult!concnro
              StrSql = StrSql & "," & acumGrupo1(Cont)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlicant)
              StrSql = StrSql & "," & numberForSQL(rsConsult!dlimonto)
              StrSql = StrSql & ",'3')"
              
              objConn.Execute StrSql, , adExecuteNoRecords
              
        End If
        
        rsConsult2.Close
        
        rsConsult.MoveNext
    
    Loop
    
    rsConsult.Close

Next

'------------------------------------------------------------------
'Busco los datos de los familiares
'------------------------------------------------------------------

If (tipoFamiliares = "1") Or (tipoFamiliares = "2") Then

    StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc,tercero.ternro,tercero.terape, tercero.ternom, tercero.terfecnac, tercero.tersex, famest,famtrab,faminc,famDGIdesde " & _
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
    
    Do Until rsConsult.EOF
    
      GuardarFam = True
    
      If (tipoFamiliares = "2") Then
          GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg, cliqnro, concFamiliar01)
          If Not GuardarFam Then
             GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg2, cliqnro, concFamiliar02)
          End If
      End If
    
      If GuardarFam Then
   
          If IsNull(rsConsult!terfecnac) Then
             terfecnac = "NULL"
          Else
             terfecnac = ConvFecha(rsConsult!terfecnac)
          End If
        
          If IsNull(rsConsult!famfecvto) Then
             famfecvto = "NULL"
          Else
             famfecvto = ConvFecha(rsConsult!famfecvto)
          End If
        
          If IsNull(rsConsult!famDGIdesde) Then
             famDGIdesde = "NULL"
          Else
             famDGIdesde = ConvFecha(rsConsult!famDGIdesde)
          End If
          
          If IsNull(rsConsult!famDGIhasta) Then
             famDGIhasta = "NULL"
          Else
             famDGIhasta = ConvFecha(rsConsult!famDGIhasta)
          End If
          
        
          StrSql = " INSERT INTO rep_libroley_fam "
          StrSql = StrSql & " (bpronro, ternro , pronro, nrodoc,"
          StrSql = StrSql & " sigladoc, ternrofam, terape, ternom,"
          StrSql = StrSql & " terfecnac , tersex, famest, famtrab,"
          StrSql = StrSql & " faminc, famsalario, famfecvto, famCargaDGI,"
          StrSql = StrSql & " famDGIdesde , famDGIhasta, famemergencia, paredesc"
          StrSql = StrSql & " ) "
          StrSql = StrSql & " VALUES "
          StrSql = StrSql & "(" & NroProceso
          StrSql = StrSql & "," & ternro
          StrSql = StrSql & "," & proNro
          StrSql = StrSql & ",'" & rsConsult!nrodoc & "'"
          StrSql = StrSql & ",'" & rsConsult!sigladoc & "'"
          StrSql = StrSql & "," & rsConsult!ternro
          StrSql = StrSql & ",'" & rsConsult!terape & "'"
          StrSql = StrSql & ",'" & rsConsult!ternom & "'"
          StrSql = StrSql & "," & terfecnac
          StrSql = StrSql & "," & strForSQL(rsConsult!tersex)
          StrSql = StrSql & "," & strForSQL(rsConsult!famest)
          StrSql = StrSql & "," & strForSQL(rsConsult!famtrab)
          StrSql = StrSql & "," & strForSQL(rsConsult!faminc)
          StrSql = StrSql & "," & strForSQL(rsConsult!famsalario)
          StrSql = StrSql & "," & famfecvto
          StrSql = StrSql & "," & strForSQL(rsConsult!famCargaDGI)
          StrSql = StrSql & "," & famDGIdesde
          StrSql = StrSql & "," & famDGIhasta
          StrSql = StrSql & "," & strForSQL(rsConsult!famemergencia)
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


'--------------------------------------------------------------------
' Se encarga de generar los datos para el Standar y Deloitte
'--------------------------------------------------------------------
Sub generarDatosEmpleado07(proNro, ternro, descripcion, orden)

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
Dim direccion As String
Dim puesto As String
Dim documento  As String
Dim fecha_nac As String
Dim est_civil As String
Dim Cuil As String
Dim estado
Dim reg_prev As String
Dim lug_trab  As String
Dim basico As String
Dim neto As String
Dim msr As String
Dim asi_flia As String
Dim dtos As String
Dim bruto As String
Dim profecpago
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim terfecnac
Dim famfecvto
Dim famDGIdesde
Dim famDGIhasta
Dim cliqnro
Dim GuardarFam As Boolean
Dim pliqdesde As Date
Dim pliqhasta As Date
Dim ObraSocial

Dim centroCosto
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
   fecpago = rsConsult!profecpago
   profecpago = rsConsult!profecpago
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
   empfecalta = rsConsult!empfaltagr
Else
   Flog.Writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close


'------------------------------------------------------------------
' Obtengo los datos de las estructuras
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de las estructuras"

 sql = " SELECT empfaltagr, empleado.empest, empfecbaja,estr1.estrdabr AS categoria, estr2.estrdabr AS contrato, estr3.estrdabr AS puesto, estr4.estrdabr AS regprev, estr5.estrdabr AS sucursal "
 sql = sql & " FROM empleado "
 'Busco la categoria
 sql = sql & " LEFT JOIN his_estructura his1 ON his1.ternro = empleado.ternro AND his1.tenro = 3 AND empleado.ternro= " & ternro & " AND his1.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his1.htethasta IS NULL OR his1.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr1 ON estr1.estrnro = his1.estrnro "
 'Busco el contrato
 sql = sql & " LEFT JOIN his_estructura his2 ON his2.ternro = empleado.ternro AND his2.tenro = 18 AND empleado.ternro= " & ternro & " AND his2.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his2.htethasta IS NULL OR his2.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr2 ON estr2.estrnro = his2.estrnro "
 'Busco el puesto
 sql = sql & " LEFT JOIN his_estructura his3 ON his3.ternro = empleado.ternro AND his3.tenro = 4 AND empleado.ternro= " & ternro & " AND his3.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his3.htethasta IS NULL OR his3.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr3 ON estr3.estrnro = his3.estrnro "
 'Busco el regimen previsional
 sql = sql & " LEFT JOIN his_estructura his4 ON his4.ternro = empleado.ternro AND his4.tenro = 15 AND empleado.ternro= " & ternro & " AND his4.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his4.htethasta IS NULL OR his4.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr4 ON estr4.estrnro = his4.estrnro "
 'Busco la sucursal
 sql = sql & " LEFT JOIN his_estructura his5 ON his5.ternro = empleado.ternro AND his5.tenro = 1 AND empleado.ternro= " & ternro & " AND his5.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his5.htethasta IS NULL OR his5.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr5 ON estr5.estrnro = his5.estrnro "
    
 sql = sql & " WHERE empleado.ternro = " & ternro

 OpenRecordset sql, rsConsult

 'Fecha Alta
 If IsNull(rsConsult!empfaltagr) Then
    fecalta = ""
 Else
    fecalta = rsConsult!empfaltagr
 End If
 
 'Contrato
 If IsNull(rsConsult!contrato) Then
   contrato = ""
 Else
   contrato = rsConsult!contrato
 End If
 
 'Categoria
 If IsNull(rsConsult!categoria) Then
   categoria = ""
 Else
   categoria = rsConsult!categoria
 End If
 
 'Puesto
 If IsNull(rsConsult!puesto) Then
   puesto = ""
 Else
   puesto = rsConsult!puesto
 End If
 
 'Reg. Prev.
 If IsNull(rsConsult!regprev) Then
   reg_prev = ""
 Else
   reg_prev = rsConsult!regprev
 End If
 
 'Sucursal
 If IsNull(rsConsult!sucursal) Then
   lug_trab = ""
 Else
   lug_trab = rsConsult!sucursal
 End If

 estado = rsConsult!empest
 
 rsConsult.Close
 
 
 ' -------------------------------------------------------------------------
' Busco la fecha de baja
'--------------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de fecha de baja"

 StrSql = " SELECT altfec, bajfec "
 StrSql = StrSql & " FROM fases "
 StrSql = StrSql & " WHERE real = -1 AND fases.empleado=" & ternro & " order by altfec DESC"

 OpenRecordset StrSql, rsConsult

 If rsConsult.EOF Then
      fecbaja = ""
  Else
    If Not IsNull(rsConsult!bajfec) Then
       fecbaja = rsConsult!bajfec
    End If
 End If
 
 rsConsult.Close
 
'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos del centro de costo"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult

centroCosto = ""

If Not rsConsult.EOF Then
   centroCosto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
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
'Busco el valor del fecha nac, cuil, doc y tipo doc
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de los documentos"

sql = " SELECT tercero.terfecnac, cuil.nrodoc AS nrocuil, docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc, estcivil.estcivdesabr"
sql = sql & " FROM tercero "
sql = sql & " LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
sql = sql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
sql = sql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
sql = sql & " LEFT JOIN estcivil ON estcivil.estcivnro= tercero.estcivnro "
sql = sql & " WHERE tercero.ternro= " & ternro

OpenRecordset sql, rsConsult

Cuil = ""
documento = ""
fecha_nac = ""
est_civil = 0

If Not rsConsult.EOF Then
   Cuil = rsConsult!nrocuil
   documento = rsConsult!sigladoc & "-" & rsConsult!nrodoc
   fecha_nac = rsConsult!terfecnac
   est_civil = rsConsult!estcivdesabr
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la obra social
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de la obra social"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND his_estructura.tenro=17 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult

ObraSocial = ""

If Not rsConsult.EOF Then
   ObraSocial = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la obra social"
'   GoTo MError
End If

rsConsult.Close


'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------

'Basico
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_basico
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del basico"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   basico = 0
Else
   basico = rsConsult!almonto
End If

rsConsult.Close

'Neto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_neto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del neto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   neto = 0
Else
   neto = rsConsult!almonto
End If

rsConsult.Close

'MSR
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_msr
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del msr"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   msr = 0
Else
   msr = rsConsult!almonto
End If

rsConsult.Close

'Asig. Familia
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_asi_flia
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de la asig. familia"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   asi_flia = 0
Else
   asi_flia = rsConsult!almonto
End If

rsConsult.Close

'Descuentos
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_Dtos
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de los descuentos"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   dtos = 0
Else
   dtos = rsConsult!almonto
End If

rsConsult.Close

'Bruto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_bruto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del bruto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   bruto = 0
Else
   bruto = rsConsult!almonto
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la direccion
'------------------------------------------------------------------

direccion = ""

Select Case zonaDomicilio
   'Direccion de la sucursal
   Case 1

        sql = " SELECT detdom.calle,detdom.nro,localidad.locdesc"
        sql = sql & " From his_estructura"
        sql = sql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND tenro=1 AND his_estructura.ternro=" & ternro
        sql = sql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
        sql = sql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        sql = sql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
    
        '---LOG---
        Flog.Writeline "Buscando datos de la direccion de la sucursal"
        
        OpenRecordset sql, rsConsult
        
        If Not rsConsult.EOF Then
           direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
        End If
        
        rsConsult.Close
    
  'Direccion de la empresa
  Case 2

        sql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
            " From his_estructura" & _
            " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
            " WHERE his_estructura.htetdesde <=" & ConvFecha(fecEstr) & " AND " & _
            " (his_estructura.htethasta >= " & ConvFecha(fecEstr) & " OR his_estructura.htethasta IS NULL)" & _
            " AND his_estructura.ternro = " & ternro & _
            " AND his_estructura.tenro  = 10"
        
        OpenRecordset sql, rsConsult
        
        EmpTernro = 0
        
        If Not rsConsult.EOF Then
            EmpTernro = rsConsult!ternro
        End If
        
        rsConsult.Close
        
        'Consulta para obtener la direccion de la empresa
        sql = "SELECT detdom.calle,detdom.nro,localidad.locdesc From cabdom " & _
            " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & EmpTernro & _
            " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
        
        '---LOG---
        Flog.Writeline "Buscando datos de la direccion de la empresa"
        
        OpenRecordset sql, rsConsult
        
        If Not rsConsult.EOF Then
           direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
        End If
        
        rsConsult.Close
    
  'Direccion del empleado
  Case 3

        sql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,detdom.piso,detdom.oficdepto,detdom.barrio"
        sql = sql & " FROM  cabdom "
        sql = sql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        sql = sql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
        sql = sql & " WHERE  cabdom.domdefault = -1 AND cabdom.ternro = " & ternro
       
        '---LOG---
        Flog.Writeline "Buscando datos de la direccion del empleado"
        
        OpenRecordset sql, rsConsult
        
        If Not rsConsult.EOF Then
           direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
        End If
        
        rsConsult.Close
    
End Select


'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & emprNro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!banco
Else
   pliqfecdep = ""
   pliqbco = ""
   Flog.Writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

If IsNull(pliqfecdep) Then
   pliqfecdep = "NULL"
Else
   If Trim(pliqfecdep) = "" Then
      pliqfecdep = "NULL"
   Else
      pliqfecdep = ConvFecha(pliqfecdep)
   End If
End If

StrSql = " INSERT INTO rep_libroley "
StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
StrSql = StrSql & " empresa , emprnro, pliqnro, fecpago,"
StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
StrSql = StrSql & " direccion , puesto, documento, fecha_nac,"
StrSql = StrSql & " est_civil , cuil, estado, reg_prev,"
StrSql = StrSql & " lug_trab , basico, neto, msr,"
StrSql = StrSql & " asi_flia , dtos, bruto, "
StrSql = StrSql & " prodesc , descripcion, pliqdesc, pliqmes, "
StrSql = StrSql & " pliqanio , profecpago, pliqfecdep, pliqbco,ultima_pag_impr, auxchar1, auxchar2, auxchar3, "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,auxchar4, auxchar5) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & proNro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & empresa & "'"
StrSql = StrSql & "," & emprNro
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & fecpago & "'"
StrSql = StrSql & ",'" & fecalta & "'"
StrSql = StrSql & ",'" & fecbaja & "'"
StrSql = StrSql & ",'" & contrato & "'"
StrSql = StrSql & ",'" & categoria & "'"
StrSql = StrSql & ",'" & Mid(direccion, 1, 50) & "'"
StrSql = StrSql & ",'" & puesto & "'"
StrSql = StrSql & ",'" & documento & "'"
StrSql = StrSql & ",'" & fecha_nac & "'"
StrSql = StrSql & ",'" & est_civil & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & "," & estado
StrSql = StrSql & ",'" & reg_prev & "'"
StrSql = StrSql & ",'" & lug_trab & "'"
StrSql = StrSql & "," & numberForSQL(basico)
StrSql = StrSql & "," & numberForSQL(neto)
StrSql = StrSql & "," & numberForSQL(msr)
StrSql = StrSql & "," & numberForSQL(asi_flia)
StrSql = StrSql & "," & numberForSQL(dtos)
StrSql = StrSql & "," & numberForSQL(bruto)
StrSql = StrSql & ",'" & prodesc & "'"
StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & "," & ConvFecha(profecpago)
StrSql = StrSql & "," & pliqfecdep
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & "," & Pagina
StrSql = StrSql & ",'" & emprCuit & "'"
StrSql = StrSql & ",'" & emprDire & "'"
StrSql = StrSql & ",'" & emprActiv & "'"
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & centroCosto & "'"
StrSql = StrSql & ",'" & ObraSocial & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Busco el detalle de la liquidacion
'------------------------------------------------------------------

StrSql = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, detliq.dlicant, detliq.dlimonto " & _
    " FROM cabliq " & _
    " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND proceso.pronro = " & proNro & _
    " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro " & _
    " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & _
    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 " & _
    " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "

'---LOG---
Flog.Writeline "Buscando datos del detalle de liquidacion"

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF

  StrSql = " INSERT INTO rep_libroley_det "
  StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
  StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
  StrSql = StrSql & " dlimonto) "
  StrSql = StrSql & " VALUES "
  StrSql = StrSql & "(" & NroProceso
  StrSql = StrSql & "," & ternro
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

'------------------------------------------------------------------
'Busco los datos de los familiares
'------------------------------------------------------------------

If (tipoFamiliares = "1") Or (tipoFamiliares = "2") Then

    StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc,tercero.ternro,tercero.terape, tercero.ternom, tercero.terfecnac, tercero.tersex, famest,famtrab,faminc,famDGIdesde " & _
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
    
    Do Until rsConsult.EOF
    
      GuardarFam = True
    
      If (tipoFamiliares = "2") Then
          GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg, cliqnro, concFamiliar01)
          If Not GuardarFam Then
             GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg2, cliqnro, concFamiliar02)
          End If
      End If
    
      If GuardarFam Then
      
          If IsNull(rsConsult!terfecnac) Then
             terfecnac = "NULL"
          Else
             terfecnac = ConvFecha(rsConsult!terfecnac)
          End If
        
          If IsNull(rsConsult!famfecvto) Then
             famfecvto = "NULL"
          Else
             famfecvto = ConvFecha(rsConsult!famfecvto)
          End If
        
          If IsNull(rsConsult!famDGIdesde) Then
             famDGIdesde = "NULL"
          Else
             famDGIdesde = ConvFecha(rsConsult!famDGIdesde)
          End If
          
          If IsNull(rsConsult!famDGIhasta) Then
             famDGIhasta = "NULL"
          Else
             famDGIhasta = ConvFecha(rsConsult!famDGIhasta)
          End If
          
        
          StrSql = " INSERT INTO rep_libroley_fam "
          StrSql = StrSql & " (bpronro, ternro , pronro, nrodoc,"
          StrSql = StrSql & " sigladoc, ternrofam, terape, ternom,"
          StrSql = StrSql & " terfecnac , tersex, famest, famtrab,"
          StrSql = StrSql & " faminc, famsalario, famfecvto, famCargaDGI,"
          StrSql = StrSql & " famDGIdesde , famDGIhasta, famemergencia, paredesc"
          StrSql = StrSql & " ) "
          StrSql = StrSql & " VALUES "
          StrSql = StrSql & "(" & NroProceso
          StrSql = StrSql & "," & ternro
          StrSql = StrSql & "," & proNro
          StrSql = StrSql & ",'" & rsConsult!nrodoc & "'"
          StrSql = StrSql & ",'" & rsConsult!sigladoc & "'"
          StrSql = StrSql & "," & rsConsult!ternro
          StrSql = StrSql & ",'" & rsConsult!terape & "'"
          StrSql = StrSql & ",'" & rsConsult!ternom & "'"
          StrSql = StrSql & "," & terfecnac
          StrSql = StrSql & "," & strForSQL(rsConsult!tersex)
          StrSql = StrSql & "," & strForSQL(rsConsult!famest)
          StrSql = StrSql & "," & strForSQL(rsConsult!famtrab)
          StrSql = StrSql & "," & strForSQL(rsConsult!faminc)
          StrSql = StrSql & "," & strForSQL(rsConsult!famsalario)
          StrSql = StrSql & "," & famfecvto
          StrSql = StrSql & "," & strForSQL(rsConsult!famCargaDGI)
          StrSql = StrSql & "," & famDGIdesde
          StrSql = StrSql & "," & famDGIhasta
          StrSql = StrSql & "," & strForSQL(rsConsult!famemergencia)
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


'--------------------------------------------------------------------
' Se encarga de generar los datos para Promofilm
'--------------------------------------------------------------------
Sub generarDatosEmpleado08(proNro, ternro, descripcion, orden)

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
Dim direccion As String
Dim puesto As String
Dim documento  As String
Dim fecha_nac As String
Dim est_civil As String
Dim Cuil As String
Dim estado
Dim reg_prev As String
Dim lug_trab  As String
Dim basico As String
Dim neto As String
Dim msr As String
Dim asi_flia As String
Dim dtos As String
Dim bruto As String
Dim profecpago
Dim prodesc
Dim pliqdesc
Dim pliqfecdep
Dim pliqbco
Dim terfecnac
Dim famfecvto
Dim famDGIdesde
Dim famDGIhasta
Dim cliqnro
Dim GuardarFam As Boolean
Dim pliqdesde As Date
Dim pliqhasta As Date
Dim ObraSocial

Dim centroCosto
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
   fecpago = rsConsult!profecpago
   profecpago = rsConsult!profecpago
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
   empfecalta = rsConsult!empfaltagr
Else
   Flog.Writeline "Error al obtener los datos del empleado"
'   GoTo MError
End If

rsConsult.Close


'------------------------------------------------------------------
' Obtengo los datos de las estructuras
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de las estructuras"

 sql = " SELECT empfaltagr, empleado.empest, empfecbaja,estr1.estrdabr AS categoria, estr2.estrdabr AS contrato, estr3.estrdabr AS puesto, estr4.estrdabr AS regprev, estr5.estrdabr AS sucursal "
 sql = sql & " FROM empleado "
 'Busco la categoria
 sql = sql & " LEFT JOIN his_estructura his1 ON his1.ternro = empleado.ternro AND his1.tenro = 3 AND empleado.ternro= " & ternro & " AND his1.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his1.htethasta IS NULL OR his1.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr1 ON estr1.estrnro = his1.estrnro "
 'Busco el contrato
 sql = sql & " LEFT JOIN his_estructura his2 ON his2.ternro = empleado.ternro AND his2.tenro = 18 AND empleado.ternro= " & ternro & " AND his2.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his2.htethasta IS NULL OR his2.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr2 ON estr2.estrnro = his2.estrnro "
 'Busco el puesto
 sql = sql & " LEFT JOIN his_estructura his3 ON his3.ternro = empleado.ternro AND his3.tenro = 4 AND empleado.ternro= " & ternro & " AND his3.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his3.htethasta IS NULL OR his3.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr3 ON estr3.estrnro = his3.estrnro "
 'Busco el regimen previsional
 sql = sql & " LEFT JOIN his_estructura his4 ON his4.ternro = empleado.ternro AND his4.tenro = 15 AND empleado.ternro= " & ternro & " AND his4.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his4.htethasta IS NULL OR his4.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr4 ON estr4.estrnro = his4.estrnro "
 'Busco la sucursal
 sql = sql & " LEFT JOIN his_estructura his5 ON his5.ternro = empleado.ternro AND his5.tenro = 1 AND empleado.ternro= " & ternro & " AND his5.htetdesde <= " & ConvFecha(pliqhasta) & " AND (his5.htethasta IS NULL OR his5.htethasta >= " & ConvFecha(pliqhasta) & ") "
 sql = sql & " LEFT JOIN estructura estr5 ON estr5.estrnro = his5.estrnro "
    
 sql = sql & " WHERE empleado.ternro = " & ternro

 OpenRecordset sql, rsConsult

 'Fecha Alta
 If IsNull(rsConsult!empfaltagr) Then
    fecalta = ""
 Else
    fecalta = rsConsult!empfaltagr
 End If
 
 'Contrato
 If IsNull(rsConsult!contrato) Then
   contrato = ""
 Else
   contrato = rsConsult!contrato
 End If
 
 'Categoria
 If IsNull(rsConsult!categoria) Then
   categoria = ""
 Else
   categoria = rsConsult!categoria
 End If
 
 'Puesto
 If IsNull(rsConsult!puesto) Then
   puesto = ""
 Else
   puesto = rsConsult!puesto
 End If
 
 'Reg. Prev.
 If IsNull(rsConsult!regprev) Then
   reg_prev = ""
 Else
   reg_prev = rsConsult!regprev
 End If
 
 'Sucursal
 If IsNull(rsConsult!sucursal) Then
   lug_trab = ""
 Else
   lug_trab = rsConsult!sucursal
 End If

 estado = rsConsult!empest
 
 rsConsult.Close
 
 
 ' -------------------------------------------------------------------------
' Busco la fecha de baja
'--------------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de fecha de baja"

StrSql = " SELECT altfec, bajfec "
StrSql = StrSql & " FROM fases "
StrSql = StrSql & " WHERE fases.empleado=" & ternro
StrSql = StrSql & " AND fases.empleado=" & ternro
StrSql = StrSql & " AND altfec <= " & ConvFecha(pliqhasta)
StrSql = StrSql & " AND ( bajfec >= " & ConvFecha(pliqdesde) & " OR bajfec is null)"
'StrSql = StrSql & " AND bajfec <= " & ConvFecha(pliqhasta) & " ) or bajfec is null "
StrSql = StrSql & " ORDER BY altfec DESC"
OpenRecordset StrSql, rsConsult

fecalta = rsConsult!altfec
If rsConsult.EOF Then
    fecbaja = ""
Else
    If Not IsNull(rsConsult!bajfec) Then
        If rsConsult!bajfec <= pliqhasta Then
            fecbaja = rsConsult!bajfec
        Else
            fecbaja = ""
        End If
    Else
        fecbaja = ""
    End If
End If
If rsConsult.State = adStateOpen Then rsConsult.Close
 
'------------------------------------------------------------------
'Busco el valor del centro de costo
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos del centro de costo"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(pliqhasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(pliqhasta) & ") AND his_estructura.tenro=5 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult

centroCosto = ""

If Not rsConsult.EOF Then
   centroCosto = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos del centro de costo"
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
'Busco el valor del fecha nac, cuil, doc y tipo doc
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de los documentos"

sql = " SELECT tercero.terfecnac, cuil.nrodoc AS nrocuil, docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc, estcivil.estcivdesabr"
sql = sql & " FROM tercero "
sql = sql & " LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
sql = sql & " LEFT JOIN ter_doc docu ON (docu.ternro= tercero.ternro and docu.tidnro>0 and docu.tidnro<5) "
sql = sql & " LEFT JOIN tipodocu ON tipodocu.tidnro= docu.tidnro "
sql = sql & " LEFT JOIN estcivil ON estcivil.estcivnro= tercero.estcivnro "
sql = sql & " WHERE tercero.ternro= " & ternro

OpenRecordset sql, rsConsult

Cuil = ""
documento = ""
fecha_nac = ""
est_civil = 0

If Not rsConsult.EOF Then
   Cuil = rsConsult!nrocuil
   documento = rsConsult!sigladoc & "-" & rsConsult!nrodoc
   fecha_nac = rsConsult!terfecnac
   est_civil = rsConsult!estcivdesabr
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la obra social
'------------------------------------------------------------------

'---LOG---
Flog.Writeline "Buscando datos de la obra social"

StrSql = " SELECT estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND his_estructura.tenro=17 AND his_estructura.ternro=" & ternro
       
OpenRecordset StrSql, rsConsult

ObraSocial = ""

If Not rsConsult.EOF Then
   ObraSocial = rsConsult!estrdabr
Else
'   Flog.writeline "Error al obtener los datos de la obra social"
'   GoTo MError
End If

rsConsult.Close


'------------------------------------------------------------------
'Busco los datos de los acumuladores
'------------------------------------------------------------------

'Basico
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_basico
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del basico"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   basico = 0
Else
   basico = rsConsult!almonto
End If

rsConsult.Close

'Neto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_neto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del neto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   neto = 0
Else
   neto = rsConsult!almonto
End If

rsConsult.Close

'MSR
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_msr
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del msr"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   msr = 0
Else
   msr = rsConsult!almonto
End If

rsConsult.Close

'Asig. Familia
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_asi_flia
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de la asig. familia"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   asi_flia = 0
Else
   asi_flia = rsConsult!almonto
End If

rsConsult.Close

'Descuentos
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_Dtos
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos de los descuentos"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   dtos = 0
Else
   dtos = rsConsult!almonto
End If

rsConsult.Close

'Bruto
sql = " SELECT almonto,acunro "
sql = sql & " FROM acu_liq"
sql = sql & " WHERE acunro = " & acum_bruto
sql = sql & " AND cliqnro =  " & cliqnro

'---LOG---
Flog.Writeline "Buscando datos del bruto"

OpenRecordset sql, rsConsult

If rsConsult.EOF Then
   bruto = 0
Else
   bruto = rsConsult!almonto
End If

rsConsult.Close

'------------------------------------------------------------------
'Busco el valor de la direccion
'------------------------------------------------------------------

direccion = ""

Select Case zonaDomicilio
   'Direccion de la sucursal
   Case 1

        sql = " SELECT detdom.calle,detdom.nro,localidad.locdesc"
        sql = sql & " From his_estructura"
        sql = sql & " INNER JOIN sucursal ON sucursal.estrnro=his_estructura.estrnro AND htetdesde <= " & ConvFecha(fecpago) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fecpago) & ") AND tenro=1 AND his_estructura.ternro=" & ternro
        sql = sql & " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro"
        sql = sql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        sql = sql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
    
        '---LOG---
        Flog.Writeline "Buscando datos de la direccion de la sucursal"
        
        OpenRecordset sql, rsConsult
        
        If Not rsConsult.EOF Then
           direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
        End If
        
        rsConsult.Close
    
  'Direccion de la empresa
  Case 2

        sql = "SELECT his_estructura.estrnro, empresa.ternro, empresa.empnom " & _
            " From his_estructura" & _
            " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro" & _
            " WHERE his_estructura.htetdesde <=" & ConvFecha(fecEstr) & " AND " & _
            " (his_estructura.htethasta >= " & ConvFecha(fecEstr) & " OR his_estructura.htethasta IS NULL)" & _
            " AND his_estructura.ternro = " & ternro & _
            " AND his_estructura.tenro  = 10"
        
        OpenRecordset sql, rsConsult
        
        EmpTernro = 0
        
        If Not rsConsult.EOF Then
            EmpTernro = rsConsult!ternro
        End If
        
        rsConsult.Close
        
        'Consulta para obtener la direccion de la empresa
        sql = "SELECT detdom.calle,detdom.nro,localidad.locdesc From cabdom " & _
            " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & EmpTernro & _
            " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
        
        '---LOG---
        Flog.Writeline "Buscando datos de la direccion de la empresa"
        
        OpenRecordset sql, rsConsult
        
        If Not rsConsult.EOF Then
           direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
        End If
        
        rsConsult.Close
    
  'Direccion del empleado
  Case 3

        sql = " SELECT detdom.calle,detdom.nro,localidad.locdesc,detdom.piso,detdom.oficdepto,detdom.barrio"
        sql = sql & " FROM  cabdom "
        sql = sql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro"
        sql = sql & " INNER JOIN localidad ON detdom.locnro = localidad.locnro"
        sql = sql & " WHERE  cabdom.domdefault = -1 AND cabdom.ternro = " & ternro
       
        '---LOG---
        Flog.Writeline "Buscando datos de la direccion del empleado"
        
        OpenRecordset sql, rsConsult
        
        If Not rsConsult.EOF Then
           direccion = rsConsult!calle & " " & rsConsult!Nro & ", " & rsConsult!locdesc
        End If
        
        rsConsult.Close
    
End Select


'------------------------------------------------------------------
'Busco los datos de la cargas sociales
'------------------------------------------------------------------
StrSql = " SELECT * FROM peri_ccss "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro = " & pliqnro
StrSql = StrSql & " AND estrnro = " & emprNro

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqfecdep = rsConsult!Fecha
   pliqbco = rsConsult!banco
Else
   pliqfecdep = ""
   pliqbco = ""
   Flog.Writeline "No se encontraron los datos de las cargas sociales"
'   GoTo MError
End If

rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos
'------------------------------------------------------------------

If IsNull(pliqfecdep) Then
   pliqfecdep = "NULL"
Else
   If Trim(pliqfecdep) = "" Then
      pliqfecdep = "NULL"
   Else
      pliqfecdep = ConvFecha(pliqfecdep)
   End If
End If

StrSql = " INSERT INTO rep_libroley "
StrSql = StrSql & " (bpronro , Legajo, ternro, pronro,"
StrSql = StrSql & " apellido , apellido2, nombre, nombre2,"
StrSql = StrSql & " empresa , emprnro, pliqnro, fecpago,"
StrSql = StrSql & " fecalta , fecbaja, contrato, categoria,"
StrSql = StrSql & " direccion , puesto, documento, fecha_nac,"
StrSql = StrSql & " est_civil , cuil, estado, reg_prev,"
StrSql = StrSql & " lug_trab , basico, neto, msr,"
StrSql = StrSql & " asi_flia , dtos, bruto, "
StrSql = StrSql & " prodesc , descripcion, pliqdesc, pliqmes, "
StrSql = StrSql & " pliqanio , profecpago, pliqfecdep, pliqbco,ultima_pag_impr, auxchar1, auxchar2, auxchar3, "
StrSql = StrSql & " estrdabr1,estrdabr2,estrdabr3,tedabr1,tedabr2,tedabr3,orden,auxchar4, auxchar5) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & "," & Legajo
StrSql = StrSql & "," & ternro
StrSql = StrSql & "," & proNro
StrSql = StrSql & ",'" & apellido & "'"
StrSql = StrSql & ",'" & apellido2 & "'"
StrSql = StrSql & ",'" & nombre & "'"
StrSql = StrSql & ",'" & nombre2 & "'"
StrSql = StrSql & ",'" & empresa & "'"
StrSql = StrSql & "," & emprNro
StrSql = StrSql & "," & pliqnro
StrSql = StrSql & ",'" & fecpago & "'"
StrSql = StrSql & ",'" & fecalta & "'"
StrSql = StrSql & ",'" & fecbaja & "'"
StrSql = StrSql & ",'" & contrato & "'"
StrSql = StrSql & ",'" & categoria & "'"
StrSql = StrSql & ",'" & Mid(direccion, 1, 50) & "'"
StrSql = StrSql & ",'" & puesto & "'"
StrSql = StrSql & ",'" & documento & "'"
StrSql = StrSql & ",'" & fecha_nac & "'"
StrSql = StrSql & ",'" & est_civil & "'"
StrSql = StrSql & ",'" & Cuil & "'"
StrSql = StrSql & "," & estado
StrSql = StrSql & ",'" & reg_prev & "'"
StrSql = StrSql & ",'" & lug_trab & "'"
StrSql = StrSql & "," & numberForSQL(basico)
StrSql = StrSql & "," & numberForSQL(neto)
StrSql = StrSql & "," & numberForSQL(msr)
StrSql = StrSql & "," & numberForSQL(asi_flia)
StrSql = StrSql & "," & numberForSQL(dtos)
StrSql = StrSql & "," & numberForSQL(bruto)
StrSql = StrSql & ",'" & prodesc & "'"
StrSql = StrSql & ",'" & Mid(descripcion, 1, 100) & "'"
StrSql = StrSql & ",'" & pliqdesc & "'"
StrSql = StrSql & "," & pliqmes
StrSql = StrSql & "," & pliqanio
StrSql = StrSql & "," & ConvFecha(profecpago)
StrSql = StrSql & "," & pliqfecdep
StrSql = StrSql & ",'" & pliqbco & "'"
StrSql = StrSql & "," & Pagina
StrSql = StrSql & ",'" & emprCuit & "'"
StrSql = StrSql & ",'" & emprDire & "'"
StrSql = StrSql & ",'" & emprActiv & "'"
StrSql = StrSql & "," & controlNull(estrnomb1)
StrSql = StrSql & "," & controlNull(estrnomb2)
StrSql = StrSql & "," & controlNull(estrnomb3)
StrSql = StrSql & "," & controlNull(tenomb1)
StrSql = StrSql & "," & controlNull(tenomb2)
StrSql = StrSql & "," & controlNull(tenomb3)
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & centroCosto & "'"
StrSql = StrSql & ",'" & ObraSocial & "'"
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Busco el detalle de la liquidacion
'------------------------------------------------------------------

StrSql = " SELECT concepto.concabr, concepto.conccod, concepto.concnro, concepto.concimp, detliq.dlicant, detliq.dlimonto " & _
    " FROM cabliq " & _
    " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND proceso.pronro = " & proNro & _
    " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro " & _
    " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & ternro & _
    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 " & _
    " ORDER BY periodo.pliqnro,cabliq.pronro, concepto.conccod "

'---LOG---
Flog.Writeline "Buscando datos del detalle de liquidacion"

OpenRecordset StrSql, rsConsult

Do Until rsConsult.EOF

  StrSql = " INSERT INTO rep_libroley_det "
  StrSql = StrSql & " (bpronro, ternro, pronro, concabr,"
  StrSql = StrSql & " conccod ,concnro , concimp , dlicant,"
  StrSql = StrSql & " dlimonto) "
  StrSql = StrSql & " VALUES "
  StrSql = StrSql & "(" & NroProceso
  StrSql = StrSql & "," & ternro
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

'------------------------------------------------------------------
'Busco los datos de los familiares
'------------------------------------------------------------------

If (tipoFamiliares = "1") Or (tipoFamiliares = "2") Then

    StrSql = " SELECT docu.nrodoc AS nrodoc, tipodocu.tidsigla AS sigladoc,tercero.ternro,tercero.terape, tercero.ternom, tercero.terfecnac, tercero.tersex, famest,famtrab,faminc,famDGIdesde " & _
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
    
    Do Until rsConsult.EOF
    
      GuardarFam = True
    
      If (tipoFamiliares = "2") Then
          GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg, cliqnro, concFamiliar01)
          If Not GuardarFam Then
             GuardarFam = HayAsignacionesFliares(pliqdesde, pliqhasta, ternro, rsConsult!ternro, NroBusqProg2, cliqnro, concFamiliar02)
          End If
      End If
    
      If GuardarFam Then
      
          If IsNull(rsConsult!terfecnac) Then
             terfecnac = "NULL"
          Else
             terfecnac = ConvFecha(rsConsult!terfecnac)
          End If
        
          If IsNull(rsConsult!famfecvto) Then
             famfecvto = "NULL"
          Else
             famfecvto = ConvFecha(rsConsult!famfecvto)
          End If
        
          If IsNull(rsConsult!famDGIdesde) Then
             famDGIdesde = "NULL"
          Else
             famDGIdesde = ConvFecha(rsConsult!famDGIdesde)
          End If
          
          If IsNull(rsConsult!famDGIhasta) Then
             famDGIhasta = "NULL"
          Else
             famDGIhasta = ConvFecha(rsConsult!famDGIhasta)
          End If
          
        
          StrSql = " INSERT INTO rep_libroley_fam "
          StrSql = StrSql & " (bpronro, ternro , pronro, nrodoc,"
          StrSql = StrSql & " sigladoc, ternrofam, terape, ternom,"
          StrSql = StrSql & " terfecnac , tersex, famest, famtrab,"
          StrSql = StrSql & " faminc, famsalario, famfecvto, famCargaDGI,"
          StrSql = StrSql & " famDGIdesde , famDGIhasta, famemergencia, paredesc"
          StrSql = StrSql & " ) "
          StrSql = StrSql & " VALUES "
          StrSql = StrSql & "(" & NroProceso
          StrSql = StrSql & "," & ternro
          StrSql = StrSql & "," & proNro
          StrSql = StrSql & ",'" & rsConsult!nrodoc & "'"
          StrSql = StrSql & ",'" & rsConsult!sigladoc & "'"
          StrSql = StrSql & "," & rsConsult!ternro
          StrSql = StrSql & ",'" & rsConsult!terape & "'"
          StrSql = StrSql & ",'" & rsConsult!ternom & "'"
          StrSql = StrSql & "," & terfecnac
          StrSql = StrSql & "," & strForSQL(rsConsult!tersex)
          StrSql = StrSql & "," & strForSQL(rsConsult!famest)
          StrSql = StrSql & "," & strForSQL(rsConsult!famtrab)
          StrSql = StrSql & "," & strForSQL(rsConsult!faminc)
          StrSql = StrSql & "," & strForSQL(rsConsult!famsalario)
          StrSql = StrSql & "," & famfecvto
          StrSql = StrSql & "," & strForSQL(rsConsult!famCargaDGI)
          StrSql = StrSql & "," & famDGIdesde
          StrSql = StrSql & "," & famDGIhasta
          StrSql = StrSql & "," & strForSQL(rsConsult!famemergencia)
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


Sub buscarDatosEmpresaDada(ByVal estrnro As Long, ByVal ternro As Long, ByVal empnom As String, empactiv)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

    empresa = ""
    emprNro = 0
    emprActiv = ""
    emprTer = 0
    emprCuit = ""
    emprDire = ""

    '---LOG---
    Flog.Writeline "Buscando datos de la empresa"
    
    If estrnro = 0 Then
        Flog.Writeline "No se encontró la empresa"
        Exit Sub
    Else
        empresa = empnom
        emprNro = estrnro
        emprActiv = empactiv
        emprTer = ternro
    End If
    
    'Consulta para obtener la direccion de la empresa
    StrSql = "SELECT detdom.calle,detdom.nro,localidad.locdesc From cabdom " & _
        " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & emprTer & _
        " INNER JOIN localidad ON detdom.locnro = localidad.locnro "
    
    '---LOG---
    Flog.Writeline "Buscando datos de la direccion de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.Writeline "No se encontró el domicilio de la empresa"
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
    Flog.Writeline "Buscando datos del cuit de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.Writeline "No se encontró el CUIT de la Empresa"
        'Exit Sub
        emprCuit = "  "
    Else
        emprCuit = rsConsult!nrodoc
    End If
    
    rsConsult.Close

End Sub

