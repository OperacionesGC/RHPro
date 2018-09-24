Attribute VB_Name = "repHorasSup"

Global Const Version = "1.06" 'Miriam Ruiz
Global Const FechaModificacion = "11/12/2015"
Global Const UltimaModificacion = "CAS-33186 - GESTIÓN COMPARTIDA -  Libro de Horas Extras CABA"
'               Se calcula el valor de la horas extras con todos los decimales y se redondea al final



'Global Const Version = "1.05" 'Miriam Ruiz
'Global Const FechaModificacion = "02/09/2015"
'Global Const UltimaModificacion = "CAS-30722 - NGA - Bug en reporte de horas suplementarias"
'               Se controla que la actividad de la empresa no sea null


'Global Const Version = "1.04" 'Miriam Ruiz
'Global Const FechaModificacion = "29/06/2015"
'Global Const UltimaModificacion = "CAS-30572 - Xerox - CUSTOM  LIBRO DE HORAS SUPLEMENTARIAS"
'               Se modificó el proceso para clientes con gti R3


'Global Const Version = "1.03" 'Miriam Ruiz
'Global Const FechaModificacion = "05/06/2015"
'Global Const UltimaModificacion = "CAS-30722 - RH Pro - Libro Registro de Horas Suplementarias"
'               Se corrigió el cambio de día

'Global Const Version = "1.02" 'Miriam Ruiz
'Global Const FechaModificacion = "26/05/2015"
'Global Const UltimaModificacion = "CAS-30722 - RH Pro - Libro Registro de Horas Suplementarias"
'               Se corrigió formato de hora

'Global Const Version = "1.01" 'Miriam Ruiz'
'Global Const FechaModificacion = "15/05/2015"
'Global Const UltimaModificacion = "CAS-30722 - RH Pro - Libro Registro de Horas Suplementarias"
'               Se agregó el procesamiento para clientes que no tienen gti

'Global Const Version = "1.00" 'Miriam Ruiz
'Global Const FechaModificacion = "20/03/2015"
'Global Const UltimaModificacion = "CAS-29425 - MEDICUS - CUSTOM  LIBRO DE HORAS SUPLEMENTARIAS"

'--------------------------------------------------------------

'--------------------------------------------------------------
Option Explicit

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

Global Pagina As Long
Global tomo As Integer
Global tope As Integer

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String

Global empresa
Global emprActiv As String
Global emprCuit As String
Global emprDire As String
Global emprNom As String

Global listapronro
Global l_orden
Global filtro
Global totalEmpleados
Global cantRegistros

Global Sueldo As String
Global Tipo_Sueldo As String ' ac-co-acc-coc
Global Total_Horas As String ' ac-co-acc-coc
Global Tipo_Horas As String ' ac-co-acc-coc
Global usa_gti As String
Global tipos_hs_extras As String
Global lista_hs_extras As String
Global lista_conceptos As String
Global objFechasHoras As New FechasHoras
Global horas_valor As Integer
Global Escala As Integer
Global TFL As Integer ' tipo estructura de la forma liquidación
Global Formaliquidacion As String
Global rsConsult2 As New ADODB.Recordset
Global empresascongti As String



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
Dim objRs2 As New ADODB.Recordset
Dim objRs3 As New ADODB.Recordset
Dim fechadesde
Dim fechahasta

Dim param
Dim Pronro
Dim Ternro As Long

Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim rsAge As New ADODB.Recordset
Dim rsEmpresas As New ADODB.Recordset
Dim rsPeriodo As New ADODB.Recordset
Dim acunroSueldo
Dim I
Dim PID As String
Dim tituloReporte As String

Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden
Dim ord
    
Dim arrpliqnro
Dim listapliqnro
Dim pliqnro
Dim auxempr As String
    
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
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteHorasSup" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Horas Suplementarias : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    TiempoInicialProceso = GetTickCount
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo CE
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       
'2)
       
       
   'Obtengo los parametros del proceso

       parametros = objRs!bprcparam
       Flog.writeline " parametros --> " & parametros
       ArrParametros = Split(parametros, "@")

       Flog.writeline " limite del array --> " & UBound(ArrParametros)

       'Obtengo la lista de procesos
       listapronro = ArrParametros(0)

       empresa = ArrParametros(1)
       Call buscarDatosEmpresa(empresa)
       
       'Obtengo los cortes de estructura
       tenro1 = CInt(ArrParametros(2))
       estrnro1 = CInt(ArrParametros(3))

       tenro2 = CInt(ArrParametros(4))
       estrnro2 = CInt(ArrParametros(5))

       tenro3 = CInt(ArrParametros(6))
       estrnro3 = CInt(ArrParametros(7))

       If UBound(ArrParametros) > 7 Then
        fecEstr = ArrParametros(8)
       End If

       'Obtengo el titulo del reporte
       If UBound(ArrParametros) > 8 Then
        tituloReporte = ArrParametros(9)
       Else
        tituloReporte = ""
       End If
       'pagina
        If UBound(ArrParametros) > 9 Then
            Pagina = ArrParametros(10)
            
        Else
            Pagina = 0
        End If
       'tomo
        If UBound(ArrParametros) > 10 Then
           tomo = ArrParametros(11)
          Else
            tomo = 0
        End If
       'tope
       If UBound(ArrParametros) > 11 Then
           tope = ArrParametros(12)
        Else
            tope = 0
        End If
      
       
       'EMPIEZA EL PROCESO
   
       'Busco en el confrep
       StrSql = " SELECT * FROM confrepAdv "
       StrSql = StrSql & " WHERE repnro = 472 "
      
       OpenRecordset StrSql, objRs2
       
        lista_hs_extras = "0"
        empresascongti = "0"
        If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep"
          Exit Sub
        End If
       
        Flog.writeline "Obtengo los datos del confrep"
       
             tipos_hs_extras = "0@0@0"
             lista_conceptos = "0"
             
       Do Until objRs2.EOF
       
          Select Case objRs2!confnrocol
             Case 2
                usa_gti = objRs2!confval
             Case 3
                Sueldo = objRs2!confval
                Escala = objRs2!confval2
                
                Tipo_Sueldo = objRs2!conftipo

             Case 4
                
                If objRs2!confval <> 0 Then
                    tipos_hs_extras = tipos_hs_extras & "%" & objRs2!confval & "@," & objRs2!confval2 & ",@" & objRs2!confval3
                    lista_hs_extras = lista_hs_extras & "," & objRs2!confval2
                    lista_conceptos = lista_conceptos & "," & objRs2!confval4
                End If
                
             Case 5
                
                If objRs2!confval <> 0 Then
                    tipos_hs_extras = tipos_hs_extras & "%" & objRs2!confval & "@," & objRs2!confval2 & ",@" & objRs2!confval3
                    lista_hs_extras = lista_hs_extras & "," & objRs2!confval2
                    lista_conceptos = lista_conceptos & "," & objRs2!confval4
                End If
             Case 6
                
                If objRs2!confval <> 0 Then
                    tipos_hs_extras = tipos_hs_extras & "%" & objRs2!confval & "@," & objRs2!confval2 & ",@" & objRs2!confval3
                    lista_hs_extras = lista_hs_extras & "," & objRs2!confval2
                    lista_conceptos = lista_conceptos & "," & objRs2!confval4
                End If
             Case 7
                
                If objRs2!confval <> 0 Then
                    tipos_hs_extras = tipos_hs_extras & "%" & objRs2!confval & "@," & objRs2!confval2 & ",@" & objRs2!confval3
                    lista_hs_extras = lista_hs_extras & "," & objRs2!confval2
                    lista_conceptos = lista_conceptos & "," & objRs2!confval4
                End If
             Case 8
                
                If objRs2!confval <> 0 Then
                     tipos_hs_extras = tipos_hs_extras & "%" & objRs2!confval & "@," & objRs2!confval2 & ",@" & objRs2!confval3
                     lista_hs_extras = lista_hs_extras & "," & objRs2!confval2
                     lista_conceptos = lista_conceptos & "," & objRs2!confval4
                End If
              Case 9
                
                If objRs2!confval <> 0 Then
                     TFL = objRs2!confval
                End If
              Case 99
                If objRs2!confval <> "" Then
                    empresascongti = empresascongti & "," & objRs2!confval
                End If
          End Select

          objRs2.MoveNext
       Loop

        If empresascongti <> "0" Then
            auxempr = "," & empresa
            If InStr(empresascongti, auxempr) > 0 Then
                usa_gti = -1
            Else
                usa_gti = 0
            End If
        
        End If
            Flog.writeline "usa_gti:" & usa_gti
       'Obtengo los empleados sobre los que tengo que generar las horas suplementarias
      
           CargarEmpleados NroProceso, rsEmpl, empresa
      
                ' Actualizo el estado del proceso
               
                      StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
            
                   objConn.Execute StrSql, , adExecuteNoRecords
                             
                    orden = 0
                    'Genero por cada empleado un registro
                    Do Until rsEmpl.EOF
                       EmpErrores = False
                       Ternro = rsEmpl!Ternro
                       'orden = orden + 1
                       'Genero una entrada para el empleado por cada proceso
                          Pronro = listapronro
                          Flog.writeline "Generando datos empleado " & Ternro & " para el proceso " & Pronro
                          
                           Call generarDatosEmpleado01(Pronro, Ternro, tituloReporte, orden)
                                              
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
                       
       End If
    

   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
End Sub

Function controlNull(Str)
  If Trim(Str) = "" Then
     controlNull = "null"
  Else
     controlNull = "'" & Str & "'"
  End If
End Function



'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal empresa As Long)

Dim StrEmpl As String

    If NroProc > 0 Then
        
        StrEmpl = "SELECT * FROM batch_empleado "
        StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
        StrEmpl = StrEmpl & " AND (agencia is null OR agencia = 0)"
        
        StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
        StrEmpl = StrEmpl & " ORDER BY progreso,estado"
    
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


Sub CalcularhorarioES(ByVal Ternro As Long, ByVal adfecha As Date, ByRef HoraEnt As String, ByRef HoraSal As String)

'Dim rsHorario As New ADODB.Recordset
'Dim StrSql
'
' StrSql = "  SELECT DISTINCT plahorent, plahorsal FROM planillahorario WHERE ternro =" & Ternro & _
'                                  " AND plahordia = '" & Weekday(adfecha) & "'"
'                        OpenRecordset StrSql, rsHorario
'If Not rsHorario.EOF Then
'    HoraEnt = Left(rsHorario!plahorent, 2) & ":" & Right(rsHorario!plahorent, 2)
'    HoraSal = Left(rsHorario!plahorsal, 2) & ":" & Right(rsHorario!plahorsal, 2)
'Else
'    HoraEnt = ""
'    HoraSal = ""
'End If
'rsHorario.Close

Dim rsHorario As New ADODB.Recordset
'Dim StrSql
'se cambia la forma que busca el horario
 If usa_gti = 0 Then
     StrSql = "  SELECT DISTINCT plahorent, plahorsal FROM planillahorario WHERE ternro =" & Ternro & _
              " AND plahordia = '" & Weekday(adfecha) & "'"
     Flog.writeline " Horario Planilla " & StrSql
     OpenRecordset StrSql, rsHorario
     If Not rsHorario.EOF Then
            HoraEnt = Left(rsHorario!plahorent, 2) & ":" & Right(rsHorario!plahorent, 2)
            HoraSal = Left(rsHorario!plahorsal, 2) & ":" & Right(rsHorario!plahorsal, 2)
     Else
            HoraEnt = ""
            HoraSal = ""
     End If
 Else
   StrSql = "  SELECT diahoradesde1,diahorahasta1,diahoradesde2,diahorahasta2, "
   StrSql = StrSql + "     diahoradesde3 , diahorahasta3 "
   StrSql = StrSql + " FROM  gti_proc_emp "
   StrSql = StrSql + " INNER JOIN gti_dias ON gti_proc_emp.dianro = gti_dias.dianro "
   StrSql = StrSql + " WHERE Ternro = " & Ternro
   StrSql = StrSql + " AND diacanthoras >0 "
   StrSql = StrSql + " AND gti_proc_emp.fecha = '" & adfecha & "'"
   Flog.writeline " Horario gti " & StrSql
   OpenRecordset StrSql, rsHorario
   
If Not rsHorario.EOF Then
   If rsHorario!diahoradesde3 <> "0000" Then
        HoraEnt = Left(rsHorario!diahoradesde3, 2) & ":" & Right(rsHorario!diahoradesde3, 2)
        HoraSal = Left(rsHorario!diahorahasta3, 2) & ":" & Right(rsHorario!diahorahasta3, 2)
   Else
        If rsHorario!diahoradesde2 <> "0000" Then
             HoraEnt = Left(rsHorario!diahoradesde2, 2) & ":" & Right(rsHorario!diahoradesde2, 2)
             HoraSal = Left(rsHorario!diahorahasta2, 2) & ":" & Right(rsHorario!diahorahasta2, 2)
        Else
             HoraEnt = Left(rsHorario!diahoradesde1, 2) & ":" & Right(rsHorario!diahoradesde1, 2)
             HoraSal = Left(rsHorario!diahorahasta1, 2) & ":" & Right(rsHorario!diahorahasta1, 2)
        End If
    End If
    
Else
    HoraEnt = "08:00"
    HoraSal = "08:00"
End If
End If
 Flog.writeline " HoraEnt  " & HoraEnt
  Flog.writeline " HoraSal  " & HoraSal
rsHorario.Close

End Sub

 Public Sub CalcularhorarioHE(adfecha As Date, HoraEnt As String, HoraSal As String, adcanthoras As String, ByRef HED As String, ByRef HEH As String, ByRef horas As Long, ByRef minutos As Long)    'HED:Hora extra desde - HEH:Hora extra hasta


 Dim fechaSal As Date
 Dim TotDias
 Dim HED_aux As String
 Dim HEH_aux As String
  
  Dim horasalaux As String


    TotDias = 0
    fechaSal = adfecha
    adcanthoras = Left(adcanthoras, 2) & Right(adcanthoras, 2)
        horasalaux = Left(HoraSal, 2) & Right(HoraSal, 2)
    objFechasHoras.SumoHoras adfecha, horasalaux, adcanthoras, fechaSal, HEH

    HED = HoraSal
    HED_aux = Left(HED, 2) & Right(HED, 2)
    HEH_aux = Left(HEH, 2) & Right(HEH, 2)
    objFechasHoras.RestaHs adfecha, HED_aux, fechaSal, HEH_aux, TotDias, horas, minutos

   'minutos = minutos * 60 / 100
   
 End Sub
 
 Public Sub CalcularhorarioHE2(ByVal Ternro As Long, ByVal THmes As Integer, adfecha As Date, adcanthoras As String, ByRef HED As String, ByRef HEH As String, ByRef horas As Long, ByRef minutos As Long)     'HED:Hora extra desde - HEH:Hora extra hasta
 Dim rsHorario As New ADODB.Recordset
 Dim fechaSal As Date
 Dim TotDias
 Dim HED_aux As String
 Dim HEH_aux As String
 
  StrSql = " SELECT horhoradesde FROM gti_horcumplido " & _
           " WHERE ternro =  " & Ternro & _
           " AND " & ConvFecha(adfecha) & ">= hordesde " & _
           " AND " & ConvFecha(adfecha) & "<= horhasta " & _
           " AND thnro = " & THmes & _
           " ORDER BY horhoradesde "
           OpenRecordset StrSql, rsHorario
           
 If Not rsHorario.EOF Then
    fechaSal = adfecha
    adcanthoras = Left(adcanthoras, 2) & Right(adcanthoras, 2)
    objFechasHoras.SumoHoras adfecha, rsHorario!horhoradesde, adcanthoras, fechaSal, HEH

    HED = Left(rsHorario!horhoradesde, 2) & ":" & Right(rsHorario!horhoradesde, 2)
    HED_aux = Left(HED, 2) & Right(HED, 2)
    HEH_aux = Left(HEH, 2) & Right(HEH, 2)
    objFechasHoras.RestaHs adfecha, HED_aux, fechaSal, HEH_aux, TotDias, horas, minutos
 End If
 rsHorario.Close
 
End Sub
 
Sub CalcularPorcentaje(thnro As Integer, ByRef Porcentaje)
Dim listaHE ' lista de hora extras configuradas
Dim enlista
Dim I As Integer
Dim encontre As Boolean


Porcentaje = 0
listaHE = Split(tipos_hs_extras, "%")
I = 0
encontre = False
Do While I <= UBound(listaHE) And Not encontre
    enlista = Split(listaHE(I), "@")
    If InStr(enlista(1), "," & thnro & ",") > 0 Then
        Porcentaje = enlista(2)
        encontre = True
    Else
       I = I + 1
    End If
Loop

End Sub

 Sub CalcularValorHE(valorHoras, Porcentaje, horas, ByRef ValorHE As Double, ByRef TotalValorHE As Double)
  Dim multiplicador
  Dim TotalMinutos
  
    multiplicador = 1 + (Porcentaje / 100)
    'ValorHE = Round(valorHoras * multiplicador, 2)
    ValorHE = valorHoras * multiplicador
  '  TotalMinutos = horas
   ' TotalValorHE = Round(ValorHE * horas, 2)
     TotalValorHE = ValorHE * horas
      Flog.writeline "valorHE : " & TotalValorHE
     
 End Sub
 
Sub buscarDatosEmpresa(emprNro)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim emprTer As Long


    emprActiv = ""
    emprTer = 0
    emprCuit = ""
    emprDire = ""
    ' -------------------------------------------------------------------------
    ' Busco los datos de la empresa
    '--------------------------------------------------------------------------
    
    StrSql = "SELECT empnom,empactiv,empresa.ternro FROM empresa" & _
             " INNER JOIN tercero ON tercero.ternro=empresa.ternro" & _
             " WHERE estrnro = " & emprNro
    
    '---LOG---
    Flog.writeline "Buscando datos de la empresa"
    
    OpenRecordset StrSql, rsConsult
    

    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró la empresa"
        Exit Sub
    Else
        emprNom = rsConsult!empnom
        If IsNull(rsConsult!empactiv) Then
            emprActiv = ""
        Else
            emprActiv = rsConsult!empactiv
        End If
        emprTer = rsConsult!Ternro
    End If
    
    rsConsult.Close
    
    'Consulta para obtener la direccion de la empresa
    StrSql = "SELECT detdom.calle,detdom.nro, detdom.piso, detdom.oficdepto From cabdom " & _
        " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro =" & emprTer & _
        " WHERE domdefault = -1 "
      
    '---LOG---
    Flog.writeline "Buscando datos de la direccion de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el domicilio de la empresa"
        'Exit Sub
        emprDire = "   "
    Else
        emprDire = rsConsult!calle & " " & rsConsult!nro
        '02/10/2006 - Martin Ferraro - Se agrego piso y dpto a la dir del la empresa
        If Not EsNulo(rsConsult!piso) Then
            emprDire = emprDire & " P. " & rsConsult!piso
        End If
        If Not EsNulo(rsConsult!oficdepto) Then
            emprDire = emprDire & " Dpto. " & rsConsult!oficdepto
        End If
        
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
        emprCuit = " "
    Else
        emprCuit = rsConsult!NroDoc
    End If
    
    rsConsult.Close



StrSql = " INSERT INTO repHorasSupEmp "
StrSql = StrSql & " (bpronro , repHsSupEmp_nombre ,repHsSupEmp_dir,repHsSupEmp_act,repHsSupEmp_cuit)"
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & ",'" & emprNom
StrSql = StrSql & "','" & emprDire
StrSql = StrSql & "','" & emprActiv
StrSql = StrSql & "','" & emprCuit
StrSql = StrSql & "')"



'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords


End Sub
 
 Sub buscarjornada(NroCab, horadesde, horahasta)
   Dim Desde
   Dim Hasta
        Desde = Left(horadesde, 2) & ":" & Right(horadesde, 2)
        Hasta = Left(horahasta, 2) & ":" & Right(horahasta, 2)
        StrSql = " SELECT * FROM repHorasSupJor WHERE  "
        StrSql = StrSql & " repHsSupCabnro =  " & NroCab
        StrSql = StrSql & " AND repHsSupJor_Desde = '" & Desde & "'"
        StrSql = StrSql & " AND repHsSupJor_Hasta = '" & Hasta & "'"
        OpenRecordset StrSql, rsConsult2

End Sub

 
'--------------------------------------------------------------------
' Se encarga de generar los datos para el Standard
'--------------------------------------------------------------------
Sub generarDatosEmpleado01(Pronro, Ternro, descripcion, orden)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset
Dim rsAcumdiario As New ADODB.Recordset


'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim Apellido As String
Dim apellido2 As String
Dim Nombre As String
Dim nombre2 As String
Dim pliqnro
Dim pliqmes
Dim pliqanio
Dim aux As String
Dim aux2 As String
Dim posdec
Dim decaux As String
Dim cuil
Dim pliqdesc As String
 Dim fechaSalaux As Date
 Dim HEHaux As String
Dim Sueldo_valor As String


Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim tenomb1
Dim tenomb2
Dim tenomb3
Dim valorHoras
Dim sql As String
Dim cliqnro
Dim NroCab
Dim CHmes
Dim THmes As Integer
Dim cambiohora
Dim FechaCompleta As String
Dim HoraEnt As String
Dim HoraSal As String
Dim horas As Long

Dim minutos As Long
Dim Porcentaje
Dim ValorHE As Double
Dim TotalValorHE As Double
Dim HED As String ' Hora Extra Desde
Dim HEH As String ' Hora Extra Hasta

On Error GoTo MError

estrnomb1 = ""
estrnomb2 = ""
estrnomb3 = ""
tenomb1 = ""
tenomb2 = ""
tenomb3 = ""


Flog.writeline Ternro
'usa_gti = -1
'------------------------------------------------------------------
'Busco los datos del empleado
'------------------------------------------------------------------
StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empremu "
StrSql = StrSql & " FROM empleado "
StrSql = StrSql & " WHERE ternro= " & Ternro

Flog.writeline "Buscando datos del empleado"
       
OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   Nombre = rsConsult!ternom
   If IsNull(rsConsult!ternom2) Then
      nombre2 = ""
   Else
      nombre2 = rsConsult!ternom2
   End If
   Apellido = rsConsult!terape
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
'Busco los datos del periodo actual
'------------------------------------------------------------------
StrSql = " SELECT periodo.pliqnro, pliqdesc, cabliq.cliqnro, pliqdesde,pliqhasta FROM periodo "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " AND proceso.pronro IN (" & Pronro & ")"
StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " AND cabliq.empleado= " & Ternro

'---LOG---
Flog.writeline "Buscando datos del periodo"

OpenRecordset StrSql, rsConsult

If Not rsConsult.EOF Then
   pliqnro = rsConsult!pliqnro
   pliqdesc = rsConsult!pliqdesc
   cliqnro = rsConsult!cliqnro
   pliqdesde = rsConsult!pliqdesde
   pliqhasta = rsConsult!pliqhasta
Else
   Flog.writeline "El empleado no se encuentra en el proceso"
   Exit Sub
End If

rsConsult.Close
   
    
'------------------------------------------------------------------
'Busco el valor del cuil
'------------------------------------------------------------------

'---LOG---
Flog.writeline "Buscando datos de los documentos"

sql = " SELECT cuil.nrodoc AS nrocuil"
sql = sql & " FROM tercero "
sql = sql & " LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
sql = sql & " WHERE tercero.ternro= " & Ternro

OpenRecordset sql, rsConsult

cuil = ""

If Not rsConsult.EOF Then
  If Not IsNull(rsConsult!nrocuil) Then
   cuil = Replace(rsConsult!nrocuil, "-", "")
  End If
End If

rsConsult.Close

'-------------------------------------------------------------------
' Busco la cantidad de horas trabajadas en el mes
'--------------------------------------------------------------------

sql = " SELECT vgrvalor FROM his_estructura "
sql = sql & " INNER JOIN valgrilla ON his_estructura.estrnro = vgrcoor_2 "
sql = sql & "  AND cgrnro = " & Escala
sql = sql & " AND vgrorden = 1 "
sql = sql & " WHERE "
sql = sql & " tenro = 19 AND "
sql = sql & " Ternro = " & Ternro
sql = sql & " AND htetdesde <= " & ConvFecha(pliqdesde)
sql = sql & " AND (htethasta >= " & ConvFecha(pliqdesde)
sql = sql & " OR htethasta IS null)"
OpenRecordset sql, rsConsult
If Not rsConsult.EOF Then
    horas_valor = rsConsult("vgrvalor")
Else
    horas_valor = 160
    Flog.writeline "No se encontró escala para el convenio, se tomará como base 160 hs mensuales"
End If
rsConsult.Close
'------------------------------------------------------------------
'Busco los datos de los acumuladores o conceptos
'------------------------------------------------------------------

'Basico
If Tipo_Sueldo = "AC" Or Tipo_Sueldo = "ACC" Then
    sql = " SELECT almonto, alcant ,acunro "
    sql = sql & " FROM acu_liq"
    sql = sql & " WHERE acunro IN (" & Sueldo & ")"
    sql = sql & " AND cliqnro =  " & cliqnro
Else
    sql = " SELECT detliq.dlimonto almonto ,detliq.dlicant alcant "
    sql = sql & " FROM cabliq "
    sql = sql & " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND cabliq.pronro IN (" & Pronro & ")"
    sql = sql & " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro "
    sql = sql & " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado = " & Ternro & " AND detliq.concnro IN (" & Sueldo & ")"
End If

'---LOG---
Flog.writeline "Buscando datos del Sueldo"

OpenRecordset sql, rsConsult

Sueldo_valor = 0
Do While Not rsConsult.EOF
         If Tipo_Sueldo = "AC" Or Tipo_Sueldo = "CO" Then
              Sueldo_valor = Sueldo_valor + rsConsult!almonto
        Else
              Sueldo_valor = Sueldo_valor + rsConsult!alcant
        End If
    rsConsult.MoveNext
  Loop
rsConsult.Close

If horas_valor > 0 Then
  'valorHoras = Round(Sueldo_valor / horas_valor, 2)
  valorHoras = Sueldo_valor / horas_valor
Else
 valorHoras = 0
End If

sql = " SELECT estructura.estrdabr  FROM his_estructura "
sql = sql & " INNER JOIN estructura ON  his_estructura.estrnro =  estructura.estrnro"
sql = sql & " WHERE his_estructura.Tenro = " & TFL
sql = sql & " AND ternro = " & Ternro
sql = sql & " AND htetdesde <= " & ConvFecha(pliqdesde)
sql = sql & " AND (htethasta >= " & ConvFecha(pliqdesde)
sql = sql & " OR htethasta IS null)"
OpenRecordset sql, rsConsult
If Not rsConsult.EOF Then
    Formaliquidacion = Left(rsConsult!estrdabr, 20)
Else
    Formaliquidacion = ""
End If
rsConsult.Close

'------------------------------------------------------------------
'Armo la SQL para guardar los datos de la cabecera
'------------------------------------------------------------------
orden = orden + 1
StrSql = " INSERT INTO repHorasSupCab "
StrSql = StrSql & " (bpronro , repHsSupCab_periodo, ternro,"
StrSql = StrSql & " repHsSupCab_Ape_Nom, repHsSupCab_leg,"
StrSql = StrSql & " repHsSupCab_cuil,"
StrSql = StrSql & " repHsSupCab_sueldo,"
StrSql = StrSql & " repHsSupCab_horas,"
StrSql = StrSql & " repHsSupCab_valor,repHsSupCab_orden,repHsSupCab_FL, "
StrSql = StrSql & " repHsSupCab_pag,repHsSupCab_tope,repHsSupCab_tomo) "
StrSql = StrSql & " VALUES "
StrSql = StrSql & "(" & NroProceso
StrSql = StrSql & ",'" & pliqdesc
StrSql = StrSql & "'," & Ternro
StrSql = StrSql & ",'" & Apellido & " " & apellido2 & " " & Nombre & " " & nombre2 & "'"
StrSql = StrSql & "," & Legajo
StrSql = StrSql & ",'" & cuil & "'"
StrSql = StrSql & "," & Sueldo_valor
StrSql = StrSql & "," & horas_valor
StrSql = StrSql & "," & valorHoras
StrSql = StrSql & "," & orden
StrSql = StrSql & ",'" & Formaliquidacion
StrSql = StrSql & "'," & Pagina
StrSql = StrSql & "," & tope
StrSql = StrSql & "," & tomo
StrSql = StrSql & ")"

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

objConn.Execute StrSql, , adExecuteNoRecords

'------------------------------------------------------------------
'Guardo los datos en la BD
'------------------------------------------------------------------

 StrSql = "Select max(repHsSupCabnro) NroCab from repHorasSupCab "
 OpenRecordset StrSql, rsConsult
 NroCab = rsConsult!NroCab
 rsConsult.Close

 'Jornada laboral

If usa_gti = 0 Then

        StrSql = "  SELECT DISTINCT plahordia,plahorent, plahorsal FROM planillahorario WHERE ternro =" & Ternro & _
                 " ORDER BY plahordia "
        OpenRecordset StrSql, rsConsult
        
        Do While Not rsConsult.EOF
            StrSql = " INSERT INTO repHorasSupJor "
            StrSql = StrSql & " (repHsSupCabnro, repHsSupJor_Desde,repHsSupJor_Hasta)"
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "(" & NroCab
            StrSql = StrSql & ",'" & Left(rsConsult!plahorent, 2) & ":" & Right(rsConsult!plahorent, 2)
            StrSql = StrSql & "','" & Left(rsConsult!plahorsal, 2) & ":" & Right(rsConsult!plahorsal, 2)
            StrSql = StrSql & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
            rsConsult.MoveNext
        Loop
        rsConsult.Close

Else
   StrSql = "  SELECT distinct diahoradesde1,diahorahasta1,diahoradesde2,diahorahasta2, "
   StrSql = StrSql + "     diahoradesde3 , diahorahasta3 "
   StrSql = StrSql + " FROM  gti_proc_emp "
   StrSql = StrSql + " INNER JOIN gti_dias ON gti_proc_emp.dianro = gti_dias.dianro "
   StrSql = StrSql + " WHERE Ternro = " & Ternro
   StrSql = StrSql + " AND diacanthoras >0 "
   StrSql = StrSql + " AND gti_proc_emp.fecha >= " & ConvFecha(pliqdesde)
   StrSql = StrSql + " AND gti_proc_emp.fecha <= " & ConvFecha(pliqhasta)
   OpenRecordset StrSql, rsConsult

    Do While Not rsConsult.EOF
    
        Call buscarjornada(NroCab, rsConsult!diahoradesde1, rsConsult!diahorahasta1)
        
        If rsConsult2.EOF Then
            StrSql = " INSERT INTO repHorasSupJor "
            StrSql = StrSql & " (repHsSupCabnro, repHsSupJor_Desde,repHsSupJor_Hasta)"
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "(" & NroCab
            StrSql = StrSql & ",'" & Left(rsConsult!diahoradesde1, 2) & ":" & Right(rsConsult!diahoradesde1, 2)
            StrSql = StrSql & "','" & Left(rsConsult!diahorahasta1, 2) & ":" & Right(rsConsult!diahorahasta1, 2)
            StrSql = StrSql & "')"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rsConsult2.Close
        If rsConsult!diahoradesde2 <> "0000" Then
              Call buscarjornada(NroCab, rsConsult!diahoradesde2, rsConsult!diahorahasta2)
              If rsConsult2.EOF Then
                    StrSql = " INSERT INTO repHorasSupJor "
                    StrSql = StrSql & " (repHsSupCabnro, repHsSupJor_Desde,repHsSupJor_Hasta)"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroCab
                    StrSql = StrSql & ",'" & Left(rsConsult!diahoradesde2, 2) & ":" & Right(rsConsult!diahoradesde2, 2)
                    StrSql = StrSql & "','" & Left(rsConsult!diahorahasta2, 2) & ":" & Right(rsConsult!diahorahasta2, 2)
                    StrSql = StrSql & "')"
                    objConn.Execute StrSql, , adExecuteNoRecords
               End If
               rsConsult2.Close
        End If
        If rsConsult!diahoradesde3 <> "0000" Then
              Call buscarjornada(NroCab, rsConsult!diahoradesde3, rsConsult!diahorahasta3)
              If rsConsult2.EOF Then
                    StrSql = " INSERT INTO repHorasSupJor "
                    StrSql = StrSql & " (repHsSupCabnro, repHsSupJor_Desde,repHsSupJor_Hasta)"
                    StrSql = StrSql & " VALUES "
                    StrSql = StrSql & "(" & NroCab
                    StrSql = StrSql & ",'" & Left(rsConsult!diahoradesde3, 2) & ":" & Right(rsConsult!diahoradesde3, 2)
                    StrSql = StrSql & "','" & Left(rsConsult!diahorahasta3, 2) & ":" & Right(rsConsult!diahorahasta3, 2)
                    StrSql = StrSql & "')"
                    objConn.Execute StrSql, , adExecuteNoRecords
               End If
               rsConsult2.Close
        End If
        rsConsult.MoveNext
    Loop
End If

'------------------------------------------------------------------
'Busco el detalle de la jornada
'------------------------------------------------------------------





         StrSql = " SELECT thnro, SUM(dlimonto) monto, SUM(dlicant) cantidad" & _
             " From cabliq  " & _
             " INNER JOIN proceso  ON proceso.pronro = cabliq.pronro AND proceso.pronro IN (" & Pronro & ")" & _
             " INNER JOIN periodo  ON proceso.pliqnro = periodo.pliqnro  " & _
             " INNER JOIN detliq   ON cabliq.cliqnro = detliq.cliqnro  AND cabliq.empleado =  " & Ternro & _
             " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1   and detliq.concnro IN (" & lista_conceptos & ")" & _
             " INNER JOIN tiph_con ON detliq.concnro = tiph_con.concnro AND   thnro IN (" & lista_hs_extras & ")" & _
             " GROUP by thnro "
            OpenRecordset StrSql, rsConsult
            
         ' Flog.writeline "liquidado" & StrSql
 If usa_gti = -1 Then
        '---LOG---
        Flog.writeline "Buscando datos del detalle de liquidacion"

                    StrSql = " SELECT thnro,adcanthoras, adfecha , horas " & _
                              " FROM gti_acumdiario " & _
                              " WHERE ternro= " & Ternro & _
                              " AND adfecha>= " & ConvFecha(pliqdesde) & _
                              " AND adfecha<= " & ConvFecha(pliqhasta) & _
                              " AND thnro IN (" & lista_hs_extras & ")" & _
                              " ORDER BY thnro "
                    OpenRecordset StrSql, rsAcumdiario
                      
         '  Flog.writeline "cargado:" & StrSql
        Do Until rsConsult.EOF
             CHmes = rsConsult!Cantidad
             THmes = rsConsult!thnro
             cambiohora = False
             Do While Not rsAcumdiario.EOF And Not cambiohora
                    FechaCompleta = Format(rsAcumdiario!adfecha, "Long Date")
                    'calcula horario de entrada y salida del día
                     Call CalcularhorarioES(Ternro, rsAcumdiario!adfecha, HoraEnt, HoraSal)
                   'calcula horario de las horas extras
                    If HoraEnt <> "" And HoraSal <> "" Then
                        Call CalcularhorarioHE(rsAcumdiario!adfecha, HoraEnt, HoraSal, rsAcumdiario!horas, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                    Else
                        ' calcula las hs extras a partir del horario cumplido
                        Call CalcularhorarioHE2(Ternro, THmes, rsAcumdiario!adfecha, rsAcumdiario!horas, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                    End If
                    Call CalcularPorcentaje(rsAcumdiario!thnro, Porcentaje)
                    Call CalcularValorHE(valorHoras, Porcentaje, rsAcumdiario!adcanthoras, ValorHE, TotalValorHE)
                    If THmes = rsAcumdiario!thnro Then
                        If rsAcumdiario!adcanthoras <= CHmes Then
                            CHmes = CHmes - rsAcumdiario!adcanthoras
                           If horas + minutos > 0 Then
                             StrSql = " INSERT INTO repHorasSupDet " & _
                                 " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                 "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                 "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                 "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                 "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                 " VALUES " & _
                                 "(" & NroCab & _
                                 ",'" & FechaCompleta & _
                                 "','" & HoraEnt & _
                                 "','" & HoraSal & _
                                 "','" & HED & _
                                 "','" & HEH & _
                                 "','" & horas & _
                                 "','" & minutos & _
                                 "','" & horas & _
                                 "','" & minutos & _
                                 "'," & valorHoras & _
                                 "," & Porcentaje & _
                                 "," & ValorHE & _
                                 "," & TotalValorHE & _
                                 "," & ConvFecha(rsAcumdiario!adfecha) & _
                                 ",0)"  'repHsSupDet_error = 0 no hubo errores
                                 objConn.Execute StrSql, , adExecuteNoRecords
                              End If
                                 rsAcumdiario.MoveNext
                         Else
                               aux = ""
                               If HoraEnt <> "" And HoraSal <> "" Then
                                  posdec = InStr(CStr(CHmes), ".")
                                  If posdec = 0 Then
                                        aux = CStr(CHmes) & ":00"
                                        
                                  Else
                                        decaux = CStr(Round(Right(CStr(CHmes), Len(CStr(CHmes)) - posdec) * 60 / 100))
                                        aux = CStr(Left(CStr(CHmes), posdec - 1)) & ":" & decaux
                                        
                                  End If
                                        If Len(aux) = 4 Then
                                            aux = "0" & aux
                                        End If
                                    aux2 = aux
                                    Call CalcularhorarioHE(rsAcumdiario!adfecha, HoraEnt, HoraSal, aux, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                                Else
                                    ' calcula las hs extras a partir del horario cumplido
                                    Call CalcularhorarioHE2(Ternro, THmes, rsAcumdiario!adfecha, aux, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                                End If
                                 Call CalcularValorHE(valorHoras, Porcentaje, CHmes, ValorHE, TotalValorHE)
                                If horas + minutos > 0 Then
                                         StrSql = " INSERT INTO repHorasSupDet " & _
                                     " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                 "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                 "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                 "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                 "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                     " VALUES " & _
                                     "(" & NroCab & _
                                     ",'" & FechaCompleta & _
                                     "','" & HoraEnt & _
                                     "','" & HoraSal & _
                                     "','" & HED & _
                                     "','" & HEH & _
                                     "','" & horas & _
                                     "','" & minutos & _
                                     "','" & horas & _
                                     "','" & minutos & _
                                     "'," & valorHoras & _
                                     "," & Porcentaje & _
                                     "," & ValorHE & _
                                     "," & TotalValorHE & _
                                     "," & ConvFecha(rsAcumdiario!adfecha) & _
                                     ",0)"
                                      objConn.Execute StrSql, , adExecuteNoRecords
                               End If
                              aux = ""
                              CHmes = Replace(rsAcumdiario!horas, ":", ".") - CHmes 'Replace(aux2, ":", "")
                              posdec = InStr(CStr(CHmes), ".")
                                  If posdec = 0 Then
                                        aux = CStr(CHmes) & ":00"
                                        
                                  Else
                                        decaux = CStr(Round(Right(CStr(CHmes), Len(CStr(CHmes)) - posdec) * 60 / 100))
                                        aux = CStr(Left(CStr(CHmes), posdec - 1)) & ":" & decaux
                                        
                                  End If
                                    If Len(aux) = 4 Then
                                            aux = "0" & aux
                                        End If
                                    aux2 = aux
                              If HoraEnt <> "" And HoraSal <> "" Then
                                    HoraEnt = HEH
                                     objFechasHoras.SumoHoras rsAcumdiario!adfecha, HEH, aux, fechaSalaux, HEHaux
                                     HoraSal = HEHaux
                                    Call CalcularhorarioHE(rsAcumdiario!adfecha, HoraEnt, HoraSal, aux, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                                Else
                                    ' calcula las hs extras a partir del horario cumplido
                                    Call CalcularhorarioHE2(Ternro, THmes, rsAcumdiario!adfecha, aux, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                                End If
                                 Call CalcularValorHE(valorHoras, Porcentaje, CHmes, ValorHE, TotalValorHE)
                                 If horas + minutos > 0 Then
                                    StrSql = " INSERT INTO repHorasSupDet " & _
                                            " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                        "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                        "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                        "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                        "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                            " VALUES " & _
                                            "(" & NroCab & _
                                            ",'" & FechaCompleta & _
                                            "','" & HoraEnt & _
                                            "','" & HoraSal & _
                                            "','" & HED & _
                                            "','" & HEH & _
                                            "','" & horas & _
                                            "','" & minutos & _
                                            "','" & horas & _
                                            "','" & minutos & _
                                            "'," & valorHoras & _
                                            "," & Porcentaje & _
                                            "," & ValorHE & _
                                            ",0" & _
                                            "," & ConvFecha(rsAcumdiario!adfecha) & _
                                            ",1)"  'repHsSupDet_error = 1 Hay mas horas en AcuDiario que en detliq
                                            objConn.Execute StrSql, , adExecuteNoRecords
                                     End If
                                     rsAcumdiario.MoveNext
                                     CHmes = 0
                        End If
                    Else
                        If THmes < rsAcumdiario!thnro Then
                            cambiohora = True
                        Else
                            If (Len(CStr(rsAcumdiario!adcanthoras)) - (InStr(CStr(rsAcumdiario!adcanthoras), "."))) = 1 And (InStr(CStr(rsAcumdiario!adcanthoras), ".")) <> 0 Then
                                posdec = Round(Right(CStr(rsAcumdiario!adcanthoras), Len(CStr(rsAcumdiario!adcanthoras)) - (InStr(CStr(rsAcumdiario!adcanthoras), "."))) * 60 / 10)
                            Else
                                 If (InStr(CStr(rsAcumdiario!adcanthoras), ".")) = 0 Then
                                     posdec = 0
                                 Else
                                      posdec = Round(Right(CStr(rsAcumdiario!adcanthoras), Len(CStr(rsAcumdiario!adcanthoras)) - (InStr(CStr(rsAcumdiario!adcanthoras), "."))) * 60 / 100)
                                End If
                            End If
                        
                          
                            '  StrSql = " INSERT INTO repHorasSupDet " & _
                                     " (repHsSupCabnro , repHsSupDet_Dia,repHsSupDet_hs, repHsSupDet_min " & _
                                     "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                     "  ,repHsSupDet_monto,repHsSupDet_fcorta, repHsSupDet_error)" & _
                                     " VALUES " & _
                                     "(" & NroCab & _
                                     ",'" & FechaCompleta & _
                                     "'," & rsAcumdiario!adcanthoras & "," & posdec & ",0," & Porcentaje & "," & valorHoras & ",0," & ConvFecha(rsAcumdiario!adfecha) & ",1)" 'repHsSupDet_error = 1 Hay mas horas en Acudiario que en Detliq
                                      StrSql = " INSERT INTO repHorasSupDet " & _
                                            " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                        "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                        "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                        "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                        "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                            " VALUES " & _
                                            "(" & NroCab & _
                                            ",'" & FechaCompleta & _
                                            "','" & HoraEnt & _
                                            "','" & HoraSal & _
                                            "','" & HED & _
                                            "','" & HEH & _
                                            "','" & horas & _
                                            "','" & minutos & _
                                            "','" & horas & _
                                            "','" & minutos & _
                                            "'," & valorHoras & _
                                            "," & Porcentaje & _
                                            "," & ValorHE & _
                                            ",0" & _
                                            "," & ConvFecha(rsAcumdiario!adfecha) & _
                                            ",1)"  'repHsSupDet_error = 1 Hay mas horas en AcuDiario que en detliq
                                     objConn.Execute StrSql, , adExecuteNoRecords
                             
                                     rsAcumdiario.MoveNext
                                     
                        End If
                    End If

             Loop
             'objConn.Execute StrSql, , adExecuteNoRecords
             If Not rsAcumdiario.EOF Then
                If CHmes > 0 Then
                     'posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 100)
                     If (Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) = 1 And (InStr(CStr(CHmes), ".")) <> 0 Then
                        posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 10)
                    Else
                          If (InStr(CStr(CHmes), ".")) = 0 Then
                            posdec = 0
                         Else
                            posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 100)
                         End If
                    End If
                    
                        StrSql = " INSERT INTO repHorasSupDet " & _
                                        " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_hs, repHsSupDet_min" & _
                                        "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                        "  ,repHsSupDet_monto ,repHsSupDet_fcorta, repHsSupDet_error)" & _
                                        " VALUES " & _
                                        "(" & NroCab & _
                                        ",'" & _
                                        "'," & CHmes & "," & posdec & ",0,0,0,0," & ConvFecha(Date) & ",2)" 'repHsSupDet_error = 2 Hay mas horas en Detliq que en Acudiario
                                     objConn.Execute StrSql, , adExecuteNoRecords
                   
                    cambiohora = False
                 End If
                  rsConsult.MoveNext
             Else
                If CHmes > 0 Then
                'posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 100)
                If (Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) = 1 And (InStr(CStr(CHmes), ".")) <> 0 Then
                        posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 10)
                    Else
                       If (InStr(CStr(CHmes), ".")) = 0 Then
                            posdec = 0
                       Else
                            posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 100)
                       End If
                    End If
                    
                            StrSql = " INSERT INTO repHorasSupDet " & _
                                                " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_hs, repHsSupDet_min" & _
                                                "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                                "  ,repHsSupDet_monto ,repHsSupDet_fcorta, repHsSupDet_error)" & _
                                                " VALUES " & _
                                                "(" & NroCab & _
                                                ",'" & _
                                                "'," & CHmes & "," & posdec & ",0,0,0,0," & ConvFecha(Date) & ",2)"   'repHsSupDet_error = 2 Hay mas horas en Detliq que en Acudiario
                                     objConn.Execute StrSql, , adExecuteNoRecords
                     
                End If
                  rsConsult.MoveNext
                  
             End If
        Loop
        If Not rsAcumdiario.EOF Then
            Do While Not rsAcumdiario.EOF
                   FechaCompleta = Format(rsAcumdiario!adfecha, "Long Date")
                    Call CalcularhorarioES(Ternro, rsAcumdiario!adfecha, HoraEnt, HoraSal)
                   ' Call CalcularhorarioHE(rsAcumdiario!adfecha, HoraEnt, HoraSal, rsAcumdiario!horas, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                    
                    If HoraEnt <> "" And HoraSal <> "" Then
                        Call CalcularhorarioHE(rsAcumdiario!adfecha, HoraEnt, HoraSal, rsAcumdiario!horas, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                    Else
                        
                        ' calcula las hs extras a partir del horario cumplido
                        Call CalcularhorarioHE2(Ternro, rsAcumdiario!thnro, rsAcumdiario!adfecha, rsAcumdiario!horas, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                        HoraEnt = HED
                        HoraSal = HEH
                    End If
                    
                    
                    Call CalcularPorcentaje(rsAcumdiario!thnro, Porcentaje)
                    Call CalcularValorHE(valorHoras, Porcentaje, rsAcumdiario!adcanthoras, ValorHE, TotalValorHE)
                        If horas + minutos > 0 Then
                                StrSql = " INSERT INTO repHorasSupDet " & _
                                                   " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                                   "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                                   "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                                   "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                                   "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                                   " VALUES " & _
                                                   "(" & NroCab & _
                                                   ",'" & FechaCompleta & _
                                                   "','" & HoraEnt & _
                                                   "','" & HoraSal & _
                                                   "','" & HED & _
                                                   "','" & HEH & _
                                                   "','" & horas & _
                                                   "','" & minutos & _
                                                   "','" & horas & _
                                                   "','" & minutos & _
                                                   "'," & valorHoras & _
                                                   "," & Porcentaje & _
                                                   "," & ValorHE & _
                                                  ",0" & _
                                                   "," & ConvFecha(rsAcumdiario!adfecha) & _
                                                   ",1)"  'repHsSupDet_error = 1 Hay mas horas en AcuDiario que en detliq
                                                   objConn.Execute StrSql, , adExecuteNoRecords
                                End If
                                     rsAcumdiario.MoveNext
            Loop
        End If
        
        rsConsult.Close
        rsAcumdiario.Close
   Else
                      '---LOG---
        Flog.writeline "Buscando datos del detalle de liquidacion"

       StrSql = " SELECT thnro,horcant adcanthoras, hordesde adfecha, horas,horhoradesde,horhorahasta, horcant " & _
                  " FROM gti_horcumplido" & _
                  " WHERE ternro= " & Ternro & _
                  " AND horfecgen>= " & ConvFecha(pliqdesde) & _
                  " AND horfecgen<= " & ConvFecha(pliqhasta) & _
                  " AND thnro IN (" & lista_hs_extras & ")" & _
                  " ORDER BY thnro "
                    OpenRecordset StrSql, rsAcumdiario
      
        'Flog.writeline "cargado" & StrSql
        Do Until rsConsult.EOF
             CHmes = rsConsult!Cantidad
             THmes = rsConsult!thnro
             cambiohora = False
             Do While Not rsAcumdiario.EOF And Not cambiohora
                    FechaCompleta = Format(rsAcumdiario!adfecha, "Long Date")
                    'calcula horario de entrada y salida del día
                    Call CalcularhorarioES(Ternro, rsAcumdiario!adfecha, HoraEnt, HoraSal)
                    
                   'calcula horario de las horas extras
                    
           HED = Left(rsAcumdiario!horhoradesde, 2) & ":" & Right(rsAcumdiario!horhoradesde, 2)
           HEH = Left(rsAcumdiario!horhorahasta, 2) & ":" & Right(rsAcumdiario!horhorahasta, 2)
           If Len(rsAcumdiario!horas) > 4 Then
                horas = Left(rsAcumdiario!horas, 2)
           Else
                horas = Left(rsAcumdiario!horas, 1)
           End If
           minutos = Right(rsAcumdiario!horas, 2)

                    Call CalcularPorcentaje(rsAcumdiario!thnro, Porcentaje)
                    Call CalcularValorHE(valorHoras, Porcentaje, rsAcumdiario!horcant, ValorHE, TotalValorHE)
                    If THmes = rsAcumdiario!thnro Then
                        If rsAcumdiario!horcant <= CHmes Then
                            CHmes = CHmes - rsAcumdiario!horcant
                            If horas + minutos > 0 Then
                                    StrSql = " INSERT INTO repHorasSupDet " & _
                                        " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                        "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                        "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                        "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                        "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                        " VALUES " & _
                                        "(" & NroCab & _
                                        ",'" & FechaCompleta & _
                                        "','" & HoraEnt & _
                                        "','" & HoraSal & _
                                        "','" & HED & _
                                        "','" & HEH & _
                                        "','" & horas & _
                                        "','" & minutos & _
                                        "','" & horas & _
                                        "','" & minutos & _
                                        "'," & valorHoras & _
                                        "," & Porcentaje & _
                                        "," & ValorHE & _
                                        "," & TotalValorHE & _
                                        "," & ConvFecha(rsAcumdiario!adfecha) & _
                                        ",0)"  'repHsSupDet_error = 0 no hubo errores
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                 End If
                                 rsAcumdiario.MoveNext
                         Else
                               aux = ""
                              
                                  posdec = InStr(CStr(CHmes), ".")
                                  If posdec = 0 Then
                                        aux = CStr(CHmes) & ":00"
                                        
                                  Else
                                        decaux = CStr(Round(Right(CStr(CHmes), Len(CStr(CHmes)) - posdec) * 60 / 100))
                                        aux = CStr(Left(CStr(CHmes), posdec - 1)) & ":" & decaux
                                        
                                  End If
                                    If Len(aux) = 4 Then
                                            aux = "0" & aux
                                        End If
                                    aux2 = aux
                                    Call CalcularhorarioHE(rsAcumdiario!adfecha, HED, HED, aux, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                              
                                 Call CalcularValorHE(valorHoras, Porcentaje, CHmes, ValorHE, TotalValorHE)
                                 If horas + minutos > 0 Then
                                            StrSql = " INSERT INTO repHorasSupDet " & _
                                        " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                    "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                    "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                    "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                    "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                        " VALUES " & _
                                        "(" & NroCab & _
                                        ",'" & FechaCompleta & _
                                        "','" & HoraEnt & _
                                        "','" & HoraSal & _
                                        "','" & HED & _
                                        "','" & HEH & _
                                        "','" & horas & _
                                        "','" & minutos & _
                                        "','" & horas & _
                                        "','" & minutos & _
                                        "'," & valorHoras & _
                                        "," & Porcentaje & _
                                        "," & ValorHE & _
                                        "," & TotalValorHE & _
                                        "," & ConvFecha(rsAcumdiario!adfecha) & _
                                        ",0)"
                                         objConn.Execute StrSql, , adExecuteNoRecords
                                 End If
                              aux = ""
                              CHmes = Replace(rsAcumdiario!horas, ":", ".") - CHmes 'Replace(aux2, ":", "")
                              posdec = InStr(CStr(CHmes), ".")
                                  If posdec = 0 Then
                                        aux = CStr(CHmes) & ":00"
                                        
                                  Else
                                        decaux = CStr(Round(Right(CStr(CHmes), Len(CStr(CHmes)) - posdec)))
                                        aux = CStr(Left(CStr(CHmes), posdec - 1)) & ":" & decaux
                                        
                                  End If
                                    If Len(decaux) = 1 Then
                                         aux = aux & "0"
                                    End If
                                    
                                    If Len(aux) = 4 Then
                                            aux = "0" & aux
                                        End If
                                    aux2 = aux
                              
                                    HoraEnt = HEH
                                     objFechasHoras.SumoHoras rsAcumdiario!adfecha, HEH, aux, fechaSalaux, HEHaux
                                     HoraSal = HEHaux
                                    Call CalcularhorarioHE(rsAcumdiario!adfecha, HoraEnt, HoraSal, aux, HED, HEH, horas, minutos) 'HED:Hora extra desde - HEH:Hora extra hasta
                                    Call CalcularValorHE(valorHoras, Porcentaje, CHmes, ValorHE, TotalValorHE)
                                    If horas + minutos > 0 Then
                                    StrSql = " INSERT INTO repHorasSupDet " & _
                                            " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                        "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                        "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                        "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                        "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                            " VALUES " & _
                                            "(" & NroCab & _
                                            ",'" & FechaCompleta & _
                                            "','" & HoraEnt & _
                                            "','" & HoraSal & _
                                            "','" & HED & _
                                            "','" & HEH & _
                                            "','" & horas & _
                                            "','" & minutos & _
                                            "','" & horas & _
                                            "','" & minutos & _
                                            "'," & valorHoras & _
                                            "," & Porcentaje & _
                                            "," & ValorHE & _
                                            "," & TotalValorHE & _
                                            "," & ConvFecha(rsAcumdiario!adfecha) & _
                                            ",1)"  'repHsSupDet_error = 1 Hay mas horas en HC que en detliq
                                     End If
                                     objConn.Execute StrSql, , adExecuteNoRecords
                                     rsAcumdiario.MoveNext
                                     CHmes = 0
                        End If
                    Else
                        If THmes < rsAcumdiario!thnro Then
                            cambiohora = True
                        Else
                            'posdec = Round(Right(CStr(rsAcumdiario!horcant), Len(CStr(rsAcumdiario!horcant)) - (InStr(CStr(rsAcumdiario!horcant), "."))) * 60 / 100)
                            'StrSql = " INSERT INTO repHorasSupDet " & _
                                     " (repHsSupCabnro , repHsSupDet_Dia,repHsSupDet_hs, repHsSupDet_min " & _
                                     "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                     "  ,repHsSupDet_monto, repHsSupDet_error)" & _
                                     " VALUES " & _
                                     "(" & NroCab & _
                                     ",'" & _
                                     "'," & horas & "," & minutos & ",0,0,0,0,1)" 'repHsSupDet_error = 1 Hay mas horas en  Horario cumplido que en det liq
                                      If horas + minutos > 0 Then
                                            StrSql = " INSERT INTO repHorasSupDet " & _
                                                 " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                             "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                             "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                             "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                             "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                                 " VALUES " & _
                                                 "(" & NroCab & _
                                                 ",'" & FechaCompleta & _
                                                 "','" & HoraEnt & _
                                                 "','" & HoraSal & _
                                                 "','" & HED & _
                                                 "','" & HEH & _
                                                 "','" & horas & _
                                                 "','" & minutos & _
                                                 "','" & horas & _
                                                 "','" & minutos & _
                                                 "'," & valorHoras & _
                                                 "," & Porcentaje & _
                                                 "," & ValorHE & _
                                                 "," & TotalValorHE & _
                                                 "," & ConvFecha(rsAcumdiario!adfecha) & _
                                                 ",1)"  'repHsSupDet_error = 1 Hay mas horas en HC que en detliq
                                     
                                         objConn.Execute StrSql, , adExecuteNoRecords
                                     End If
                                     rsAcumdiario.MoveNext
                                     
                        End If
                    End If

             Loop
             'objConn.Execute StrSql, , adExecuteNoRecords
             If Not rsAcumdiario.EOF Then
                If CHmes > 0 Then
                     'posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 100)
                     If (Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) = 1 And (InStr(CStr(CHmes), ".")) <> 0 Then
                        posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 10)
                    Else
                          If (InStr(CStr(CHmes), ".")) = 0 Then
                            posdec = 0
                       Else
                            posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 100)
                       End If
                    End If
                    
                        StrSql = " INSERT INTO repHorasSupDet " & _
                                        " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_hs, repHsSupDet_min" & _
                                        "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                        "  ,repHsSupDet_monto , repHsSupDet_fcorta,repHsSupDet_error)" & _
                                        " VALUES " & _
                                        "(" & NroCab & _
                                        ",'" & _
                                        "'," & CHmes & "," & posdec & ",0,0,0,0," & ConvFecha(Date) & ",2)"  'repHsSupDet_error = 2 Hay mas horas en Detliq que en Acudiario
                                     objConn.Execute StrSql, , adExecuteNoRecords
                    
                    
                    cambiohora = False
                 End If
                  rsConsult.MoveNext
             Else
                If CHmes > 0 Then
                    If (Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) = 1 And (InStr(CStr(CHmes), ".")) <> 0 Then
                        posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 10)
                    Else
                       If (InStr(CStr(CHmes), ".")) = 0 Then
                            posdec = 0
                       Else
                            posdec = Round(Right(CStr(CHmes), Len(CStr(CHmes)) - (InStr(CStr(CHmes), "."))) * 60 / 100)
                       End If
                    End If
                   
                            StrSql = " INSERT INTO repHorasSupDet " & _
                                                " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_hs, repHsSupDet_min" & _
                                                "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                                "  ,repHsSupDet_monto ,repHsSupDet_fcorta, repHsSupDet_error)" & _
                                                " VALUES " & _
                                                "(" & NroCab & _
                                                ",'" & _
                                                "'," & CHmes & "," & posdec & ",0,0,0,0," & ConvFecha(Date) & ",2)"  'repHsSupDet_error = 2 Hay mas horas en Detliq que en Acudiario
                                     objConn.Execute StrSql, , adExecuteNoRecords
                   
                End If
                 ' Exit Do
                 rsConsult.MoveNext
                  
             End If
        Loop
        If Not rsAcumdiario.EOF Then
            Do While Not rsAcumdiario.EOF
                   FechaCompleta = Format(rsAcumdiario!adfecha, "Long Date")
           Call CalcularhorarioES(Ternro, rsAcumdiario!adfecha, HoraEnt, HoraSal)
           HED = Left(rsAcumdiario!horhoradesde, 2) & ":" & Right(rsAcumdiario!horhoradesde, 2)
           HEH = Left(rsAcumdiario!horhorahasta, 2) & ":" & Right(rsAcumdiario!horhorahasta, 2)
           horas = Left(rsAcumdiario!horas, 2)
           minutos = Right(rsAcumdiario!horas, 2)

                    Call CalcularPorcentaje(rsAcumdiario!thnro, Porcentaje)
                    Call CalcularValorHE(valorHoras, Porcentaje, rsAcumdiario!adcanthoras, ValorHE, TotalValorHE)
                    If horas + minutos > 0 Then
                            StrSql = " INSERT INTO repHorasSupDet " & _
                                               " (repHsSupCabnro , repHsSupDet_Dia, repHsSupDet_DiaHD, repHsSupDet_DiaHH" & _
                                               "  ,repHsSupDet_horaD,repHsSupDet_horaH,repHsSupDet_hs " & _
                                               "  ,repHsSupDet_min,repHsSupDet_hsr,repHsSupDet_minr " & _
                                               "  ,repHsSupDet_valor,repHsSupDet_inc,repHsSupDet_valor_he" & _
                                               "  ,repHsSupDet_monto,repHsSupDet_fcorta,repHsSupDet_error)" & _
                                               " VALUES " & _
                                               "(" & NroCab & _
                                               ",'" & FechaCompleta & _
                                               "','" & HoraEnt & _
                                               "','" & HoraSal & _
                                               "','" & HED & _
                                               "','" & HEH & _
                                               "','" & horas & _
                                               "','" & minutos & _
                                               "','" & horas & _
                                               "','" & minutos & _
                                               "'," & valorHoras & _
                                               "," & Porcentaje & _
                                               "," & ValorHE & _
                                               "," & TotalValorHE & _
                                               "," & ConvFecha(rsAcumdiario!adfecha) & _
                                               ",1)"  'repHsSupDet_error = 1 Hay mas horas en AcuDiario que en detliq
                                               objConn.Execute StrSql, , adExecuteNoRecords
                                     End If
                                     rsAcumdiario.MoveNext
            Loop
        End If
        
        rsConsult.Close
        rsAcumdiario.Close
         End If
Exit Sub

MError:
    Flog.writeline "Error en empleado: " & Legajo & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub



