Attribute VB_Name = "repFormularioF1827"
Global Const Version = "1.00" ' sebastian stremel
Global Const FechaModificacion = "11/10/2011"
Global Const UltimaModificacion = "" 'Version Inicial

'--------------------------------------------------------------
'--------------------------------------------------------------
Option Explicit

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

Global Pagina As Long
Global tipoModelo As Integer
Global arrTipoConc(1000) As Integer
Global tituloReporte As String

Global tenro1 As Integer
Global estrnro1 As Integer
Global tenro2 As Integer
Global estrnro2 As Integer
Global tenro3 As Integer
Global estrnro3 As Integer
Global fecEstr As String

Global empresa As String
Global Empnro As Long
Global Empnroestr As Long
Global Centcostnroestr As Long
Global emprTer As Long
Global emprDire As String
Global emprCuit

Global IdUser As String
Global Fecha As Date
Global Hora As String

Global listapronro As String        'Lista de procesos

Global totalEmpleados
Global cantRegistros

Global incluyeAgencia As Integer
Global NroAcDiasTrabajados As Long

Global TipoEstructura As Long

Global VectorEsConc(40) As Boolean 'True = Concepto    False = Acumulador
Global VectorNroACCO(40) As Long
Global VectorValorACCO(40) As Double

Global CantEmpGrabados As Long 'Cantidad de empleados grabados
Global acuLiqPago As Double
Global acuImpUnicoR As Double
Global acuCotPrev As Double
Global acuIsapre As Double
Global acuRentaTotal As Double
Global acuRentasPagadas As Double
Global acuRentasAccesorias As Double
Global acuRemImponible As Double
Global acuRentaTotalNeta As Double
Global acuImpUnicoRet As Double
Global acuTotalRem As Double
Global acuTotalEmp As Integer
Global esCotPrev As Boolean
Global rutAfp As String
Global NuevaFecha As String
Global filtro
Global Ordenamiento As String
Global orden2
'Global orden As String
'Global ord As String

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

Dim historico As Boolean
'Dim param
Dim proNro As Long
Dim ternro  As Long
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim rsAge As New ADODB.Recordset
Dim rsEmpresas As New ADODB.Recordset
Dim rsPeriodo As New ADODB.Recordset
'Dim acunroSueldo
Dim I
Dim PID As String
Dim tituloReporte As String

Dim parametros As String
Dim ArrParametros
Dim strTempo As String
Dim orden As Long
Dim ord
    
'Dim orden As String
'Dim ord As String

    
Dim arrpliqnro
Dim listapliqnro
Dim pliqNro As Long
Dim pliqMes As Long
Dim pliqAnio As Long
Dim rsConsult2 As New ADODB.Recordset

Dim IdUser As String
Dim Fecha  As Date
Dim Hora As String

Dim fecdesde
Dim fechasta
Dim aprov
Dim razSoc As String
Dim Rut As String
Dim domicilio As String
Dim comuna As String
Dim email As String
Dim fax As String
Dim tel As String
Dim rutEmp As String
Dim mes As String
Dim tipoContrato As String
Dim domnro As String
Dim empNom As String
'Dim acuImpUnicoR
'Dim acuCotPrev
'Dim acuIsapre
'Dim acuRentaTotal
'Dim acuRentasPagadas
'Dim acuRentasAccesorias
'Dim acuRemImponible
'Dim acuRentaTotalNeta
'Dim acuImpUnicoRet
'Dim acuTotalRem
'Dim acuTotalEmp
'Dim esCotPrev
'Dim empNom As String
'Dim filtro
'Dim Ordenamiento As String
'Dim NuevaFecha
'Dim orden2

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
    
    Nombre_Arch = PathFLog & "ReporteFormularioF1827" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    Flog.writeline "Inicio Proceso: " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'OpenConnection strconexion, objConn
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    HuboErrores = False
    
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
       
       'Obtengo los parametros del proceso
       IdUser = objRs!IdUser
       Fecha = objRs!bprcfecha
       Hora = objRs!bprchora
       parametros = objRs!bprcparam
       Flog.writeline " parametros del proceso --> " & parametros
       ArrParametros = Split(parametros, "@")
       Flog.writeline " limite del array --> " & UBound(ArrParametros)
       'Dim cparam
       'cparam = UBound(ArrParametros)
       
       'obtengo el perido desde
       'If CLng(ArrParametros(3)) <> 0 Then
       '     pliqdesde = ArrParametros(3)
       'Else
       '     Flog.writeline "No Se encontro el periodo desde"
       '     HuboErrores = True
       'End If
       
       
       'Obtengo el titulo del reporte
       tituloReporte = ArrParametros(0)
       
       'obtengo el periodo liq desde
       If ((ArrParametros(1)) <> 0) Then
            pliqdesde = (ArrParametros(1))
       Else
            Flog.writeline "el periodo liq desde dio error"
            HuboErrores = True
       End If
       
       'obtengo el periodo liq hasta
       If ((ArrParametros(2)) <> 0) Then
            pliqhasta = (ArrParametros(2))
       Else
            Flog.writeline "el periodo liq desde dio error"
            HuboErrores = True
       End If
       
       'obtengo la fecha desde
       If ArrParametros(3) <> 0 Then
            fecdesde = (ArrParametros(3))
       Else
            Flog.writeline "el periodo liq desde dio error"
            HuboErrores = True
       End If
       
       'obtengo la fecha hasta
       If ArrParametros(4) <> 0 Then
            fechasta = (ArrParametros(4))
       Else
            Flog.writeline "el periodo liq hasta dio error"
            HuboErrores = True
       End If
       
       'obtengo si el proceso esta aprobado o no
       If (ArrParametros(5)) <> 0 Then
            aprov = (ArrParametros(5))
       Else
            Flog.writeline "estado del proceso dio error"
            HuboErrores = True
       End If
       
       'Obtengo la lista de procesos
       If ArrParametros(6) <> "" Then
            listapronro = ArrParametros(6)
       Else
            Flog.writeline "estado del proceso dio error"
            HuboErrores = True
       End If
       
       'tenro1
       If ArrParametros(7) <> "" Then
            tenro1 = ArrParametros(7)
       Else
            Flog.writeline "tenro1 error"
            HuboErrores = True
       End If
       
       'estrnro1
       If ArrParametros(8) <> "" Then
            estrnro1 = ArrParametros(8)
       Else
            Flog.writeline "tenro1 error"
            HuboErrores = True
       End If
       
       'tenro2
       If ArrParametros(9) <> "" Then
            tenro2 = ArrParametros(9)
       Else
            Flog.writeline "tenro2 error"
            HuboErrores = True
       End If
       
       'estrnro2
       If ArrParametros(10) <> "" Then
            estrnro2 = ArrParametros(10)
       Else
            Flog.writeline "estrnro2 error"
            HuboErrores = True
       End If
       
       'tenro3
       If ArrParametros(11) <> "" Then
            tenro3 = ArrParametros(11)
       Else
            Flog.writeline "tenro3 error"
            HuboErrores = True
       End If
       
       'estrnro3
       If ArrParametros(12) <> "" Then
            estrnro3 = ArrParametros(12)
       Else
            Flog.writeline "estrnro3 error"
            HuboErrores = True
       End If
       
              
       'Obtengo el numero de empresa
       If (ArrParametros(13)) <> 0 Then
            Empnroestr = (ArrParametros(13))
       Else
            Flog.writeline "No Se selecciono el parametro Empresa. "
            HuboErrores = True
       End If
       
       'obtengo nombre de la empresa
       If (ArrParametros(14) <> 0) Then
            empNom = ArrParametros(14)
       Else
            Flog.writeln "No se selecciono el parametro nombre de empresa. "
            
       End If
       
    
       
       'fecha
       NuevaFecha = ArrParametros(15)
        
       'orden
       Ordenamiento = ArrParametros(16)

       'obtengo el tipo de filtro empleado
       filtro = ArrParametros(17)
       
       'orden
       ord = ArrParametros(18)
       
       'orden2
       orden2 = ArrParametros(19)
       
              
       'EMPIEZA EL PROCESO
       Flog.writeline "Generando el reporte"
                
        'busco el rut de la empresa
         StrSql = "select * FROM estructura "
         StrSql = StrSql & " inner join empresa on empresa.estrnro = estructura.estrnro "
         StrSql = StrSql & " inner join tercero on empresa.ternro = tercero.ternro "
         StrSql = StrSql & " inner join ter_doc on ter_doc.ternro= tercero.ternro "
         StrSql = StrSql & " inner join cabdom on cabdom.ternro = tercero.ternro "
         StrSql = StrSql & " inner join detdom on detdom.domnro = cabdom.domnro "
         StrSql = StrSql & " WHERE estructura.estrnro = " & Empnroestr
         OpenRecordset StrSql, objRs2
         If Not objRs2.EOF Then
            If objRs2!domnro <> "" Then
                domnro = objRs2!domnro
            End If
            
            If objRs2!terrazsoc <> "" Then
                razSoc = objRs2!terrazsoc
            Else
                razSoc = ""
            End If
            
            If objRs2!nrodoc <> "" Then
                Rut = objRs2!nrodoc
            Else
                Rut = ""
            End If
            
            If objRs2!calle <> "" Then
                domicilio = objRs2!calle
            Else
                domicilio = ""
            End If
            
            If objRs2!nro <> "" Then
                domicilio = domicilio & ", " & objRs2!nro
            Else
                domicilio = domicilio
            End If
            
            If objRs2!sector <> "" Then
                domicilio = domicilio & ", " & objRs2!sector
            Else
                domicilio = domicilio
            End If
            
            If objRs2!torre <> "" Then
                domicilio = domicilio & ", " & objRs2!torre
            Else
                domicilio = domicilio
            End If
            
            If objRs2!piso <> "" Then
                domicilio = domicilio & ", " & objRs2!piso
            Else
                domicilio = domicilio
            End If
            
            If objRs2!oficdepto <> "" Then
                domicilio = domicilio & ", " & objRs2!oficdepto & " ,"
            Else
                domicilio = domicilio
            End If
            
         End If
         
         domicilio = domicilio
         
         If objRs2!auxchr2 <> "" Then
            comuna = objRs2!auxchr2
         Else
            comuna = ""
         End If
         
         If objRs2!email <> "" Then
            email = objRs2!email
         Else
            email = ""
         End If
         objRs2.Close
         
         'busco el fax y telefono
         StrSql = " select * from telefono "
         StrSql = StrSql & " inner join tipotel on telefono.tipotel = tipotel.titelnro "
         StrSql = StrSql & " where telefono.domnro =  " & domnro
         OpenRecordset StrSql, objRs2
         Do While Not objRs2.EOF
            If objRs2!tipotel = 3 Then
                fax = objRs2!telnro
            End If
            
            If objRs2!tipotel = 1 Then
                tel = objRs2!telnro
            End If
            
         objRs2.MoveNext
         Loop
         
         StrSql = "insert into rep_formularioF1827 (bpronro,rut,razsoc,domicilio,comuna,email,fax,tel,rentaTotalNetaPag,ImpUnicRet1,ImpUnicRet2,totalRemImpo,rentaTotalNetaPag2,impUnicoRetenido,mayorRetsol,totalCasosInf,prodesc)"
         StrSql = StrSql & " VALUES "
         StrSql = StrSql & "(" & NroProceso
         StrSql = StrSql & "," & Rut
         StrSql = StrSql & ",'" & razSoc & "'"
         StrSql = StrSql & ",'" & domicilio & "'"
         StrSql = StrSql & ",'" & comuna & "'"
         StrSql = StrSql & ",'" & email & "'"
         StrSql = StrSql & ",'" & fax & "'"
         StrSql = StrSql & ",'" & tel & "'"
         StrSql = StrSql & "," & 0
         StrSql = StrSql & "," & 0
         StrSql = StrSql & "," & 0
         StrSql = StrSql & "," & 0
         StrSql = StrSql & "," & 0
         StrSql = StrSql & "," & 0
         StrSql = StrSql & "," & 0
         StrSql = StrSql & "," & 0
         StrSql = StrSql & ",'" & tituloReporte & "'"
         StrSql = StrSql & ")"
         objConn.Execute StrSql, , adExecuteNoRecords
         Flog.writeline " Se Grabo el Formulario F1827"
            
    
        'empieza el detalle del reporte
        
        'Obtengo la Configuracion del Confrep
        objRs2.Close
        
        StrSql = "SELECT * FROM confrep"
        StrSql = StrSql & " WHERE repnro = 356"
        StrSql = StrSql & " ORDER BY confnrocol"
        OpenRecordset StrSql, objRs2
        
        Flog.writeline "Obtengo los datos del confrep"
        
        If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep para el reporte"
          HuboErrores = True
        End If
       
       
        If Not HuboErrores Then
            Do While Not objRs2.EOF
                Flog.writeline "Columna " & objRs2!confnrocol
            Select Case objRs2!confnrocol
                Case 1
                    'acum liq a pago
                    acuLiqPago = objRs2!confval
                Case 2
                    'acum imp unico ret
                    acuImpUnicoR = objRs2!confval
                Case 3
                    'acumulador o concepto cotizacion previsional
                    If objRs2!conftipo = "CO" Then
                        esCotPrev = True
                  
                    StrSql = "SELECT concnro FROM concepto WHERE conccod = " & objRs2!confval
                    If Not EsNulo(objRs2!confval2) Then
                        StrSql = StrSql & " OR conccod = '" & objRs2!confval2 & "'"
                    End If
                   OpenRecordset StrSql, objRs3
                  
                   If objRs3.EOF Then
                     acuCotPrev = 0
                   Else
                     acuCotPrev = objRs3!concnro
                   End If
                  
                   objRs3.Close
                Else
                   esCotPrev = False
                   acuCotPrev = objRs2!confval
                End If
                    
                Case 4
                    'acu issapre
                    acuIsapre = objRs2!confval
                Case 5
                    'acu renta total
                    acuRentaTotal = objRs2!confval
                Case 6
                    'acu rentas pagadas
                    acuRentasPagadas = objRs2!confval
                Case 7
                    'acu renta accesorias
                    acuRentasAccesorias = objRs2!confval
                Case 8
                    'acu rem imponible
                    acuRemImponible = objRs2!confval
                Case 9
                    'acu renta total neta
                    acuRentaTotalNeta = objRs2!confval
                Case 10
                    'acu imp unico retenido
                    acuImpUnicoRet = objRs2!confval
                Case 12
                    'acu totalEmp
                    acuTotalEmp = objRs2!confval
                
                Case 13
                    'rut afp
                    rutAfp = objRs2!confval
            End Select
            objRs2.MoveNext
            Loop
            
           'Obtengo los empleados sobre los que tengo que generar el reporte
           If filtro <> 3 Then
                'CargarEmpleados(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal empresa As Long)
                CargarEmpleados NroProceso, rsEmpl, 0
                If Not rsEmpl.EOF Then
                    cantRegistros = rsEmpl.RecordCount
                    Flog.writeline "Cantidad de empleados a procesar: " & cantRegistros
                    CantEmpGrabados = 0 'Cantidad de empleados Guardados
                Else
                    Flog.writeline "No hay empleados para el filtro seleccionado."
                    StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Exit Sub
                End If
           
           Else 'si se selecciono todos los empleados
                Flog.writeline " se selecciono la opcion todos los empleados."
                'arrpronro = Split(listapronro, ",")
                CargarEmpleados_v1 listapronro, rsEmpl, 0
                If Not rsEmpl.EOF Then
                    cantRegistros = rsEmpl.RecordCount
                    Flog.writeline "Cantidad de empleados a procesar: " & cantRegistros
                    CantEmpGrabados = 0 'Cantidad de empleados Guardados
                Else
                    Flog.writeline "No hay empleados para el filtro seleccionado."
                    StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Exit Sub
                End If
           End If
        
           'Actualizo Barch Proceso
           StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(rsEmpl.RecordCount) & "' WHERE bpronro = " & NroProceso
        
           objConn.Execute StrSql, , adExecuteNoRecords
     
                   
           'Obtengo la lista de procesos
           arrpronro = Split(listapronro, ",")
            
           ord = 0
     
           'Genero por cada empleado un registro
           IncPorc = (99 / CEmpleadosAProc)
           While Not rsEmpl.EOF
                'arrpronro = Split(listapronro, ",")
                EmpErrores = False
                ternro = rsEmpl!ternro
                orden = ord
                'Ordenamiento = Ordenamiento
                Flog.writeline ""
                Flog.writeline "Generando datos empleado " & ternro
                        
                'Call ReporteFormularioF1827(arrpronro(), ternro, tituloReporte, orden)
                Call ReporteFormularioF1827(listapronro, ternro, tituloReporte, orden)
                
                                                
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                
                'Resto uno a la cantidad de registros
                cantRegistros = rsEmpl.RecordCount
                
                'Actualizo
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                         ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                        
                objConn.Execute StrSql, , adExecuteNoRecords
                
                ord = ord + 1
                
                'Borro batch empleado
                '****************************************************************
                StrSql = "DELETE  FROM batch_empleado "
                StrSql = StrSql & " WHERE bpronro = " & NroProceso
                StrSql = StrSql & " AND ternro = " & ternro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                
                'Incremento el progreso
                Progreso = Progreso + IncPorc
                'Progreso = Replace(Progreso, ",", ".")
                'Inserto progreso
                'StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Int(Progreso)
                TiempoInicialProceso = GetTickCount
                MyBeginTrans
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Replace(Progreso, ",", ".")
                    StrSql = StrSql & " , bprctiempo = " & TiempoInicialProceso
                    StrSql = StrSql & " WHERE bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'objconnProgreso.Execute StrSql, , adExecuteNoRecords
                MyCommitTrans
                'lo muevo al último registro
                rsEmpl.MoveNext
                
           Wend
           
          
        End If
    Else
        
        Exit Sub
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline
    Flog.writeline "************************************************************"
    Flog.writeline "Fin :" & Now
    Flog.writeline "Cantidad de empleados guardados en el reporte: " & CantEmpGrabados
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
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub ReporteFormularioF1827(ListaPro As String, ternro As Long, descripcion As String, orden As Long)

'Sub ReporteFormularioF1827(ListaPro As String, ternro As Long, descripcion As String)
Dim StrSql As String
Dim rsConsult As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim pliqNro As Long
Dim pliqMes As Integer
Dim pliqAnio As Long
Dim documento  As String
Dim cliqnro As Long
Dim EmpTernro As Long
Dim DiasTrabajados
Dim estrnomb1
Dim estrnomb2
Dim estrnomb3
Dim tenomb1
Dim tenomb2
Dim tenomb3
Dim proNro
Dim DescEstructura As String
Dim sql As String
Dim I As Integer
Dim direccion As String
Dim G
Dim GrabaEmpleado As Boolean
Dim Rut As String
Dim prodesc As String
Dim pliqdesde
Dim pliqhasta
Dim relContractual
Dim aculiq
Dim rentaTotal
Dim rutInstitucional
Dim rutSalud

'variables para insertar detalle del empleado
Dim det_rut As String
Dim det_mes As String
Dim det_relcontractual As String
Dim det_rentatotal As Double
Dim det_impunico As Double
Dim det_mayor_ret_sol As Double
Dim det_monto_cot_prev As Double
Dim det_rut_prev As String
Dim det_monto_salud As Double
Dim det_rut_salud As String

'variables totales anuales sin actualizar
Dim totalneto As Double
Dim rentasanuales As Double
Dim rentasaccesorias As Double
Dim totalimponible As Double
Dim ListaProAux() As String
Dim rsaux As New ADODB.Recordset

'variables totales anuales actualizados
Dim l_acuRentaTotalNeta As Double
Dim l_acuImpUnicoRetenido As Double
Dim l_acuTotalEmp As Integer

On Error GoTo MError

'Inicializo Conceptos
'For I = 1 To 40
'   VectorValorACCO(I) = 0 'As Double
'Next

StrSql = "select * from rep_formulariof1827 inner join batch_proceso on rep_formulariof1827.bpronro = batch_proceso.bpronro where rep_formulariof1827.bpronro=" & NroProceso
OpenRecordset StrSql, rsaux

If Not rsaux.EOF Then
    totalneto = rsaux!rentaTotalNetaPag
    rentasanuales = rsaux!impUnicRet1
    rentasaccesorias = rsaux!impUnicRet2
    totalimponible = rsaux!totalRemImpo
    l_acuRentaTotalNeta = rsaux!rentaTotalNetaPag2
    l_acuImpUnicoRetenido = rsaux!impUnicoRetenido
    l_acuTotalEmp = rsaux!totalCasosInf
    'Progreso = rsaux!bprcprogreso
End If
rsaux.Close


'*********************************************************************
'Ciclo por todos los procesos seleccionados del periodo
'*********************************************************************
GrabaEmpleado = False
det_rut = ""
det_mes = ""
det_relcontractual = ""
det_rentatotal = 0
det_impunico = 0
det_mayor_ret_sol = 0
det_monto_cot_prev = 0
det_rut_prev = ""
det_monto_salud = 0
det_rut_salud = ""

ListaProAux = Split(listapronro, ",")

For I = 0 To UBound(ListaProAux)
                          
   proNro = ListaProAux(I)


        '------------------------------------------------------------------
        'Busco los datos del periodo actual
        '------------------------------------------------------------------
        StrSql = " SELECT periodo.*, proceso.profecpago, proceso.prodesc, cabliq.cliqnro FROM periodo "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " AND proceso.pronro= " & proNro
        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " AND cabliq.empleado= " & ternro
        
        '---LOG---
        Flog.writeline "Buscando datos del periodo para el proceso: " & proNro
        
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           pliqNro = rsConsult!pliqNro
           pliqMes = rsConsult!pliqMes
           pliqAnio = rsConsult!pliqAnio
           prodesc = rsConsult!prodesc
           cliqnro = rsConsult!cliqnro
           pliqdesde = rsConsult!pliqdesde
           pliqhasta = rsConsult!pliqhasta
        
          rsConsult.Close
        
            'busco el rut del empleado sebastian stremel 06-10-2011
            StrSql = "select * from empleado "
            StrSql = StrSql & " inner join ter_doc on empleado.ternro=ter_doc.ternro "
            StrSql = StrSql & " Where tidnro = 1 And Empleado.ternro = " & ternro
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                det_rut = rsConsult!nrodoc
            Else
                Flog.writelinen "el empleado no tiene un rut asociado"
            End If
            rsConsult.Close
            
            'busco la relacion contractual del empleado sebastian stremel 06-10-2011
            StrSql = " select * from his_estructura "
            StrSql = StrSql & " inner join estructura on estructura.estrnro=his_estructura.estrnro "
            StrSql = StrSql & " Where his_estructura.ternro = " & ternro
            StrSql = StrSql & " and his_estructura.tenro=18 "
            StrSql = StrSql & " and " & ConvFecha(pliqdesde) & " >= htetdesde and ((" & ConvFecha(pliqhasta) & " <=htethasta) or ((htethasta is null)))"
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                det_relcontractual = rsConsult!estrdabr
            Else
                Flog.writeline "el empleado no tiene un tipo de estrucutura contrato para el periodo"
            End If
            rsConsult.Close
            
            'busco valor  de la renta total neta pagada sebastian stremel
            sql = " SELECT almonto "
            sql = sql & " FROM acu_liq"
            sql = sql & " WHERE acunro = " & acuLiqPago
            sql = sql & " AND cliqnro =  " & cliqnro
            OpenRecordset sql, rsConsult
            If Not rsConsult.EOF Then
                If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                    det_rentatotal = det_rentatotal + 0
                Else
                    det_rentatotal = det_rentatotal + Replace(CDbl(rsConsult!almonto), ",", ".")
                End If
            Else
                Flog.writeline "no hay valor para renta total neta pagada"
                
            End If
            rsConsult.Close
            
            'busco valor del impuesto unico retenido actualizado
            sql = " SELECT almonto "
            sql = sql & " FROM acu_liq"
            sql = sql & " WHERE acunro = " & acuImpUnicoR
            sql = sql & " AND cliqnro =  " & cliqnro
            OpenRecordset sql, rsConsult
            If Not rsConsult.EOF Then
                If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                    det_impunico = det_impunico + 0
                Else
                    det_impunico = det_impunico + CDbl(rsConsult!almonto)
                    det_impunico = Replace(acuImpUnicoR, ",", ".")
                End If
            Else
                Flog.writeline "no hay valor para el valor del impuesto unico retenido actualizado"
                
            End If
            rsConsult.Close
            
            'busco el valor de monto cotizacion previsional
            If esCotPrev Then
                sql = " SELECT almonto "
                sql = sql & " FROM acu_liq"
                sql = sql & " WHERE acunro = " & acuCotPrev
                sql = sql & " AND cliqnro =  " & cliqnro
                OpenRecordset sql, rsConsult
                If Not rsConsult.EOF Then
                    
                    If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                        det_monto_cot_prev = det_monto_cot_prev + 0
                    Else
                        det_monto_cot_prev = det_monto_cot_prev + CDbl(rsConsult!almonto)
                        det_monto_cot_prev = Replace(det_monto_cot_prev, ",", ".")
                    End If
                    
                Else
                    Flog.writeline "no hay valor para el valor del monto cotizacion previsional"
                    
                End If
            Else
                sql = " SELECT SUM (DISTINCT detliq.dlimonto) AS almonto "
                sql = sql & " FROM cabliq "
                sql = sql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro AND cabliq.pronro = " & proNro
                sql = sql & " INNER JOIN periodo ON proceso.pliqnro = periodo.pliqnro "
                sql = sql & " INNER JOIN detliq  ON cabliq.cliqnro = detliq.cliqnro  "
                sql = sql & " INNER JOIN batch_empleado  ON batch_empleado.ternro = cabliq.empleado "
                sql = sql & " AND detliq.concnro = " & acuCotPrev
                OpenRecordset sql, rsConsult
                If Not rsConsult.EOF Then
                    If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                        det_monto_cot_prev = det_monto_cot_prev + 0
                    Else
                        det_monto_cot_prev = det_monto_cot_prev + Replace(CDbl(rsConsult!almonto), ",", ".")
                    End If
                Else
                    Flog.writeline "no hay valor para el valor de monto de cotizacion previsional"
                    
                End If
            End If
            rsConsult.Close
            
            'busco rut institucional previsional
            StrSql = " select * from his_estructura "
            StrSql = StrSql & " inner join estructura on estructura.estrnro=his_estructura.estrnro "
            StrSql = StrSql & " inner join ter_doc on ter_doc.ternro = his_estructura.ternro "
            StrSql = StrSql & " Where his_estructura.ternro = " & ternro
            StrSql = StrSql & " and his_estructura.tenro= " & rutAfp
            StrSql = StrSql & " and " & ConvFecha(pliqdesde) & " >= htetdesde and ((" & ConvFecha(pliqhasta) & " <=htethasta) or ((htethasta is null)))"
            StrSql = StrSql & " and ter_doc.tidnro=1"
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                det_rut_prev = rsConsult!nrodoc
            Else
                Flog.writeline "no tiene un rut para el tipo de estructura afp en este periodo"
            End If
            rsConsult.Close
            
            'busco monto isapre
            sql = " SELECT almonto "
            sql = sql & " FROM acu_liq"
            sql = sql & " WHERE acunro = " & acuIsapre
            sql = sql & " AND cliqnro =  " & cliqnro
            OpenRecordset sql, rsConsult
            If Not rsConsult.EOF Then
                If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                    det_monto_salud = det_monto_salud + 0
                Else
                    det_monto_salud = det_monto_salud + Replace(CDbl(rsConsult!almonto), ",", ".")
                End If
            Else
                Flog.writeline "no hay valor para el acumulador isapre, monto cotizacion salud"
                'acuIsapre = acuIsapre + 0
            End If
            rsConsult.Close
            
            'rut institucional prev de salud
            StrSql = " select * from his_estructura "
            StrSql = StrSql & " inner join estructura on estructura.estrnro=his_estructura.estrnro "
            StrSql = StrSql & " inner join ter_doc on ter_doc.ternro = his_estructura.ternro "
            StrSql = StrSql & " Where his_estructura.ternro = " & ternro
            StrSql = StrSql & " and his_estructura.tenro= 17 "
            StrSql = StrSql & " and " & ConvFecha(pliqdesde) & " >= htetdesde and ((" & ConvFecha(pliqhasta) & " <=htethasta) or ((htethasta is null)))"
            StrSql = StrSql & " and ter_doc.tidnro=1"
            OpenRecordset StrSql, rsConsult
            If Not rsConsult.EOF Then
                det_rut_salud = rsConsult!nrodoc
            Else
                Flog.writeline "no tiene un rut para el tipo de estructura sistema de salud en este periodo"
            End If
            rsConsult.Close
            
            
            'total montos anuales sin actualizar
            
            'busco valor  de la renta total neta pagada sebastian stremel
            sql = " SELECT almonto "
            sql = sql & " FROM acu_liq"
            sql = sql & " WHERE acunro = " & acuRentaTotal
            sql = sql & " AND cliqnro =  " & cliqnro
            OpenRecordset sql, rsConsult
            If Not rsConsult.EOF Then
                If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                    totalneto = totalneto + 0
                Else
                    totalneto = totalneto + CDbl(Replace(rsConsult!almonto, ",", "."))
                End If
            Else
                Flog.writeline "no hay valor para renta total neta pagada"
                
            End If
            rsConsult.Close
            
            
            'busco valor  de la rentas pagadas durante el año
            sql = " SELECT almonto "
            sql = sql & " FROM acu_liq"
            sql = sql & " WHERE acunro = " & acuRentasPagadas
            sql = sql & " AND cliqnro =  " & cliqnro
            OpenRecordset sql, rsConsult
            If Not rsConsult.EOF Then
                If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                    rentasanuales = rentasanuales + 0
                Else
                    rentasanuales = rentasanuales + CDbl(Replace(rsConsult!almonto, ",", "."))
                End If
            Else
                Flog.writeline "no hay valor para rentas pagadas durante el año"
                
            End If
            rsConsult.Close
            
            
            'busco valor  de la rentas accesorias
            sql = " SELECT almonto "
            sql = sql & " FROM acu_liq"
            sql = sql & " WHERE acunro = " & acuRentasAccesorias
            sql = sql & " AND cliqnro =  " & cliqnro
            OpenRecordset sql, rsConsult
            If Not rsConsult.EOF Then
                If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                    rentasaccesorias = rentasaccesorias + 0
                Else
                    rentasaccesorias = rentasaccesorias + CDbl(Replace(rsConsult!almonto, ",", "."))
                End If
            Else
                Flog.writeline "no hay valor para rentas accesorias"
                
            End If
            rsConsult.Close

            'busco valor total imponible
            sql = " SELECT almonto "
            sql = sql & " FROM acu_liq"
            sql = sql & " WHERE acunro = " & acuRemImponible
            sql = sql & " AND cliqnro =  " & cliqnro
            OpenRecordset sql, rsConsult
            If Not rsConsult.EOF Then
                If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                    totalimponible = totalimponible + 0
                    
                Else
                    totalimponible = totalimponible + CDbl(Replace(rsConsult!almonto, ",", "."))
                End If
            Else
                Flog.writeline "no hay valor para total imponible"
                
            End If
            rsConsult.Close
            
            
            'total montos anuales actualizados
            
            'renta total neta pagada
            sql = " SELECT almonto "
            sql = sql & " FROM acu_liq"
            sql = sql & " WHERE acunro = " & acuRentaTotalNeta
            sql = sql & " AND cliqnro =  " & cliqnro
            OpenRecordset sql, rsConsult
            If Not rsConsult.EOF Then
                If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                    l_acuRentaTotalNeta = l_acuRentaTotalNeta + 0
                Else
                    l_acuRentaTotalNeta = l_acuRentaTotalNeta + CDbl(Replace(rsConsult!almonto, ",", "."))
                End If
            Else
                Flog.writeline "no hay valor para renta total neta pagada"
                
            End If
            rsConsult.Close
    
            'impuesto unico retenido
            sql = " SELECT almonto "
            sql = sql & " FROM acu_liq"
            sql = sql & " WHERE acunro = " & acuImpUnicoRet
            sql = sql & " AND cliqnro =  " & cliqnro
            OpenRecordset sql, rsConsult
            If Not rsConsult.EOF Then
                If IsNull(rsConsult!almonto) Or (rsConsult!almonto = "") Then
                    l_acuImpUnicoRetenido = l_acuImpUnicoRetenido + 0
                Else
                    l_acuImpUnicoRetenido = l_acuImpUnicoRetenido + CDbl(Replace(rsConsult!almonto, ",", "."))
                End If
            Else
                Flog.writeline "no hay valor para acumulador impuesto unico retenido, sector actualizados"
                
            End If
            rsConsult.Close

          
            GrabaEmpleado = True
            Flog.writeline "   Se encontraron datos para el empleado en el proceso "
            
  
            
            'agrego esto para grabar
            '------------------------------------------------------------------
            'Armo la SQL para guardar los datos
            '------------------------------------------------------------------

            StrSql = " INSERT INTO rep_formulariof1827_det "
            StrSql = StrSql & " (bpronro, empRut, mes, relContractual, renTotalNetaPagada, impUnicoRet, mayorRetSolicitada,"
            StrSql = StrSql & " montoCotPrevisional, rutPrevisional, montoSalud, rutSalud, capSence, nroCertificado "
            StrSql = StrSql & ")"
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "(" & NroProceso
            StrSql = StrSql & ",'" & det_rut & "'"
            StrSql = StrSql & "," & pliqMes
            StrSql = StrSql & ",'" & det_relcontractual & "'"
            StrSql = StrSql & "," & det_rentatotal
            StrSql = StrSql & "," & det_impunico
            StrSql = StrSql & "," & 0
            StrSql = StrSql & "," & det_monto_cot_prev
            StrSql = StrSql & ",'" & det_rut_prev & "'"
            StrSql = StrSql & "," & det_monto_salud
            StrSql = StrSql & ",'" & det_rut_salud & "'"
            StrSql = StrSql & ",'" & Null & "' "
            StrSql = StrSql & ",'" & Null & "' "
            StrSql = StrSql & ")"
            
            '------------------------------------------------------------------
            'Guardo los datos en la BD
            '------------------------------------------------------------------
            objConn.Execute StrSql, , adExecuteNoRecords
            l_acuTotalEmp = l_acuTotalEmp + 1
            Flog.writeline " Se Grabo el Formulario F1827"
        
                    
    Else
        Flog.writeline "El empleado no se encuentra en el proceso. Nro de proceso: " & proNro
    End If
        




Next
        
StrSql = " UPDATE rep_formulariof1827 set rentaTotalNetaPag= " & totalneto & ",impUnicRet1= " & rentasanuales & ",impUnicRet2=" & rentasaccesorias & ",totalRemImpo=" & totalimponible & ",rentaTotalNetaPag2=" & l_acuRentaTotalNeta & ",impUnicoRetenido=" & l_acuImpUnicoRetenido & ",totalCasosInf=" & l_acuTotalEmp & " where bpronro= " & NroProceso
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline " Se actualizo el Formulario F1827"
  
'If GrabaEmpleado Then
  
    '------------------------------------------------------------------
    'Armo la SQL para guardar los datos
    '------------------------------------------------------------------

 '   StrSql = " INSERT INTO rep_formulario1827_det "
  '  StrSql = StrSql & " (bpronro, empRut, mes, relContractual, renTotalNetaPagada, impUnicoRet, mayorRetSolicitada,"
  '  StrSql = StrSql & " montoCotPrevisional, rutPrevisional, montoSalud, rutSalud, capSence, nroCertificado "
    'StrSql = StrSql & " (bpronro, pronro, prodesc, descripcion, fecha, hora, iduser,"
    'StrSql = StrSql & " empnro, empnom, empdir, emprut, pliqnro, pliqmes, pliqanio"
    'For I = 1 To 40
    '    StrSql = StrSql & "," & " valor" & I
    'Next I
    
   ' StrSql = StrSql & ")"
   ' StrSql = StrSql & " VALUES "
   ' StrSql = StrSql & "(" & NroProceso
   ' StrSql = StrSql & ",'" & Rut & "'"
   ' StrSql = StrSql & ",'" & pliqMes & "'"
   ' StrSql = StrSql & ",'" & relContractual & "'"
   ' StrSql = StrSql & "," & rentaTotal
   ' StrSql = StrSql & "," & acuImpUnicoR
   ' StrSql = StrSql & ",'" & Null & "' "
   ' StrSql = StrSql & "," & acuCotPrev
   ' StrSql = StrSql & ",'" & rutInstitucional & "'"
   ' StrSql = StrSql & "," & acuIsapre
   ' StrSql = StrSql & ",'" & rutSalud & "'"
   ' StrSql = StrSql & ",'" & Null & "' "
   ' StrSql = StrSql & ",'" & Null & "' "
    'StrSql = StrSql & "," & pliqAnio
    'For I = 1 To 40
    '    StrSql = StrSql & "," & VectorValorACCO(I)
    'Next I
   ' StrSql = StrSql & ")"
    
    '------------------------------------------------------------------
    'Guardo los datos en la BD
    '------------------------------------------------------------------
   ' objConn.Execute StrSql, , adExecuteNoRecords
    
    'Flog.writeline " Se Grabo el Formulario F1827"
'End If

Exit Sub

MError:
    Flog.writeline "Error en Formulario F1827: " & NroProceso & " Error: " & Err.Description
    Flog.writeline "Ultimo Sql Ejecutado: " & StrSql
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
    
    'If NroProc > 0 Then
    '    StrEmpl = "SELECT * FROM batch_empleado "
    '    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
    '    StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
    '    StrEmpl = StrEmpl & " ORDER BY progreso,estado"
    'End If
   
    If NroProc > 0 Then
        If tenro3 <> 0 Then
                StrEmpl = "SELECT * FROM batch_empleado "
                StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN his_estructura on his_Estructura.ternro=empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN tipoestructura on tipoestructura.tenro = his_estructura.tenro"
                StrEmpl = StrEmpl & " INNER JOIN estructura on estructura.tenro=tipoestructura.tenro"
                StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
                StrEmpl = StrEmpl & " AND ((his_estructura.htethasta <=" & NuevaFecha & ") or (his_estructura.htethasta is NULL))"
                If estrnro3 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro3 & " And estructura.estrnro = " & estrnro3
                End If
                If estrnro2 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro2 & " And estructura.estrnro = " & estrnro2
                End If
                If estrnro1 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro1 & " And estructura.estrnro = " & estrnro1
                End If
                'StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro1 & " And estructura.estrnro = " & estrnro1
                StrEmpl = StrEmpl & " ORDER BY progreso,estado"
                'StrEmpl = StrEmpl & " ORDER BY " & Ordenamiento & " " & orden2
                   
 
        Else
            If tenro2 <> 0 Then
                StrEmpl = "SELECT * FROM batch_empleado "
                StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN his_estructura on his_Estructura.ternro=empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN tipoestructura on tipoestructura.tenro = his_estructura.tenro"
                StrEmpl = StrEmpl & " INNER JOIN estructura on estructura.tenro=tipoestructura.tenro"
                StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
                StrEmpl = StrEmpl & " AND ((his_estructura.htethasta <=" & NuevaFecha & ") or (his_estructura.htethasta is NULL))"
                If estrnro2 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro2 & " And estructura.estrnro = " & estrnro2
                End If
                If estrnro1 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro1 & " And estructura.estrnro = " & estrnro1
                End If
                StrEmpl = StrEmpl & " ORDER BY progreso,estado"
                'StrEmpl = StrEmpl & " ORDER BY " & Ordenamiento & " " & orden2
            Else
                If tenro1 <> 0 Then
                    StrEmpl = "SELECT * FROM batch_empleado "
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
                    StrEmpl = StrEmpl & " INNER JOIN his_estructura on his_Estructura.ternro=empleado.ternro "
                    StrEmpl = StrEmpl & " INNER JOIN tipoestructura on tipoestructura.tenro = his_estructura.tenro"
                    StrEmpl = StrEmpl & " INNER JOIN estructura on estructura.tenro=tipoestructura.tenro"
                    StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
                    StrEmpl = StrEmpl & " AND ((his_estructura.htethasta <=" & NuevaFecha & ") or (his_estructura.htethasta is NULL))"
                    If estrnro1 <> 0 Then
                        StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro1 & " And estructura.estrnro = " & estrnro1
                    End If
                    StrEmpl = StrEmpl & " ORDER BY progreso,estado"
                    'StrEmpl = StrEmpl & " ORDER BY " & Ordenamiento & " " & orden2
                Else
                    StrEmpl = "SELECT * FROM batch_empleado "
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
                    StrEmpl = StrEmpl & " WHERE bpronro = " & NroProc
                    StrEmpl = StrEmpl & " ORDER BY progreso,estado"
                    'StrEmpl = StrEmpl & " ORDER BY " & Ordenamiento & " " & orden2
                
                End If
     
            End If
        End If
    End If
    OpenRecordset StrEmpl, rsEmpl
    
    cantRegistros = rsEmpl.RecordCount
    CEmpleadosAProc = cantRegistros
    'IncPorc = (99 / CEmpleadosAProc)
    totalEmpleados = cantRegistros
    
End Sub
Sub CargarEmpleados_v1(ByVal NroProc, ByRef rsEmpl As ADODB.Recordset, ByVal empresa As Long)

Dim StrEmpl As String

    'If NroProc > 0 Then
    '   StrEmpl = "select distinct empleado ternro from cabliq where pronro in(" & NroProc & " )"
    'End If
    If NroProc > 0 Then
        If tenro3 <> 0 Then
                StrEmpl = "select distinct empleado ternro from cabliq "
                StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN his_estructura on his_Estructura.ternro=empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN tipoestructura on tipoestructura.tenro = his_estructura.tenro"
                StrEmpl = StrEmpl & " INNER JOIN estructura on estructura.tenro=tipoestructura.tenro"
                StrEmpl = StrEmpl & " WHERE pronro in(" & NroProc & " )"
                StrEmpl = StrEmpl & " AND ((his_estructura.htethasta <=" & NuevaFecha & ") or (his_estructura.htethasta is NULL))"
                If estrnro3 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro3 & " And estructura.estrnro = " & estrnro3
                End If
                If estrnro2 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro2 & " And estructura.estrnro = " & estrnro2
                End If
                If estrnro1 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro1 & " And estructura.estrnro = " & estrnro1
                End If
                'StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro1 & " And estructura.estrnro = " & estrnro1
                StrEmpl = StrEmpl & " ORDER BY progreso,estado"
                'StrEmpl = StrEmpl & " ORDER BY " & Ordenamiento & " " & orden2
 
        Else
            If tenro2 <> 0 Then
                StrEmpl = "select distinct empleado ternro from cabliq "
                StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN his_estructura on his_Estructura.ternro=empleado.ternro "
                StrEmpl = StrEmpl & " INNER JOIN tipoestructura on tipoestructura.tenro = his_estructura.tenro"
                StrEmpl = StrEmpl & " INNER JOIN estructura on estructura.tenro=tipoestructura.tenro"
                StrEmpl = StrEmpl & " WHERE pronro in(" & NroProc & " )"
                StrEmpl = StrEmpl & " AND ((his_estructura.htethasta <=" & NuevaFecha & ") or (his_estructura.htethasta is NULL))"
                If estrnro2 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro2 & " And estructura.estrnro = " & estrnro2
                End If
                If estrnro1 <> 0 Then
                    StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro1 & " And estructura.estrnro = " & estrnro1
                End If
                StrEmpl = StrEmpl & " ORDER BY progreso,estado"
                'StrEmpl = StrEmpl & " ORDER BY " & Ordenamiento & " " & orden2
            Else
                If tenro1 <> 0 Then
                    StrEmpl = "select distinct empleado ternro from cabliq "
                    StrEmpl = StrEmpl & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
                    StrEmpl = StrEmpl & " INNER JOIN his_estructura on his_Estructura.ternro=empleado.ternro "
                    StrEmpl = StrEmpl & " INNER JOIN tipoestructura on tipoestructura.tenro = his_estructura.tenro"
                    StrEmpl = StrEmpl & " INNER JOIN estructura on estructura.tenro=tipoestructura.tenro"
                    StrEmpl = StrEmpl & " WHERE pronro in(" & NroProc & " )"
                    StrEmpl = StrEmpl & " AND ((his_estructura.htethasta <=" & NuevaFecha & ") or (his_estructura.htethasta is NULL))"
                    If estrnro1 <> 0 Then
                        StrEmpl = StrEmpl & " AND tipoestructura.tenro =" & tenro1 & " And estructura.estrnro = " & estrnro1
                    End If
                    StrEmpl = StrEmpl & " ORDER BY progreso,estado"
                    'StrEmpl = StrEmpl & " ORDER BY " & Ordenamiento & " " & orden2
                Else
                    StrEmpl = "select distinct empleado ternro from cabliq where pronro in(" & NroProc & " )"
                End If
           End If
     
            'End If
        End If
    End If
   
    OpenRecordset StrEmpl, rsEmpl
    
    cantRegistros = rsEmpl.RecordCount
    CEmpleadosAProc = cantRegistros
    'IncPorc = (99 / CEmpleadosAProc)
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


Sub buscarDatosEmpresa(Empnroestr)

Dim StrSql As String
Dim rsConsult As New ADODB.Recordset



    empresa = ""
    emprTer = 0
    emprCuit = ""
    emprDire = ""
    

    ' -------------------------------------------------------------------------
    'Busco los datos Basicos de la Empresa
    ' -------------------------------------------------------------------------
    Flog.writeline "Buscando datos de la empresa"
    
    StrSql = "SELECT * FROM empresa WHERE Estrnro = " & Empnroestr
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
       Flog.writeline "Error: Buscando datos de la empresa: al obtener el empleado"
       HuboErrores = True
    Else
        empresa = rsConsult!empNom
        emprTer = rsConsult!ternro
        Empnro = rsConsult!Empnro
    End If
    
    rsConsult.Close
            
    'Consulta para obtener el RUT de la empresa
    StrSql = "SELECT nrodoc FROM tercero " & _
             " INNER JOIN ter_doc ON (tercero.ternro = ter_doc .ternro and ter_doc.tidnro = 1)" & _
             " Where tercero.ternro =" & emprTer
    
    Flog.writeline "Buscando datos del RUT de la empresa"
    
    OpenRecordset StrSql, rsConsult
    
    If rsConsult.EOF Then
        Flog.writeline "No se encontró el RUT de la Empresa"
        emprCuit = "  "
    Else
        emprCuit = rsConsult!nrodoc
    End If
    rsConsult.Close
End Sub

