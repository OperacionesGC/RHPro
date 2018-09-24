Attribute VB_Name = "mdlRepAguinaldo"
'Global Const Version = "1.00"
'Global Const FechaModificacion = "19/11/2012"
'Global Const UltimaModificacion = " Sebastian Stremel "

'Global Const Version = "1.01"
'Global Const FechaModificacion = "28/11/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - correccion de errores  - CAS-16938 - Sykes - Nuevo Comprobante de Aguinaldo - 2da parte "

'Global Const Version = "1.02"
'Global Const FechaModificacion = "23/01/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - Se agrego los conceptos imprimibles  - CAS-18106 - Sykes - Agregado Comprobante de Aguinaldo "

'Global Const Version = "1.03"
'Global Const FechaModificacion = "30/01/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - Se corrigio la query que muestra los conceptos imprimibles  - CAS-18106 - Sykes - Agregado Comprobante de Aguinaldo "

'Global Const Version = "1.04"
'Global Const FechaModificacion = "22/04/2014"
'Global Const UltimaModificacion = " CAS-22452 - SYKES CR - Custom reporte comprobante de aguinaldos - LED - Se analizan solo los procesos que caen dentro de la fase activa del empleado "

'Global Const Version = "1.05"
'Global Const FechaModificacion = "24/04/2014"
'Global Const UltimaModificacion = " CAS-22452 - SYKES CR - Custom asunto mail comprobante aguinaldo - LED - Se analizan solo los empleados que tienen fases activas "

'Global Const Version = "1.06"
'Global Const FechaModificacion = "24/07/2014"
'Global Const UltimaModificacion = " LED - CAS-22452 - SYKES CR - CUSTOM COMPROBANTE DE AGUINALDO MODIFICACION DE FECHAS - Se obtienen la fecha desde y hasta del campo profecpago, antes se obtenia desde profecini "

'Global Const Version = "1.07"
'Global Const FechaModificacion = "04/11/2014"
'Global Const UltimaModificacion = " Sebastian Stremel - Se agregan 4 nuevos AC/CO - CAS-22452 - SYKES CR - CUSTOM DETALLE Y TOTALES DE COMPROBANTE DE AGUINALDO"

'Global Const Version = "1.08"
'Global Const FechaModificacion = "04/11/2014"
'Global Const UltimaModificacion = " Sebastian Stremel - Se corrigio IN (listaprocesos) ya que tenia comillas en algunos casos y rompia - CAS-22452 - SYKES CR - CUSTOM DETALLE Y TOTALES DE COMPROBANTE DE AGUINALDO [Entrega 2]"

Global Const Version = "1.09"
Global Const FechaModificacion = "18/09/2015"
Global Const UltimaModificacion = " LED - CAS-31849 - SYKES CR - Custom comprobando de pago de aguinaldo - Se resta al Aguinaldo neto los valores de Adelanto de Aguinaldo y pensión alimentaria"


Global ActId As String
Global centrocosto As String
Global neto As String
Global aniodesde As Integer
Global anioactual As Integer
Global arregloProcesos()
Global esConceptoAcu1 As Boolean
Global acu1 As String
Global pos As Integer
Global nroproceso As Long
Global fecEstr As Date
Global orden As Integer
Global total As Integer
Global esConceptoTotal As Boolean
Global lic1 As String
Global lic2 As String
Global lic3 As String
Global lic4 As String
Global esConceptoLic1 As Boolean
Global esConceptoLic2 As Boolean
Global esConceptoLic3 As Boolean
Global esConceptoLic4 As Boolean
Global valorLic1 As Double
Global valorLic2 As Double
Global valorLic3 As Double
Global valorLic4 As Double
Global etiquetaLic1 As String
Global etiquetaLic2 As String
Global etiquetaLic3 As String
Global etiquetaLic4 As String
Global valorTotal As Double
Global valorNeto As Double
Global valorSalBruto As Double
Global valorPensionAlimentaria As Double
Global valorAdelantoAguinaldo As Double
Global valorAguinaldoNeto As Double
Global listapronro


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
Dim objRs4 As New ADODB.Recordset
Dim fechadesde
Dim fechahasta
Dim tipoDepuracion
Dim historico As Boolean
Dim param
'Dim listapronro
Dim Pronro
Dim Ternro
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim acunroSueldo
Dim i
Dim totalEmpleados
Dim cantRegistros
Dim PID As String
Dim tituloReporte As String
Dim parametros As String
Dim ArrParametros
Dim strTempo As String

Dim pliqdesde
Dim pliahasta
Dim cant As Integer

Dim codSalBruto
Dim esConceptoSalBruto As Boolean

Dim codPensionAlimentaria
Dim esConceptoPensionAlimentaria As Boolean

Dim codAdelantoAguinaldo
Dim esConceptoAdelantoAguinaldo As Boolean

Dim codAguinaldoNeto
Dim esConceptoAguinaldoNeto As Boolean


cant = 0
    strCmdLine = Command()
    
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProcesoBatch = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProcesoBatch = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If

    
    nroproceso = NroProcesoBatch
    

    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    tituloReporte = ""

    TiempoInicialProceso = GetTickCount

    Nombre_Arch = PathFLog & "DetalleAguinaldo" & "-" & nroproceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    Flog.writeline "Inicio Proceso de Recibos de Sueldo : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
   
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline

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

    HuboErrores = False
    On Error GoTo CE
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & nroproceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        cantRegistros = CLng(objRs!bprcempleados)
        totalEmpleados = cantRegistros
    
    End If
    objRs.Close
   
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & nroproceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & nroproceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       'Obtengo los parametros del proceso
       parametros = objRs!bprcparam
       
       Flog.writeline "Parametros del proceso: " & parametros
       
       ArrParametros = Split(parametros, "@")
       
       'Obtengo la lista de procesos
       listapronro = ArrParametros(0)
       
       tipoRecibo = CLng(ArrParametros(1))
       
       tenro1 = CLng(ArrParametros(2))
       estrnro1 = CLng(ArrParametros(3))
       tenro2 = CLng(ArrParametros(4))
       estrnro2 = CLng(ArrParametros(5))
       tenro3 = CLng(ArrParametros(6))
       estrnro3 = CLng(ArrParametros(7))
       fecEstr = ArrParametros(8)
       
       
       'pliqdesde
       pliqdesde = ArrParametros(9)
       
       'pliqhasta
       pliqhasta = ArrParametros(10)
       
       'Armo el titulo del reporte
       strTempo = ArrParametros(11)
       
       
       ArrParametros = Split(strTempo, "<br>")
       If UBound(ArrParametros) >= 1 Then
          tituloReporte = ArrParametros(1)
       Else
          tituloReporte = ""
       End If
       
       ArrParametros = Split(ArrParametros(0), "- Periodos")
       tituloReporte = ArrParametros(0) & tituloReporte
       
       'EMPIEZA EL PROCESO
       
       'Busco la configuracion del confrep
       Flog.writeline "Obtengo los datos del confrep"
       
       StrSql = " SELECT * FROM confrep "
       StrSql = StrSql & " WHERE repnro = 389"
       OpenRecordset StrSql, objRs2
       If objRs2.EOF Then
          Flog.writeline "No esta configurado el ConfRep"
          Exit Sub
       End If
       
       Do Until objRs2.EOF
          Flog.writeline "Columna " & objRs2!confnrocol
          Select Case objRs2!confnrocol
             Case 1
                'Cod Ext. Estructura Act. ID de la persona a la cual corresponde el recibo (57 - configurable)
                If objRs2!conftipo = "TE" Then
                    ActId = objRs2!confval
                    Flog.writeline "Codigo a buscar de estructura act id:" & ActId
                Else
                    Flog.writeline "La columna 1 debe ser configurada como TE"
                End If
             
             Case 2
                'centro de costo
                If objRs2!conftipo = "TE" Then
                    centrocosto = objRs2!confval
                    Flog.writeline "Codigo a buscar de estructura centro de costo:" & centrocosto
                Else
                    Flog.writeline "La columna 2 debe ser configurada como TE"
                End If
                
            Case 3
                'acumulador neto
                If objRs2!conftipo = "AC" Then
                    neto = objRs2!confval
                    esConceptoNeto = False
                    Flog.writeline "Codigo a buscar del neto:" & neto
                Else
                    If objRs2!conftipo = "CO" Then
                        neto = objRs2!confval2
                        esConceptoNeto = True
                        Flog.writeline "Se busca el concepto: " & neto
                    Else
                        Flog.writeline "La columna 3 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 4
                'acumulador o concepto que busca los valores
                If objRs2!conftipo = "AC" Then
                    acu1 = objRs2!confval
                    esConceptoAcu1 = False
                    Flog.writeline "Codigo a buscar del acumulador 1:" & acu1
                Else
                    If objRs2!conftipo = "CO" Then
                        acu1 = objRs2!confval2
                        esConceptoAcu1 = True
                        Flog.writeline "Se busca el concepto: " & acu1
                    Else
                        Flog.writeline "La columna 4 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 5
                'acumulador o concepto que busca los valores
                If objRs2!conftipo = "AC" Then
                    lic1 = objRs2!confval
                    etiquetaLic1 = objRs2!confetiq
                    esConceptoLic1 = False
                    Flog.writeline "Codigo a buscar de licencia1:" & lic1
                Else
                    If objRs2!conftipo = "CO" Then
                        lic1 = objRs2!confval2
                        etiquetaLic1 = objRs2!confetiq
                        esConceptoLic1 = True
                        Flog.writeline "Se busca el concepto: " & lic1
                    Else
                        Flog.writeline "La columna 5 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 6
                'acumulador o concepto que busca los valores
                If objRs2!conftipo = "AC" Then
                    lic2 = objRs2!confval
                    etiquetaLic2 = objRs2!confetiq
                    esConceptoLic2 = False
                    Flog.writeline "Codigo a buscar de licencia2:" & lic2
                Else
                    If objRs2!conftipo = "CO" Then
                        lic2 = objRs2!confval2
                        etiquetaLic2 = objRs2!confetiq
                        esConceptoLic2 = True
                        Flog.writeline "Se busca el concepto: " & lic2
                    Else
                        Flog.writeline "La columna 6 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 7
                'acumulador o concepto que busca los valores
                If objRs2!conftipo = "AC" Then
                    lic3 = objRs2!confval
                    etiquetaLic3 = objRs2!confetiq
                    esConceptoLic3 = False
                    Flog.writeline "Codigo a buscar de licencia3:" & lic3
                Else
                    If objRs2!conftipo = "CO" Then
                        lic3 = objRs2!confval2
                        etiquetaLic3 = objRs2!confetiq
                        esConceptoLic3 = True
                        Flog.writeline "Se busca el concepto: " & lic3
                    Else
                        Flog.writeline "La columna 7 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 8
                'acumulador o concepto que busca los valores
                If objRs2!conftipo = "AC" Then
                    lic4 = objRs2!confval
                    etiquetaLic4 = objRs2!confetiq
                    esConceptoLic4 = False
                    Flog.writeline "Codigo a buscar de licencia4:" & lic4
                Else
                    If objRs2!conftipo = "CO" Then
                        lic4 = objRs2!confval2
                        etiquetaLic4 = objRs2!confetiq
                        esConceptoLic4 = True
                        Flog.writeline "Se busca el concepto: " & lic4
                    Else
                        Flog.writeline "La columna 8 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 9
                'acumulador o concepto que busca los valores
                If objRs2!conftipo = "AC" Then
                    total = objRs2!confval
                    esConceptoTotal = False
                    Flog.writeline "Codigo a buscar del total:" & total
                Else
                    If objRs2!conftipo = "CO" Then
                        total = objRs2!confval2
                        esConceptoTotal = True
                        Flog.writeline "Se busca el concepto: " & total
                    Else
                        Flog.writeline "La columna 9 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 10
                'total salario bruto
                 If objRs2!conftipo = "AC" Then
                    codSalBruto = objRs2!confval
                    esConceptoSalBruto = False
                    Flog.writeline "Codigo a buscar del salario Bruto:" & codSalBruto
                Else
                    If objRs2!conftipo = "CO" Then
                        codSalBruto = objRs2!confval2
                        esConceptoSalBruto = True
                        Flog.writeline "Se busca el concepto: " & codSalBruto
                    Else
                        Flog.writeline "La columna 10 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 11
                'Pension Alimentaria
                 If objRs2!conftipo = "AC" Then
                    codPensionAlimentaria = objRs2!confval
                    esConceptoPensionAlimentaria = False
                    Flog.writeline "Codigo a buscar de la pension alimentaria:" & codPensionAlimentaria
                Else
                    If objRs2!conftipo = "CO" Then
                        codPensionAlimentaria = objRs2!confval2
                        esConceptoPensionAlimentaria = True
                        Flog.writeline "Se busca el concepto: " & codPensionAlimentaria
                    Else
                        Flog.writeline "La columna 11 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 12
                'Adelanto de Aguinaldo
                 If objRs2!conftipo = "AC" Then
                    codAdelantoAguinaldo = objRs2!confval
                    esConceptoAdelantoAguinaldo = False
                    Flog.writeline "Codigo a buscar de adelanto de aguinaldo:" & codAdelantoAguinaldo
                Else
                    If objRs2!conftipo = "CO" Then
                        codAdelantoAguinaldo = objRs2!confval2
                        esConceptoAdelantoAguinaldo = True
                        Flog.writeline "Se busca el concepto: " & codAdelantoAguinaldo
                    Else
                        Flog.writeline "La columna 12 debe ser un acumulador o un concepto."
                    End If
                End If
            Case 13
                'Aguinaldo Neto
                 If objRs2!conftipo = "AC" Then
                    codAguinaldoNeto = objRs2!confval
                    esConceptoAguinaldoNeto = False
                    Flog.writeline "Codigo a buscar de aguinaldo neto:" & codAguinaldoNeto
                Else
                    If objRs2!conftipo = "CO" Then
                        codAguinaldoNeto = objRs2!confval2
                        esConceptoAguinaldoNeto = True
                        Flog.writeline "Se busca el concepto: " & codAguinaldoNeto
                    Else
                        Flog.writeline "La columna 13 debe ser un acumulador o un concepto."
                    End If
                End If
          End Select
       
          objRs2.MoveNext
       Loop
       
       objRs2.Close
       
        'busco el año del proceso hasta
        StrSql = " SELECT pliqanio FROM periodo "
        StrSql = StrSql & " WHERE pliqnro=" & pliqhasta
        OpenRecordset StrSql, objRs2
        If Not objRs2.EOF Then
            anioactual = objRs2!pliqanio
            Flog.writeline "Se encontro el año actual"
        Else
            Flog.writeline "No se encontro el año actual"
        End If
        objRs2.Close
        'hasta aca
       

       
        aniodesde = anioactual - 1
        'busco los periodos de un año para atras
        StrSql = " SELECT * FROM periodo "
        StrSql = StrSql & " WHERE (pliqmes = 12 And pliqanio =" & aniodesde & ") "
        StrSql = StrSql & " OR (pliqmes <=11 and pliqanio=" & anioactual & ")"
        StrSql = StrSql & " ORDER BY pliqdesde  ASC"
        OpenRecordset StrSql, objRs2
        Do While Not objRs2.EOF
            'para cada periodo busco las liquidaciones semanales
            StrSql = " SELECT distinct proceso.pronro, prodesc,proceso.pliqnro,tprocnro, profecini, profecfin, periodo.pliqanio, periodo.pliqmes, proceso.profecpago FROM proceso "
            StrSql = StrSql & " INNER JOIN cabliq on cabliq.pronro= proceso.pronro "
            StrSql = StrSql & " INNER JOIN periodo on periodo.pliqnro = proceso.pliqnro "
            StrSql = StrSql & " WHERE proceso.pliqnro=" & objRs2!pliqnro & " AND tprocnro in(1,15)" '1 y 15 son los procesos semanales
            StrSql = StrSql & " ORDER BY profecpago asc"
            OpenRecordset StrSql, objRs4

            Do While Not objRs4.EOF
                'armar arreglo con todas las liquidaciones a evaluar
                cant = cant + 1
                ReDim Preserve arregloProcesos(cant)
                arregloProcesos(cant) = objRs4!Pronro & "@" & objRs4!prodesc & "@" & objRs4!pliqnro
                'hasta aca
            objRs4.MoveNext
            Loop
            objRs4.Close
        objRs2.MoveNext
        Loop
        objRs2.Close
        'hasta aca
       
        If cant = 0 Then
            Flog.writeline "no hay procesos"
            StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & nroproceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            GoTo FIN
        End If
        
       'Obtengo los empleados sobre los que tengo que generar los recibos
       Flog.writeline "Obtengo los empleados sobre los que tengo que generar los recibos"
       Call CargarEmpleados(nroproceso, rsEmpl)
       
       Flog.writeline "Inicializo progreso"
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & nroproceso
       objconnProgreso.Execute StrSql, , adExecuteNoRecords
       
       orden = 0
       
       Flog.writeline "Genero por cada empleado un recibo de sueldo"
       'Genero por cada empleado un recibo de sueldo
       If rsEmpl.RecordCount <= 0 Then
            Flog.writeline "No hay empleados para procesar"
            StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & nroproceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            GoTo FIN
       End If
       
       Do Until rsEmpl.EOF
          
       '__________________________________________________________________
       'BUSCO LOS ACUMULADORES DE LA LIQUIDACION SELECCIONADA EN EL FILTRO
       Ternro = rsEmpl!Ternro
       
       'Salario Bruto seba 30/10/2014
        If EsNulo(codSalBruto) Then
            Flog.writeline "El concepto/acumulador salario bruto es nulo"
        Else
            If esConceptoSalBruto Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & codSalBruto & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorSalBruto = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & codSalBruto
                    Else
                        valorSalBruto = 0
                    End If
                Else
                    valorSalBruto = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & codSalBruto
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorSalBruto = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & codSalBruto
                    Else
                        valorSalBruto = 0
                    End If
                Else
                    valorSalBruto = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
            objRs2.Close
       End If
       'hasta aca
       
       'Pension Alimentaria
        If EsNulo(codPensionAlimentaria) Then
            Flog.writeline "El concepto/acumulador pension alimentaria es nulo"
        Else
            If esConceptoPensionAlimentaria Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & codPensionAlimentaria & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorPensionAlimentaria = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & codPensionAlimentaria
                    Else
                        valorPensionAlimentaria = 0
                    End If
                Else
                    valorPensionAlimentaria = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & codPensionAlimentaria
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorPensionAlimentaria = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & codPensionAlimentaria
                    Else
                        valorPensionAlimentaria = 0
                    End If
                Else
                    valorPensionAlimentaria = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
            objRs2.Close
       End If
       'hasta aca
       
       'Adelanto de aguinaldo
        If EsNulo(codAdelantoAguinaldo) Then
            Flog.writeline "El concepto/acumulador adelanto aguinaldo es nulo"
        Else
            If esConceptoAdelantoAguinaldo Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & codAdelantoAguinaldo & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorAdelantoAguinaldo = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & codAdelantoAguinaldo
                    Else
                        valorAdelantoAguinaldo = 0
                    End If
                Else
                    valorAdelantoAguinaldo = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & codAdelantoAguinaldo
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorAdelantoAguinaldo = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & codAdelantoAguinaldo
                    Else
                        valorAdelantoAguinaldo = 0
                    End If
                Else
                    valorAdelantoAguinaldo = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
            objRs2.Close
       End If
       'hasta aca
       
       
       'Aguinaldo neto
        If EsNulo(codAguinaldoNeto) Then
            Flog.writeline "El concepto/acumulador aguinaldo neto es nulo"
        Else
            If esConceptoAguinaldoNeto Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & codAguinaldoNeto & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorAguinaldoNeto = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & codAguinaldoNeto
                    Else
                        valorAguinaldoNeto = 0
                    End If
                Else
                    valorAguinaldoNeto = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & codAguinaldoNeto
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorAguinaldoNeto = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & codAguinaldoNeto
                    Else
                        valorAguinaldoNeto = 0
                    End If
                Else
                    valorAguinaldoNeto = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
            objRs2.Close
            
            'LED (v1.09) - se resta al valor del aguinaldo neto la pension alimentaria y adelanto de aguinaldo
            If Not EsNulo(codAguinaldoNeto) Then
                valorAguinaldoNeto = valorAguinaldoNeto - Abs(valorPensionAlimentaria)
            End If
            
            If Not EsNulo(valorAdelantoAguinaldo) Then
                valorAguinaldoNeto = valorAguinaldoNeto - Abs(valorAdelantoAguinaldo)
            End If
            'Fin - LED (v1.09) - se resta al valor del aguinaldo neto la pension alimentaria y adelanto de aguinaldo
       End If
       'hasta aca
       
       'ACUMULADOR TOTAL
       
        If EsNulo(total) Then
            Flog.writeline "El concepto/acumulador total es nulo"
        Else
            If esConceptoTotal Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & total & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorTotal = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & total
                    Else
                        valorTotal = 0
                    End If
                Else
                    valorTotal = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & total
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorTotal = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & acu1
                    Else
                        valorTotal = 0
                    End If
                Else
                    valorTotal = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
            objRs2.Close
       End If
       
       
       
       'ACUMULADOR AGUINALDO O NETO
        If EsNulo(neto) Then
            Flog.writeline "El concepto/acumulador neto es nulo"
        Else
            If esConceptoNeto Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & neto & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorNeto = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & neto
                    Else
                        valorNeto = 0
                    End If
                Else
                    valorNeto = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & neto
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorNeto = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & neto
                    Else
                        valorNeto = 0
                    End If
                Else
                    valorNeto = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
            objRs2.Close
        End If
        
        
       'ACUMULADOR LIC1
        If EsNulo(lic1) Then
            Flog.writeline "El concepto/acumulador lic1 es nulo"
        Else
            If esConceptoLic1 Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & lic1 & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorLic1 = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & lic1
                    Else
                        valorLic1 = 0
                    End If
                Else
                    valorLic1 = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto)monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & lic1
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorLic1 = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & lic1
                    Else
                        valorLic1 = 0
                    End If
                Else
                    valorLic1 = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
            objRs2.Close
        End If
        
       'ACUMULADOR LIC2
        If EsNulo(lic2) Then
            Flog.writeline "El concepto/acumulador neto es nulo"
        Else
            If esConceptoLic2 Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & lic2 & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorLic2 = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & lic2
                    Else
                        valorLic2 = 0
                    End If
                Else
                    valorLic2 = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & lic2
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorLic2 = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & lic2
                    Else
                        valorLic2 = 0
                    End If
                Else
                    valorLic2 = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
            objRs2.Close
        End If
        
        
        
        'ACUMULADOR LIC3
        If EsNulo(lic3) Then
            Flog.writeline "El concepto/acumulador lic3 es nulo"
        Else
            If esConceptoLic3 Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & lic3 & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorLic3 = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & lic3
                    Else
                        valorLic3 = 0
                    End If
                Else
                    valorLic3 = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & lic3
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorLic3 = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & lic3
                    Else
                        valorLic3 = 0
                    End If
                Else
                    valorLic3 = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
           objRs2.Close
        End If
       
        'ACUMULADOR LIC4
        If EsNulo(lic4) Then
            Flog.writeline "El concepto/acumulador lic4 es nulo"
        Else
            If esConceptoLic4 Then
                StrSql = "SELECT sum(dlimonto) monto FROM cabliq"
                StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
                StrSql = StrSql & " WHERE concepto.conccod = '" & lic4 & "'"
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorLic4 = objRs2!Monto
                        Flog.writeline "Se encontraron datos del concepto:" & lic4
                    Else
                        valorLic4 = 0
                    End If
                Else
                    valorLic4 = 0
                    Flog.writeline "No hay datos del concepto."
                End If
            Else
                StrSql = "SELECT sum(almonto) monto  FROM cabliq"
                StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
                StrSql = StrSql & " WHERE acu_liq.acuNro = " & lic4
                StrSql = StrSql & " AND cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
                OpenRecordset StrSql, objRs2
                If Not objRs2.EOF Then
                    If Not EsNulo(objRs2!Monto) Then
                        valorLic4 = objRs2!Monto
                        Flog.writeline "Se encontraron datos del acumulador:" & lic4
                    Else
                        valorLic4 = 0
                    End If
                Else
                    valorLic4 = 0
                    Flog.writeline "No hay datos del acumulador."
                End If
            End If
           objRs2.Close
        End If
       '__________________________________________________________________
          
          
          'orden = orden + 1
          Flog.writeline "Lista de procesos " & listapronro
          Flog.writeline "tercero " & rsEmpl!Ternro
          
          arrpronro = Split(listapronro, ",")
          EmpErrores = False
          'Ternro = rsEmpl!Ternro
          generoRecibo = False
                    
          Call generarReciboAguinaldo(Ternro)

          
             
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
           
        cantRegistros = cantRegistros - 1
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                    ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                    ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & nroproceso
           
        objConn.Execute StrSql, , adExecuteNoRecords
         
        'Si se generaron todos los recibos de sueldo del empleado correctamente lo borro
        If Not EmpErrores Then
              StrSql = " DELETE FROM batch_empleado "
              StrSql = StrSql & " WHERE bpronro = " & nroproceso
              StrSql = StrSql & " AND ternro = " & Ternro
    
              objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rsEmpl.MoveNext
    Loop
Else
    Exit Sub
End If
   
    If Not HuboErrores Then
        'Actualizo el estado del proceso
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & nroproceso
    Else
        StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & nroproceso
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
FIN:
    HuboErrores = False
    Flog.writeline "************************************************************"
    Flog.writeline " Fin: " & Now
    Flog.writeline "************************************************************"
    Exit Sub
CE:
    HuboErrores = True
    Flog.writeline "************************************************************"
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo sql ejecutado: " & StrSql
    Flog.writeline "************************************************************"
End Sub

'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset)

Dim StrEmpl As String

    StrEmpl = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProc & " ORDER BY progreso,estado "
    OpenRecordset StrEmpl, rsEmpl
End Sub

Sub generarReciboAguinaldo(ByVal Ternro)
Dim rsconsult As New ADODB.Recordset
Dim rsconsult2 As New ADODB.Recordset

'variables para el insert
Dim id As String
Dim ccosto As String
Dim empleg As Long
Dim proc As Integer
Dim descProc As String
Dim valorAcu1 As Double
Dim primeravez As Boolean
Dim pliq As Integer
Dim mesAnt As Integer
Dim anioAnt As Integer
Dim apellido As String
Dim Nombre As String
Dim procFechaDesde As Date
Dim procFechaHasta As Date
Dim profecpago As Date
Dim altfec As String
Dim bajfec As String

Dim sumaConceptos As Double
Dim listaNueva As String


apellido = ""
Nombre = ""
'busco el codigo externo de la estructura configurada
StrSql = " SELECT estrcodext, estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fecEstr) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(fecEstr) & ") And his_estructura.tenro =" & ActId & " And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsconsult

If Not rsconsult.EOF Then
    If Not EsNulo(rsconsult!estrcodext) Then
        id = rsconsult!estrcodext
        Flog.writeline "ID: " & id
    Else
        id = ""
        Flog.writeline "El codext del id esta vacio"
    End If
Else
   Flog.writeline "Error al obtener los datos del id"
'   GoTo MError
End If
rsconsult.Close

'busco el centro de costo
StrSql = " SELECT estrcodext, estrdabr "
StrSql = StrSql & " From his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro=his_estructura.estrnro "
StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fecEstr) & " And (htethasta Is Null Or htethasta >= " & ConvFecha(fecEstr) & ") And his_estructura.tenro =" & centrocosto & " And his_estructura.ternro = " & Ternro
       
OpenRecordset StrSql, rsconsult

If Not rsconsult.EOF Then
    If Not EsNulo(rsconsult!estrdabr) Then
        ccosto = rsconsult!estrdabr
        Flog.writeline "centro de costo: " & ccosto
    Else
        ccosto = ""
        Flog.writeline "El centro de costo esta vacio."
    End If
Else
   Flog.writeline "Error al obtener los datos del centro de costo."
'   GoTo MError
End If
rsconsult.Close

'busco el legajo
StrSql = "SELECT empleg, terape, terape2, ternom, ternom2 FROM empleado WHERE ternro=" & Ternro
OpenRecordset StrSql, rsconsult
If Not rsconsult.EOF Then
    empleg = rsconsult!empleg
    Flog.writeline "Legajo del empleado:" & rsconsult!empleg
    If Not EsNulo(rsconsult!ternom2) Then
        Nombre = rsconsult!ternom & " " & rsconsult!ternom2
    Else
        Nombre = rsconsult!ternom
    End If
    
    If Not EsNulo(rsconsult!terape2) Then
        apellido = rsconsult!terape & " " & rsconsult!terape2
    Else
        apellido = rsconsult!terape
    End If
    
Else
    Flog.writeline "No se encontro el legajo del empleado"
End If
rsconsult.Close


'busco la fase activa del empleado
StrSql = "SELECT altfec, bajfec FROM fases WHERE empleado = " & Ternro & " AND estado = -1 "
OpenRecordset StrSql, rsconsult
If Not rsconsult.EOF Then
    altfec = rsconsult!altfec
    If Not EsNulo(rsconsult!bajfec) Then
        bajfec = rsconsult!bajfec
    Else
        bajfec = ""
    End If
    Flog.writeline "El empleado posee fase activa ternro: " & Ternro
Else
    Flog.writeline "El empleado no posee fase activa, no se genera recibo para el ternro: " & Ternro
    Exit Sub
End If
rsconsult.Close

'para el empleado voy a buscar todas las liquidaciones
mesAnt = 0
anioAnt = 0
orden = orden + 1
'inserto en la tabla cabecera
StrSql = " INSERT INTO rep_aguinaldo "
StrSql = StrSql & "(bpronro, ternro, idpersona, legajo "
StrSql = StrSql & ", apellido, nombre, centrocosto,pronro, auxdeci1, auxdeci2, auxdeci3, auxdeci4 "
StrSql = StrSql & ", auxchar1, auxchar2, auxchar3, auxchar4,total, aguinaldo, auxdeci5, auxdeci6, auxdeci7, auxdeci8)"
StrSql = StrSql & " VALUES "
StrSql = StrSql & "("
StrSql = StrSql & nroproceso
StrSql = StrSql & ", " & Ternro
StrSql = StrSql & ", '" & id & "'"
StrSql = StrSql & ", " & empleg
StrSql = StrSql & ", '" & apellido & "'"
StrSql = StrSql & ", '" & Nombre & "'"
StrSql = StrSql & ", '" & ccosto & "'"
StrSql = StrSql & ", " & orden
StrSql = StrSql & ", " & Replace(valorLic1, ",", ".")
StrSql = StrSql & ", " & Replace(valorLic2, ",", ".")
StrSql = StrSql & ", " & Replace(valorLic3, ",", ".")
StrSql = StrSql & ", " & Replace(valorLic4, ",", ".")
StrSql = StrSql & ", '" & etiquetaLic1 & "'"
StrSql = StrSql & ", '" & etiquetaLic2 & "'"
StrSql = StrSql & ", '" & etiquetaLic3 & "'"
StrSql = StrSql & ", '" & etiquetaLic4 & "'"
StrSql = StrSql & ", " & Replace(valorTotal, ",", ".")
StrSql = StrSql & ", " & Replace(valorNeto, ",", ".")
StrSql = StrSql & ", " & Replace(valorSalBruto, ",", ".")
StrSql = StrSql & ", " & Replace(valorPensionAlimentaria, ",", ".")
StrSql = StrSql & ", " & Replace(valorAdelantoAguinaldo, ",", ".")
StrSql = StrSql & ", " & Replace(valorAguinaldoNeto, ",", ".")
StrSql = StrSql & ")"
objconnProgreso.Execute StrSql, , adExecuteNoRecords

sumaConceptos = 0
listaNueva = "0"

Call filtrarProcFases(altfec, bajfec, arregloProcesos)

For k = 1 To UBound(arregloProcesos)


    cadena = Split(arregloProcesos(k), "@")
    proc = cadena(0)
    descProc = cadena(1)
    pliq = cadena(2)
    
    'busco la fecha de pago del proceso
    StrSql = "SELECT profecpago FROM proceso "
    StrSql = StrSql & "WHERE pronro =" & proc
    OpenRecordset StrSql, rsconsult
    If Not rsconsult.EOF Then
        profecpago = rsconsult!profecpago
    End If
    'hasta aca
    
    If k = 1 Then
        'busco la fecha desde de el primer proceso
            StrSql = "SELECT profecpago FROM proceso "
            StrSql = StrSql & "WHERE pronro =" & proc
            OpenRecordset StrSql, rsconsult
            If Not rsconsult.EOF Then
                procFechaDesde = rsconsult!profecpago
            End If
        'hasta aca
        'actualizo algunos datos de la tabla cabecera
        StrSql = " UPDATE rep_aguinaldo "
        StrSql = StrSql & " SET fechaDesde=" & ConvFecha(procFechaDesde)
        StrSql = StrSql & " WHERE bpronro=" & nroproceso & " AND ternro=" & Ternro
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        'hasta aca
    End If
    
    If k = UBound(arregloProcesos) Then
        'busco la fecha desde del ultimo proceso
            StrSql = "SELECT profecpago FROM proceso "
            StrSql = StrSql & "WHERE pronro =" & proc
            OpenRecordset StrSql, rsconsult
            If Not rsconsult.EOF Then
                procFechaHasta = rsconsult!profecpago
            End If
            'actualizo algunos datos de la tabla cabecera
            StrSql = " UPDATE rep_aguinaldo "
            StrSql = StrSql & " SET fechahasta = " & ConvFecha(procFechaHasta)
            StrSql = StrSql & " WHERE bpronro=" & nroproceso & " AND ternro=" & Ternro
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            'hasta aca
        'hasta aca
    End If
    
    Call buscarPosProceso(proc, pliq)
    
        If esConceptoAcu1 Then
            StrSql = "SELECT * FROM cabliq"
            StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
            StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
            StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
            StrSql = StrSql & " WHERE concepto.conccod = " & acu1
            StrSql = StrSql & " AND cabliq.pronro =" & proc & " AND empleado=" & Ternro
            OpenRecordset StrSql, rsconsult2
            If Not rsconsult2.EOF Then
                If Not EsNulo(rsconsult2!dlimonto) Then
                    valorAcu1 = rsconsult2!dlimonto
                    Flog.writeline "Se encontraron datos del concepto:" & acu1
                    
                End If
            Else
                Flog.writeline "No hay datos del concepto."
            End If
        Else
            StrSql = "SELECT * FROM cabliq"
            StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro "
            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
            StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
            StrSql = StrSql & " WHERE acu_liq.acuNro = " & acu1
            StrSql = StrSql & " AND cabliq.pronro =" & proc & " AND empleado=" & Ternro
            OpenRecordset StrSql, rsconsult2
            If Not rsconsult2.EOF Then
                If Not EsNulo(rsconsult2!almonto) Then
                    valorAcu1 = rsconsult2!almonto
                    Flog.writeline "Se encontraron datos del acumulador:" & acu1
                    Seguir = True
                End If
            Else
                Flog.writeline "No hay datos del acumulador."
            End If
        End If
        
        'If (rsconsult2!pliqmes = mesAnt) And (rsconsult2!pliqanio = anioAnt) Then
        'Else
        '    mesAnt = rsconsult2!pliqmes
        '    anioAnt = rsconsult2!pliqanio
        'End If
        
        'hacer update de la columa o arreglo de 12 filas x 5 columnas
        If Seguir Then
            If (mesAnt <> rsconsult2!pliqmes) Or (anioAnt <> rsconsult2!pliqanio) Then
                primeravez = True
            Else
                primeravez = False
            End If
            If (primeravez) Then
                StrSql = "INSERT INTO rep_aguinaldo_det"
                StrSql = StrSql & "(bpronro, ternro, pliqmes, pliqanio, pliqdesc, prodesc, fechaPago" & pos
                StrSql = StrSql & ", valor" & pos & ")"
                StrSql = StrSql & " Values "
                StrSql = StrSql & "("
                StrSql = StrSql & nroproceso
                StrSql = StrSql & ", " & Ternro
                StrSql = StrSql & ", " & rsconsult2!pliqmes
                StrSql = StrSql & ", " & rsconsult2!pliqanio
                StrSql = StrSql & ", '" & rsconsult2!pliqdesc & "'"
                StrSql = StrSql & ", '" & rsconsult2!prodesc & "'"
                StrSql = StrSql & ", '" & profecpago & "'"
                StrSql = StrSql & ", " & valorAcu1
                StrSql = StrSql & ")"
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                Flog.writeline "se inserto el primer registro"
                'priveravez = False
                mesAnt = rsconsult2!pliqmes
                anioAnt = rsconsult2!pliqanio
            Else
                'UPDATE
                StrSql = "UPDATE rep_aguinaldo_det "
                StrSql = StrSql & " SET valor" & pos & " =" & valorAcu1
                StrSql = StrSql & " ,fechaPago" & pos & " ='" & profecpago & "'"
                StrSql = StrSql & " WHERE bpronro=" & nroproceso
                StrSql = StrSql & " AND pliqmes=" & rsconsult2!pliqmes
                StrSql = StrSql & " AND pliqanio=" & rsconsult2!pliqanio
                StrSql = StrSql & " AND ternro=" & rsconsult2!Empleado
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                Flog.writeline "se actualizo el registro"
                mesAnt = rsconsult2!pliqmes
                anioAnt = rsconsult2!pliqanio
            End If
            'hasta aca
            rsconsult2.Close
            Seguir = False
        Else
        End If
    
    
        '=============================================
        'ARMO LA LISTA CON LOS PROCESOS DEL EMPLEADO
        '=============================================
        listaNueva = listaNueva & ", " & proc
        '=============================================
    
    'actualizo algunos datos de la tabla cabecera
    'StrSql = " UPDATE rep_aguinaldo "
    'StrSql = StrSql & " SET fechaDesde=" & procFechaDesde & ", fechahasta = " & procFechaHasta
    'StrSql = StrSql & " WHERE bpronro=" & nroproceso & " AND ternro=" & Ternro
    'objconnProgreso.Execute StrSql, , adExecuteNoRecords
    'hasta aca
Next
'hasta aca

    '================================================================
    'BUSCO EL VALOR DE LOS CONCEPTOS IMPRIMIBLE EN EL PROCESO
    StrSql = "SELECT sum(dlimonto) monto, concepto.concabr, concepto.conccod FROM cabliq "
    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
    StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
    StrSql = StrSql & " WHERE cabliq.pronro IN(" & listapronro & ") AND empleado=" & Ternro
    StrSql = StrSql & " GROUP BY dlimonto,concepto.conccod, concepto.concabr"
    StrSql = StrSql & " ORDER BY concepto.conccod"
    
    'StrSql = "SELECT distinct concepto.conccod, concepto.concabr, SUM(dlimonto) monto FROM cabliq "
    'StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    'StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
    'StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro"
    'StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.concimp = -1 "
    'StrSql = StrSql & " WHERE cabliq.pronro IN (" & listaNueva & ") AND empleado=" & Ternro
    'StrSql = StrSql & " GROUP BY concepto.conccod, concepto.concabr"
    OpenRecordset StrSql, rsconsult2
    Do While Not rsconsult2.EOF
        'hacer insert en la tabla
        StrSql = "INSERT INTO rep_aguinaldo_det "
        StrSql = StrSql & " (bpronro, ternro, pronro, prodesc, valor1) "
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & " (" & nroproceso
        StrSql = StrSql & ", " & Ternro
        StrSql = StrSql & ", 1 "
        StrSql = StrSql & ", '" & rsconsult2!concabr & "'"
        StrSql = StrSql & ", " & Replace(rsconsult2!Monto, ",", ".")
        StrSql = StrSql & ")"
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Se inserto el concepto imprimible:" & rsconsult2!concabr
    rsconsult2.MoveNext
    Loop
    '================================================================

End Sub
Sub buscarPosProceso(ByVal proc, ByVal pliq)

Dim rsconsult As New ADODB.Recordset
Dim Encontro As Boolean


Encontro = False
pos = 0
StrSql = " SELECT * FROM proceso "
StrSql = StrSql & " WHERE pliqnro=" & pliq & " AND tprocnro IN(1,15)"
StrSql = StrSql & " ORDER BY profecini ASC"
OpenRecordset StrSql, rsconsult
Do While Not rsconsult.EOF And Not (Encontro)
    pos = pos + 1
    If (rsconsult!Pronro = proc) Then
        Encontro = True
    End If
rsconsult.MoveNext
Loop
rsconsult.Close
End Sub

Sub filtrarProcFases(ByVal altfec As String, ByVal bajfec As String, ByRef arregloProcesos)
'busca en la lista de procesos solo los procesos que pertenecen a la fase del empleado
Dim rsconsultAux As New ADODB.Recordset
Dim salida As String
Dim proc As Long
Dim arreglo
    
    salida = ""
    
    For i = 1 To UBound(arregloProcesos)
        proc = Split(arregloProcesos(i), "@")(0)
        StrSql = " SELECT pronro FROM proceso WHERE pronro = " & proc
        If Not EsNulo(bajfec) Then
            StrSql = StrSql & " AND ((profecini <= " & ConvFecha(altfec) & " AND (profecfin >= " & ConvFecha(bajfec) & " or profecfin >= " & ConvFecha(altfec) & "))"
            StrSql = StrSql & " OR (profecini >= " & ConvFecha(altfec) & " AND (profecini <= " & ConvFecha(bajfec) & ")))"
        Else
            StrSql = StrSql & " AND (profecini >= " & ConvFecha(altfec) & " OR profecfin >= " & ConvFecha(altfec) & ")"
        End If
        OpenRecordset StrSql, rsconsultAux
        If Not rsconsultAux.EOF Then
            salida = salida & "!!" & arregloProcesos(i)
        End If
        rsconsultAux.Close
    Next
    arreglo = Split(salida, "!!")
    If UBound(arreglo) <> -1 Then
        ReDim Preserve arregloProcesos(UBound(arreglo))
        For i = 0 To UBound(arreglo)
            arregloProcesos(i) = arreglo(i)
        Next
    Else
        ReDim Preserve arregloProcesos(0)
    End If
    
End Sub
