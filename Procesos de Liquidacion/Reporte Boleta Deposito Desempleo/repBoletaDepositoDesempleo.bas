Attribute VB_Name = "repBoletoDepositoDesempleo"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "21/01/2008"
'Global Const UltimaModificacion = "Inicial"

'Global Const Version = "1.01"
'Global Const FechaModificacion = "25/01/2008"
'Global Const UltimaModificacion = "Permite varias estructuras de banco."

'Global Const Version = "1.02" ' Cesar Stankunas
'Global Const FechaModificacion = "07/08/2009"
'Global Const UltimaModificacion = ""    'Encriptacion de string connection

'Global Const Version = "1.03"
'Global Const FechaModificacion = "01/07/2015"
'Global Const UltimaModificacion = "se agregaron controles sobre la configuración-Miriam Ruiz-CAS-30812 - H&A - Errores R4 V2"

Global Const Version = "1.04"
Global Const FechaModificacion = "06/07/2015"
Global Const UltimaModificacion = "se agregaron controles cuando no había datos-Miriam Ruiz-CAS-30812 - H&A - Errores R4 V2"

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

Global TEnro1 As Integer
Global Estrnro1 As Integer
Global TEnro2 As Integer
Global Estrnro2 As Integer
Global TEnro3 As Integer
Global Estrnro3 As Integer
Global fecEstr As String
Global Formato As Integer

Dim repboldesnro As Long
Dim bpronro As Long
Dim pliqnro As Long
Dim tfiltro As String
Dim Ternro As Long
Dim empleg As Long
Dim Bannro As Long
Dim bancdesc As String
Dim bancta As String
Dim terape As String
Dim terape2 As String
Dim ternom As String
Dim ternom2 As String
Dim calle As String
Dim calleNro As String
Dim Localidad As String
Dim Provincia As String
Dim pliqmes  As String
Dim pliqanio As String
Dim Aporte  As String
Dim AporteDesc  As String
Dim Empnro As Long
Dim empnom  As String
Dim empcalle  As String
Dim empcalleNro  As String
Dim empLocalidad  As String
Dim RNIC As String
Dim IERIC  As String
Dim repfec As String

Dim tipoprocesos As Integer
Dim listaprocesos
Dim proaprob As Integer
Dim proaprobdes As String

Dim COAC_FONDO As String
Dim COAC_TIPO As String
Dim EST_BCO_FONDO As String
Dim TD_IERIC As Long
Dim TD_RNIC As Long


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
    Dim rsPeriodos As New ADODB.Recordset
    'Dim pliqnro
    'Dim pliqdesde
    'Dim pliqhasta
    Dim tipoDepuracion
    Dim historico As Boolean
    Dim param
    Dim listapronro
    Dim proNro
    Dim Ternro
    Dim arrpronro
    Dim Periodos
    Dim rsEmpl As New ADODB.Recordset
    Dim i
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
    
    Nombre_Arch = PathFLog & "RepBoletaDepositoDesempleo" & "-" & NroProceso & ".log"
    
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
    
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    On Error GoTo CE
    
    HuboErrores = False
    
    'Obtengo la cantidad de empledos a procesar
    StrSql = "SELECT count(*) AS total FROM batch_empleado WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    cantRegistros = CInt(objRs!total)
    totalEmpleados = cantRegistros
    
    objRs.Close
    
    Flog.writeline "Inicio Proceso de Boleta Deposito Fondo Desempleo : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
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
        
        
        'Obtengo el periodo de liquidacion
        Flog.writeline "Obtengo el Período de Liquidacion"
        pliqnro = CLng(ArrParametros(0))
        Flog.writeline "Período = " & pliqnro
        
        'Obtengo el tipo de proceso (-1 = todos)
        Flog.writeline "Obtengo el tipo de proceso "
        tipoprocesos = CLng(ArrParametros(1))
        Flog.writeline "tipo de procesos = " & tipoprocesos
        
        'Obtengo la lista de procesos
        Flog.writeline "Obtengo la Lista de Procesos"
        listapronro = ArrParametros(2)
        Flog.writeline "Lista de Procesos = " & listapronro
        
        'Obtengo el estado de los procesos
        Flog.writeline "estado de los procesos"
        proaprob = CLng(ArrParametros(3))
        Flog.writeline "estado de los procesos = " & proaprob
       
        'Obtengo la empresa
        Flog.writeline "Empresa"
        Empnro = CLng(ArrParametros(4))
        Flog.writeline "empresa(nro) = " & Empnro
               
        'Obtengo los cortes de estructura
        Flog.writeline "Obtengo los cortes de estructuras"
        Flog.writeline "Obtengo estructura 1"
        TEnro1 = CInt(ArrParametros(5))
        Estrnro1 = CInt(ArrParametros(6))
        Flog.writeline "Corte 1 = " & TEnro1 & " - " & Estrnro1
        
        Flog.writeline "Obtengo estructura 2"
        TEnro2 = CInt(ArrParametros(7))
        Estrnro2 = CInt(ArrParametros(8))
        Flog.writeline "Corte 2 = " & TEnro2 & " - " & Estrnro2
        
        Flog.writeline "Obtengo estructura 3"
        TEnro3 = CInt(ArrParametros(9))
        Estrnro3 = CInt(ArrParametros(10))
        Flog.writeline "Corte 3 = " & TEnro3 & " - " & Estrnro3
        
        'Flog.writeline "Obtengo la Fecha"
        'fecCorte = ArrParametros(11)
        'Flog.writeline "Fecha = " & fecEstr
        
        'EMPIEZA EL PROCESO
        Select Case CInt(proaprob)
            Case 0: proaprobdes = "Sin Liquidar"
            Case 1: proaprobdes = "Liquidado"
            Case 2: proaprobdes = "Aprob. Prov."
            Case 3: proaprobdes = "Aprob. Def."
        End Select
  
 'Busco el periodo desde
'       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqdesde
'       OpenRecordset StrSql, objRs
'
'       If Not objRs.EOF Then
'          FechaDesde = objRs!pliqdesde
'          descDesde = objRs!pliqDesc
'       Else
'          Flog.writeline "No se encontro el periodo desde."
'          Exit Sub
'       End If
'
'       objRs.Close
'
'       'Busco el periodo hasta
'       StrSql = "SELECT * FROM periodo WHERE pliqnro = " & pliqhasta
'       OpenRecordset StrSql, objRs
'
'       If Not objRs.EOF Then
'          FechaHasta = objRs!pliqhasta
'          descHasta = objRs!pliqDesc
'       Else
'          Flog.writeline "No se encontro el periodo hasta."
'          Exit Sub
'       End If
'
'       objRs.Close
        'Cargo la configuracion del reporte
        Flog.writeline "Cargo la Configuración del Reporte"
       EST_BCO_FONDO = "0"
        StrSql = " SELECT * FROM confrep WHERE repnro = 226 "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Do While Not objRs.EOF
                Select Case objRs!confnrocol
                    Case 1: 'Conceto o acumulador
                       If objRs!confval2 = "" Then
                            COAC_FONDO = "0"
                            Flog.writeline "El valor 2 de la columna 1 no está configurado"
                       Else
                            COAC_FONDO = CStr(objRs!confval2)
                       End If
                        COAC_TIPO = CStr(objRs!conftipo)
                        If objRs!conftipo = "CO" Then
                            StrSql = " SELECT * FROM concepto WHERE conccod = '" & COAC_FONDO & "'"
                            OpenRecordset StrSql, objRs2
                            If objRs2.EOF Then
                                Flog.writeline "El concepto no existe."
                                Exit Sub
                            End If
                            objRs2.Close
                        ElseIf objRs!conftipo = "AC" Then
                            StrSql = " SELECT * FROM acumulador WHERE acunro = '" & COAC_FONDO & "'"
                            OpenRecordset StrSql, objRs2
                            If objRs2.EOF Then
                                Flog.writeline "El acumulador no existe."
                                Exit Sub
                            End If
                            objRs2.Close
                        End If
                    Case 2: 'Estructura del banco del fondo de desempleo
                        If objRs!confval2 = "" Then
                            EST_BCO_FONDO = "0"
                            Flog.writeline "El valor 2 de la columna 2 no está configurado"
                        Else
                             EST_BCO_FONDO = EST_BCO_FONDO & "," & CStr(objRs!confval2)
                        End If
                        'StrSql = " SELECT * FROM estructura WHERE estrnro = '" & EST_BCO_FONDO & "'"
                        'OpenRecordset StrSql, objRs2
                        'If objRs2.EOF Then
                        '    Flog.writeline "La estructura Banco en la columna 2 no existe."
                        '    Exit Sub
                        'End If
                        'objRs2.Close
                    Case 3: 'Tipo doc IERIC
                        If objRs!confval2 = "" Then
                            TD_IERIC = 0
                            Flog.writeline "El valor 2 de la columna 3 no está configurado"
                        Else
                            TD_IERIC = CLng(objRs!confval2)
                        End If
                        StrSql = " SELECT * FROM tipodocu WHERE tidnro = " & TD_IERIC
                        OpenRecordset StrSql, objRs2
                        If objRs2.EOF Then
                            Flog.writeline "El tipo de documento IERIC no existe."
                            Exit Sub
                        End If
                        objRs2.Close
                    Case 4: 'Tipo doc RNIC
                        If objRs!confval2 = "" Then
                            TD_RNIC = 0
                            Flog.writeline "El valor 2 de la columna 4 no está configurado"
                        Else
                            TD_RNIC = CLng(objRs!confval2)
                        End If
                        StrSql = " SELECT * FROM tipodocu WHERE tidnro = " & TD_RNIC
                        OpenRecordset StrSql, objRs2
                        If objRs2.EOF Then
                            Flog.writeline "El tipo de documento RNIC no existe."
                            Exit Sub
                        End If
                        objRs2.Close
                    Case Else
                End Select
                objRs.MoveNext
            Loop
            'Limpio la variable para las consultas sql
           ' EST_BCO_FONDO = Right(EST_BCO_FONDO, Len(EST_BCO_FONDO) - 1)
        Else
            Flog.writeline "No se encontró la configuración del reporte."
            Exit Sub
        End If
       
       
        'Cargo la configuracion del reporte
        'Call CargarConfiguracionReporte(Modelo)
        
        'Obtengo los empleados sobre los que tengo que generar los recibos
        
        Flog.writeline "Cargo los Datos generales del reporte"
        HuboError = False
        HuboErrores = False
        Call cargarDatosUnicos
        If HuboError = True Then
            Exit Sub
        End If
        
        'Call CargarEmpleados(NroProceso, rsEmpl)
        Flog.writeline "Cargo los Empleados "
        Call CargarEmpleados(pliqnro, listapronro, rsEmpl)
       cantRegistros = rsEmpl.RecordCount
       totalEmpleados = cantRegistros
        'Guardo en la BD el encabezado
        'Flog.writeline "Genero el encabezado del Reporte"
        'EMPRESA - PERIODO - PRODESO ESTADO - FECHA
        'Call GenerarEncabezadoReporte
       
        StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
       
        If Not rsEmpl.EOF Then
            '-------------------------------------------------------------------
            'Genero por cada empleado un registro
           
            Do Until rsEmpl.EOF
                'arrpronro = Split(listapronro, ",")
                EmpErrores = False
                'ternro = rsEmpl!ternro
                'orden = rsEmpl!estado
              
                'Genero una entrada para el empleado por cada proceso
                'For I = 0 To UBound(arrpronro)
                'proNro = arrpronro(I)
                'Flog.writeline "Generando datos empleado " & ternro & " para el proceso " & proNro
                 
                 Call GenerarDatosEmpleado(rsEmpl, listapronro)
                 
                'Next
              
                'Actualizo el estado del proceso
                TiempoAcumulado = GetTickCount
                cantRegistros = cantRegistros - 1
                
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((totalEmpleados - cantRegistros) * 100) / totalEmpleados) & _
                       ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                       ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
                 
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
              
              
                rsEmpl.MoveNext
            Loop
        Else
            
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



Sub GenerarDatosEmpleado(rsEmpl, listapronro)
    '--------------------------------------------------------------------
    ' Se encarga de generar los datos para el empleado por cada proceso
    '--------------------------------------------------------------------

    Dim StrSql As String
    Dim rsConsult As New ADODB.Recordset
    
    'Variables donde se guardan los datos del INSERT final
    
    'Inicializo x empleados
    Ternro = 0
    empleg = 0
    bancta = ""
    terape = ""
    terape2 = ""
    ternom = ""
    ternom2 = ""
    calle = ""
    calleNro = ""
    Localidad = ""
    Provincia = ""
    'IERIC = ""
    'pliqmes = ""
    'pliqanio = ""
    Aporte = 0
    AporteDesc = ""
    
    '------------------------------------------------------------------
    '   Busco los datos del empleado
    '------------------------------------------------------------------
    StrSql = " SELECT empleg, terape, terape2, ternom, ternom2 "
    StrSql = StrSql & " , calle, nro, piso, oficdepto, locdesc, provdesc "
    StrSql = StrSql & " FROM empleado "
    StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = empleado.ternro "
    StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro "
    StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
    StrSql = StrSql & " INNER JOIN provincia ON provincia.provnro = detdom.provnro "
    StrSql = StrSql & " WHERE empleado.ternro = " & rsEmpl!Ternro
    StrSql = StrSql & " AND domdefault = -1 "
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        terape = nullToString(rsConsult!terape)
        terape2 = nullToString(rsConsult!terape2)
        ternom = nullToString(rsConsult!ternom)
        ternom2 = nullToString(rsConsult!ternom2)
        calle = nullToString(rsConsult!calle)
        calleNro = nullToString(rsConsult!nro)
        Localidad = nullToString(rsConsult!locdesc)
        Provincia = nullToString(rsConsult!provDesc)
        rsConsult.Close
    Else
        Flog.writeline "No se encontro el empleado (" & CStr(rsEmpl!empleg) & ")"
        rsConsult.Close
        Exit Sub
    End If

'    ------------------------------------------------------------------
'       Busco IERIC del empleado
'    ------------------------------------------------------------------
'    StrSql = " SELECT nrodoc "
'    StrSql = StrSql & " FROM empleado "
'    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro "
'    StrSql = StrSql & " WHERE empleado.ternro = " & rsEmpl!ternro
'    StrSql = StrSql & " AND ter_doc.tidnro = " & TD_IERIC
'    OpenRecordset StrSql, rsConsult
'    If Not rsConsult.EOF Then
'        IERIC = CStr(rsConsult!nrodoc)
'        rsConsult.Close
'    Else
'        Flog.writeline "No se encontro el IERIC para el empleado (" & CStr(rsEmpl!empleg) & ")"
'        rsConsult.Close
'        Exit Sub
'    End If

    '------------------------------------------------------------------
    '   Busco los datos del banco de la cuenta bancaria del empleado
    '------------------------------------------------------------------
    ''Busco los datos del Banco (EST_BCO_FONDO)
    'StrSql = " SELECT bandesc "
    'StrSql = StrSql & " FROM banco "
    'StrSql = StrSql & " WHERE estrnro = " & EST_BCO_FONDO
    'OpenRecordset StrSql, rsConsult
    'If Not rsConsult.EOF Then
    '    bancdesc = CStr(rsConsult!bandesc)
    'Else
    '    Flog.writeline "No se encontro el BANCO configurado como Fondo de Desempleo xx."
    '    HuboError = True
    '    Exit Sub
    'End If
    'rsConsult.Close

    
    'Busco los datos de la cuenta bancaria del empleado
    StrSql = " SELECT ctabnro, bandesc "
    StrSql = StrSql & " FROM ctabancaria "
    StrSql = StrSql & " INNER JOIN banco ON ctabancaria.banco = banco.ternro "
    StrSql = StrSql & " WHERE ctabancaria.ternro = " & rsEmpl!Ternro
    StrSql = StrSql & " AND banco.estrnro in (" & EST_BCO_FONDO & ")"
    StrSql = StrSql & " AND ctabancaria.ctabestado = 0 "
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        bancta = CStr(rsConsult!ctabnro)
        bancdesc = CStr(rsConsult!bandesc)
        rsConsult.Close
    Else
        rsConsult.Close
        StrSql = " SELECT ctabnro, bandesc "
        StrSql = StrSql & " FROM ctabancaria "
        StrSql = StrSql & " INNER JOIN banco ON ctabancaria.banco = banco.ternro "
        StrSql = StrSql & " WHERE ctabancaria.ternro = " & rsEmpl!Ternro
        StrSql = StrSql & " AND banco.estrnro in (" & EST_BCO_FONDO & ")"
        StrSql = StrSql & " AND ctabancaria.ctabestado = -1 "
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            bancta = CStr(rsConsult!ctabnro)
            bancdesc = CStr(rsConsult!bandesc)
            rsConsult.Close
        Else
            Flog.writeline "No se encontro la Cuenta Bancaria para el empleado (" & CStr(rsEmpl!empleg) & ")"
            rsConsult.Close
            Exit Sub
        End If
    End If

    '------------------------------------------------------------------
    '   Busco los datos de los aportes del empleado
    '   Busco los valores de los conceptos y acumuladores
    '------------------------------------------------------------------
    If COAC_TIPO = "CO" Then 'Concepto
        StrSql = " SELECT SUM(detliq.dlimonto) Monto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
        StrSql = StrSql & " INNER JOIN concepto ON  concepto.concnro = detliq.concnro "
        StrSql = StrSql & " WHERE cabliq.empleado = " & rsEmpl!Ternro
        StrSql = StrSql & " AND proceso.pliqnro = " & pliqnro
        If tipoprocesos <> -1 Then
            StrSql = StrSql & " AND proceso.pronro in (" & CStr(Replace(listapronro, "-", ",")) & ")"
        End If
        StrSql = StrSql & " AND concepto.conccod = '" & COAC_FONDO & "'"
    Else 'Acumulador
        StrSql = " SELECT SUM(acu_liq.almonto) Monto "
        StrSql = StrSql & " FROM cabliq "
        StrSql = StrSql & " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
        StrSql = StrSql & " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro"
        StrSql = StrSql & " And cabliq.Empleado = " & rsEmpl!Ternro
        StrSql = StrSql & " AND proceso.pliqnro = " & pliqnro
        If tipoprocesos <> -1 Then
            StrSql = StrSql & " AND proceso.pronro in (" & CStr(Replace(listapronro, "-", ",")) & ")"
        End If
        StrSql = StrSql & " AND acu_liq.acunro = " & COAC_FONDO
    End If
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        Aporte = nullTobouble(rsConsult!Monto)
        If Aporte <> 0 Then
            AporteDesc = EnLetras(Aporte)
        Else
            AporteDesc = "Cero"
        End If
        rsConsult.Close
    Else
        Flog.writeline "No se encontro el Concepto o Acumulador para el empleado (" & CStr(rsEmpl!empleg) & ")"
        rsConsult.Close
        Exit Sub
    End If
    


    '-------------------------------------------------------------------------------
    'Inserto los datos en la BD
    '-------------------------------------------------------------------------------
    Flog.writeline "Inserto los datos en la BD - Legajo = " & rsEmpl!empleg & " TERNRO " & rsEmpl!Ternro & " proceso " & NroProceso & " periodo " & pliqnro
    'repboldesnro = 0

    tfiltro = empnom & " - " & pliqmes & " " & pliqanio & " - " & proaprobdes & " - " & repfec

    StrSql = " INSERT INTO rep_boletadesempleo ( "
    StrSql = StrSql & " bpronro ,pliqnro ,tfiltro ,ternro ,empleg ,Bannro ,bancdesc ,bancta "
    StrSql = StrSql & " ,terape ,terape2 ,ternom ,ternom2 ,calle ,calleNro ,Localidad ,Provincia "
    StrSql = StrSql & " ,IERIC ,pliqmes, pliqanio,Aporte ,AporteDesc ,Empnro ,empnom ,empcalle "
    StrSql = StrSql & " ,empcalleNro ,empLocalidad ,RNIC ,repfec "
    StrSql = StrSql & " ) values ( "
    StrSql = StrSql & NroProceso & ","
    StrSql = StrSql & pliqnro & ","
    StrSql = StrSql & "'" & tfiltro & "',"
    StrSql = StrSql & CStr(rsEmpl!Ternro) & ","
    StrSql = StrSql & CStr(rsEmpl!empleg) & ","
    StrSql = StrSql & Bannro & ","
    StrSql = StrSql & "'" & bancdesc & "',"
    StrSql = StrSql & "'" & bancta & "',"
    StrSql = StrSql & "'" & terape & "',"
    StrSql = StrSql & "'" & terape2 & "',"
    StrSql = StrSql & "'" & ternom & "',"
    StrSql = StrSql & "'" & ternom2 & "',"
    StrSql = StrSql & "'" & calle & "',"
    StrSql = StrSql & "'" & calleNro & "',"
    StrSql = StrSql & "'" & Localidad & "',"
    StrSql = StrSql & "'" & Provincia & "',"
    StrSql = StrSql & "'" & IERIC & "',"
    StrSql = StrSql & "'" & pliqmes & "',"
    StrSql = StrSql & "'" & pliqanio & "',"
    StrSql = StrSql & "" & Aporte & ","
    StrSql = StrSql & "'" & AporteDesc & "',"
    StrSql = StrSql & Empnro & ","
    StrSql = StrSql & "'" & empnom & "',"
    StrSql = StrSql & "'" & empcalle & "',"
    StrSql = StrSql & "'" & empcalleNro & "',"
    StrSql = StrSql & "'" & empLocalidad & "',"
    StrSql = StrSql & "'" & RNIC & "',"
    StrSql = StrSql & "'" & repfec & "'"
    StrSql = StrSql & ")"
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
Sub CargarEmpleados(pliqnro, listaprocesos, ByRef rsEmpl As ADODB.Recordset)

    Dim StrEmpl As String
    Dim stremplsel As String
    Dim Fecha_Fin_Periodo As String

    StrEmpl = "SELECT pliqhasta "
    StrEmpl = StrEmpl & " FROM periodo "
    StrEmpl = StrEmpl & " WHERE pliqnro = " & pliqnro
    OpenRecordset StrEmpl, rsEmpl
    If Not rsEmpl.EOF Then
        Fecha_Fin_Periodo = rsEmpl!pliqhasta
    End If
    rsEmpl.Close


    'Busco los Empleados de los procesos a evaluar
    StrEmpl = "SELECT DISTINCT empleado.ternro, empleg "
    'stremplsel = "SELECT DISTINCT empleado.ternro, empleg "
    StrEmpl = StrEmpl & " FROM empleado "
    StrEmpl = StrEmpl & " INNER JOIN tercero ON empleado.ternro = tercero.ternro "
    StrEmpl = StrEmpl & " INNER JOIN his_estructura his_Empresa ON his_empresa.ternro = empleado.ternro "
    StrEmpl = StrEmpl & " INNER JOIN cabliq ON empleado.ternro = cabliq.empleado "
    StrEmpl = StrEmpl & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
    StrEmpl = StrEmpl & " INNER JOIN empresa ON his_empresa.estrnro = empresa.estrnro "
    
    If TEnro1 <> 0 Then
        'stremplsel = stremplsel & ", TE1.tedabr te1dabr "
        StrEmpl = StrEmpl & " INNER JOIN his_estructura TE1 ON te1.ternro = empleado.ternro "
    End If
    If TEnro2 <> 0 Then
        'stremplsel = stremplsel & ", TE2.tedabr te2dabr "
        StrEmpl = StrEmpl & " INNER JOIN his_estructura TE2 ON te2.ternro = empleado.ternro "
    End If
    If TEnro3 <> 0 Then
        'stremplsel = stremplsel & ", TE3.tedabr te3dabr "
        StrEmpl = StrEmpl & " INNER JOIN his_estructura TE3 ON te3.ternro = empleado.ternro "
    End If
    
    StrEmpl = StrEmpl & " WHERE his_empresa.tenro = 10 AND empresa.empnro = " & Empnro '& " AND " & Filtro
    StrEmpl = StrEmpl & " AND pliqnro = " & pliqnro
    'StrEmpl = StrEmpl & " WHERE pliqnro = " & pliqnro
    'StrEmpl = StrEmpl & " AND pliqnro = " & pliqnro
    
    If tipoprocesos <> -1 Then
        StrEmpl = StrEmpl & " AND proceso.pronro in (" & CStr(Replace(listaprocesos, "-", ",")) & ")"
    End If
    
    'If tipoprocesos = 0 Then
    '    StrEmpl = StrEmpl & " AND proaprob " = proaprob
    'End If
    
    If TEnro1 <> 0 Then
        StrEmpl = StrEmpl & " AND te1.tenro = " & TEnro1 & " AND "
        If Estrnro1 <> 0 Then
            StrEmpl = StrEmpl & " te1.estrnro = " & Estrnro1 & " AND "
        End If
        StrEmpl = StrEmpl & " (te1.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
        StrEmpl = StrEmpl & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te1.htethasta) or (te1.htethasta is null)) "
    End If
    If TEnro2 <> 0 Then
        StrEmpl = StrEmpl & " AND te2.tenro = " & TEnro2 & " AND "
        If Estrnro2 <> 0 Then
            StrEmpl = StrEmpl & " te2.estrnro = " & Estrnro2 & " AND "
        End If
        StrEmpl = StrEmpl & " (te2.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
        StrEmpl = StrEmpl & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te2.htethasta) or (te2.htethasta is null))  "
    End If
    If TEnro3 <> 0 Then
        StrEmpl = StrEmpl & " AND te3.tenro = " & TEnro3 & " AND "
        If Estrnro3 <> 0 Then
            StrEmpl = StrEmpl & " te3.estrnro = " & Estrnro3 & " AND "
        End If
        StrEmpl = StrEmpl & " (te3.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
        StrEmpl = StrEmpl & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= te3.htethasta) or (te3.htethasta is null))"
    End If
    StrEmpl = StrEmpl & " ORDER BY empleg " '& orden
    'OpenRecordset StrEmpl, rs_Empleados
    OpenRecordset StrEmpl, rsEmpl
    
End Sub



Sub cargarDatosUnicos()
    Dim StrSql As String
    Dim rsConsult As New ADODB.Recordset
        
    'Busco los datos de la Empresa
    StrSql = " SELECT empnom, calle, nro, piso, oficdepto, locdesc"
    StrSql = StrSql & " FROM empresa "
    StrSql = StrSql & " INNER JOIN cabdom ON cabdom.ternro = empresa.ternro "
    StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro "
    StrSql = StrSql & " INNER JOIN localidad ON localidad.locnro = detdom.locnro "
    StrSql = StrSql & " WHERE empnro = " & Empnro
    StrSql = StrSql & " AND domdefault = -1 "
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        empnom = CStr(rsConsult!empnom)
        empcalle = CStr(rsConsult!calle)
        empcalleNro = CStr(rsConsult!nro)
        empLocalidad = CStr(rsConsult!locdesc)
    Else
        Flog.writeline "No se encontro la Empresa."
        HuboErrores = True
        Exit Sub
    End If
    rsConsult.Close
    
    'Busco el IERIC de la empresa
    StrSql = " SELECT nrodoc "
    StrSql = StrSql & " FROM empresa "
    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empresa.ternro "
    StrSql = StrSql & " WHERE empnro = " & Empnro
    StrSql = StrSql & " AND tidnro = " & TD_IERIC
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        IERIC = CStr(rsConsult!NroDoc)
    Else
        Flog.writeline "No se encontro el IERIC de la empresa."
        HuboError = True
        HuboErrores = True
        Exit Sub
    End If
    rsConsult.Close
    
    
    'Busco el RNIC de la empresa
    StrSql = " SELECT nrodoc "
    StrSql = StrSql & " FROM empresa "
    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empresa.ternro "
    StrSql = StrSql & " WHERE empnro = " & Empnro
    StrSql = StrSql & " AND tidnro = " & TD_RNIC
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        RNIC = CStr(rsConsult!NroDoc)
    Else
        Flog.writeline "No se encontro el RNIC de la empresa."
        HuboError = True
        HuboErrores = True
        Exit Sub
    End If
    rsConsult.Close
    
    
    'Convierto la fecha en cadena
    repfec = Day(Now) & " de "
    Select Case Month(Now)
        Case 1:
            repfec = repfec & "Enero"
        Case 2:
            repfec = repfec & "Febrero"
        Case 3:
            repfec = repfec & "Marzo"
        Case 4:
            repfec = repfec & "Abril"
        Case 5:
            repfec = repfec & "Mayo"
        Case 6:
            repfec = repfec & "Junio"
        Case 7:
            repfec = repfec & "Julio"
        Case 8:
            repfec = repfec & "Agosto"
        Case 9:
            repfec = repfec & "Septiembre"
        Case 10:
            repfec = repfec & "Octubre"
        Case 11:
            repfec = repfec & "Noviembre"
        Case 12:
            repfec = repfec & "Diciembre"
    End Select
    repfec = repfec & " de " & Year(Now)
    
    
'    'Busco los datos del Banco (EST_BCO_FONDO)
'    StrSql = " SELECT bandesc "
'    StrSql = StrSql & " FROM banco "
'    StrSql = StrSql & " WHERE estrnro = " & EST_BCO_FONDO
'    OpenRecordset StrSql, rsConsult
'    If Not rsConsult.EOF Then
'        bancdesc = CStr(rsConsult!bandesc)
'    Else
'        Flog.writeline "No se encontro el BANCO configurado como Fondo de Desempleo xx."
'        HuboError = True
'        Exit Sub
'    End If
'    rsConsult.Close
    
    'Busco el mes del periodo de liquidacion
    StrSql = " SELECT pliqmes, pliqanio "
    StrSql = StrSql & " FROM periodo "
    StrSql = StrSql & " WHERE pliqnro = " & pliqnro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        Select Case rsConsult!pliqmes
            Case 1:
                pliqmes = "Enero"
            Case 2:
                pliqmes = "Febrero"
            Case 3:
                pliqmes = "Marzo"
            Case 4:
                pliqmes = "Abril"
            Case 5:
                pliqmes = "Mayo"
            Case 6:
                pliqmes = "Junio"
            Case 7:
                pliqmes = "Julio"
            Case 8:
                pliqmes = "Agosto"
            Case 9:
                pliqmes = "Septiembre"
            Case 10:
                pliqmes = "Octubre"
            Case 11:
                pliqmes = "Noviembre"
            Case 12:
                pliqmes = "Diciembre"
        End Select
        pliqanio = rsConsult!pliqanio
    Else
        Flog.writeline "No se encontro el Periodo de liquidacion."
        HuboError = True
        HuboErrores = True
        Exit Sub
    End If
    rsConsult.Close
    
    
End Sub
Function nullToString(cadena As Variant) As String
    If IsNull(cadena) Then
        nullToString = ""
    Else
        nullToString = CStr(cadena)
    End If
End Function

Function nullTobouble(cadena As Variant) As Double
    If IsNull(cadena) Then
        nullTobouble = 0
    Else
        nullTobouble = CDbl(cadena)
    End If
End Function

