Attribute VB_Name = "repDDJJSindicatos"
Option Explicit

'Global Const Version = "1.00"
'Global Const FechaModificacion = "05/09/2013"
'Global Const UltimaModificacion = " " 'LED - Inicial - CAS-20033 - HDI - CUSTOM REPORTE DDJJ SINDICATO DEL SEGURO

'Global Const Version = "1.01"
'Global Const FechaModificacion = "30/10/2013"
'Global Const UltimaModificacion = " " 'LED - CAS-20033 - HDI - CUSTOM REPORTE DDJJ SINDICATO DEL SEG - Se aplico valor absoluto a los montos de los empleados

Global Const Version = "1.02"
Global Const FechaModificacion = "15/11/2013"
Global Const UltimaModificacion = " " 'LED - CAS-20033 - HDI - CUSTOM REPORTE DDJJ SINDICATO DEL SEG - Se aplico valor absoluto al monto total


'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Dim fs, f

Dim NroProceso As Long

Global Path As String
Global HuboErrores As Boolean
Global EmpErrores As Boolean

Global tenro1 As Long
Global estrnro1 As Long
Global tenro2 As Long
Global estrnro2 As Long
Global tenro3 As Long
Global estrnro3 As Long
Global fecEstr As String
Global fecEstr2 As String
Global Formato As Long
Global ListaProcesos As String
Global Modelo As Long
Global ModeloDesc As String
Global CantColumnas As Long
Global descDesde
Global descHasta
Global fechaHasta
Global fechaDesde
Global Nro_Col
Global listaPer
Global concAnt
Global Desde
Global Hasta
Global nomape
Global prog As Double



Private Sub Main()
Dim NombreArchivo As String
Dim directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String
Dim objRs As New ADODB.Recordset
Dim Parametros As String

Dim rep_DC_IDuser As String
Dim rep_DC_Fecha As String
Dim rep_DC_Hora As String

Dim strTempo As String
Dim orden
Dim rs_confrep

Dim ArrParametros
Dim PID As String


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
    
    Nombre_Arch = PathFLog & "DDJJSindicatos_" & NroProceso & ".log"
    
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
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID
    StrSql = StrSql & " WHERE btprcnro = 406 AND bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 406 AND bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
       'Obtengo los parametros del proceso
       Parametros = objRs!bprcparam
       rep_DC_IDuser = objRs!Iduser
       rep_DC_Fecha = objRs!bprcfecha
       rep_DC_Hora = objRs!bprchora
       Call GenerarReporte(NroProceso, Parametros, rep_DC_IDuser, rep_DC_Fecha, rep_DC_Hora)
    End If
   
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Fin :" & Now
    Flog.Close

    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline
    Flog.writeline " Error: " & Err.Description & Now
    Flog.writeline " Ultimo SQL: " & StrSql
    Flog.writeline
    MyRollbackTransliq
End Sub


Public Sub GenerarReporte(ByVal bpronro As Long, ByVal Parametros As String, ByVal Iduser As String, ByVal Fecha As String, ByVal hora As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : LED
' Fecha      : 05/09/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pliqnro
Dim Ternro
Dim totalempleados
Dim CantRegistros
Dim ArrParametros
Dim repnro
Dim repnrodet
Dim Legajo As Long
Dim apellido As String
Dim apellido2 As String
Dim nombre As String
Dim nombre2 As String
Dim tipoDoc As String
Dim nroDoc As String
Dim Cuil As String
Dim pliqmes As Long
Dim pliqanio As Long
Dim indice As Integer
Dim Valor As String
Dim descripcionConAcu As String
Dim estrnroEmpresa As Long
Dim tipocodigo As Long
Dim codempresa As String
Dim codmovimiento As String
Dim sucursal As String
Dim estrnrosucursal As Long
Dim cpsucursal As String
Dim codCategoria As String
Dim fechaIngreso As String
Dim fechaegreso As String
Dim codbaja As String
Dim codestado As String
Dim montoTotal As String
Dim montoRetenido As String
Dim rs As New ADODB.Recordset
Dim rsEmpl As New ADODB.Recordset
Dim rs_confrep As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset

    On Error GoTo MError
    
    Flog.writeline "Lista de Parametros = " & Parametros
    ArrParametros = Split(Parametros, "@")
           
    'Obtengo el periodo
    Flog.writeline "Obtengo el Período"
    pliqnro = CLng(ArrParametros(0))
    Flog.writeline "Período = " & pliqnro
    
    
    'Obtengo los cortes de estructura
    Flog.writeline "Obtengo los cortes de estructuras"
    
    Flog.writeline "Obtengo estructura 1"
    tenro1 = CLng(ArrParametros(1))
    estrnro1 = CLng(ArrParametros(2))
    Flog.writeline "Corte 1 = " & tenro1 & " - " & estrnro1
    
    Flog.writeline "Obtengo estructura 2"
    tenro2 = CLng(ArrParametros(3))
    estrnro2 = CLng(ArrParametros(4))
    Flog.writeline "Corte 2 = " & tenro2 & " - " & estrnro2
    
    Flog.writeline "Obtengo estructura 3"
    tenro3 = CLng(ArrParametros(5))
    estrnro3 = CLng(ArrParametros(6))
    Flog.writeline "Corte 3 = " & tenro3 & " - " & estrnro3
    
    
    Flog.writeline "Obtengo las Fechas Desde y Hasta"
    fecEstr = ArrParametros(7)
    fecEstr2 = ArrParametros(8)
    Flog.writeline "Fecha Desde = " & fecEstr
    Flog.writeline "Fecha Hasta = " & fecEstr2
    
    'Lista de procesos
    Flog.writeline "Obtengo la lista de procesos"
    ListaProcesos = "0,"
    ListaProcesos = ListaProcesos & ArrParametros(9)
    Flog.writeline "Lista de procesos = " & ListaProcesos
    
    'estrnro de la empresa
    Flog.writeline "Obtengo la empresa"
    estrnroEmpresa = ArrParametros(10)
    Flog.writeline "Empresa = " & estrnroEmpresa
        
    
    '============================================================================================
    'EMPIEZA EL PROCESO
    
    'Cargo la configuracion del reporte
    Flog.writeline "Cargo la Configuración del Reporte"
        
    'Obtengo los empleados sobre los que tengo que generar los recibos
    Flog.writeline "Cargo los Empleados "
    Call CargarEmpleados(NroProceso, rsEmpl, ListaProcesos)
    
    'Borro todos los registros (Para Reprocesamiento)--------------------------
    MyBeginTrans
        'Detalles de montos
        StrSql = "DELETE rep_ddjj_sindicato_det_det WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                      
        'Detalle
        StrSql = "DELETE rep_ddjj_sindicato_det WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
                      
        'Cabecera
        StrSql = "DELETE rep_ddjj_sindicato WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    'Borro todos los registros -------------------------------------------------
    
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                ", bprcempleados ='" & CStr(CantRegistros) & "' WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    prog = 0
    
    If (rsEmpl.RecordCount <> 0) Then
        Progreso = 100 / (rsEmpl.RecordCount)
        totalempleados = rsEmpl.RecordCount
        CantRegistros = rsEmpl.RecordCount
        rsEmpl.MoveFirst
    Else
        totalempleados = 0
        CantRegistros = 0
    End If
         
    If CLng(totalempleados) > 0 Then
        'Guardo en la BD el encabezado
        Flog.writeline "Genero el encabezado del Reporte"
        Call GenerarEncabezadoReporte(NroProceso, pliqnro, Iduser, Fecha, hora, totalempleados, tenro1, tenro2, tenro3, pliqmes, pliqanio)
        repnro = getLastIdentity(objConn, "rep_ddjj_sindicato")
    End If
    totalempleados = 0 'Recalculamos los empleados por si alguno no tiene documentos cargados
    montoTotal = 0
    Do Until rsEmpl.EOF
        prog = prog + Progreso
        EmpErrores = False
        Ternro = rsEmpl!Ternro
        
        Flog.writeline "--------------------------------------------------------------------------------"
        Flog.writeline "Comienza el analisis para el empleado, ternro: " & Ternro
        '-----------------------------------------------------------------------------------------------------------------
        'Busco los datos del empleado - (nombre, apellido, tipo y nro de documento (dni,le,lc), cuil)
        '-----------------------------------------------------------------------------------------------------------------
        StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfaltagr,empremu, doc.tidnro tipdocbasico, doc.nrodoc nrodocbasico " & _
                 " ,cuil.nrodoc cuil FROM empleado " & _
                 " INNER JOIN ter_doc doc ON doc.ternro = empleado.ternro AND doc.tidnro <= 3 " & _
                 " INNER JOIN ter_doc cuil ON cuil.ternro = empleado.ternro AND cuil.tidnro = 10 " & _
                 " WHERE empleado.ternro= " & Ternro
        
        Flog.writeline "-----------------------------------------------------------------"
        Flog.writeline "Buscando datos del empleado, ternro: " & Ternro
        
        OpenRecordset StrSql, rsConsult
        nombre = ""
        If Not rsConsult.EOF Then
           nombre = rsConsult!ternom
           'nomape = nombre
           If Not IsNull(rsConsult!ternom2) And rsConsult!ternom2 <> "" Then
              nombre = nombre & " " & rsConsult!ternom2
           End If
           
           apellido = rsConsult!terape
           
           'nomape = nomape & " " & apellido
           If Not IsNull(rsConsult!terape2) And rsConsult!terape2 <> "" Then
              apellido = apellido & " " & rsConsult!terape2
           End If
           Legajo = rsConsult!empleg
           
         tipoDoc = ""
         nroDoc = ""
         Cuil = ""
         Select Case CLng(rsConsult!tipdocbasico)
           Case 1:
               tipoDoc = "04"
           Case 2:
               tipoDoc = "01"
           Case 3:
               tipoDoc = "02"
           Case Else:
               Flog.writeline "El empleado no pose dni, le o lc."
               Flog.writeline "No se analizara el empleado."
               GoTo prox_emp
         End Select
                   
         nroDoc = Left(rsConsult!nrodocbasico, 8)
         If Not IsNull(rsConsult!Cuil) Then
           Cuil = Replace(rsConsult!Cuil, "-", "")
           Flog.writeline "Cuil del empleado encontrado."
         Else
            Flog.writeline "El empleado no posee cuil (tipo documento 10)"
            Flog.writeline "No se analizara el empleado."
            GoTo prox_emp
             
         End If
         totalempleados = totalempleados + 1
        Else
           Flog.writeline "Error al obtener los datos del empleado (revisar documento (DNI, LE, LC o CUIL))"
           Flog.writeline "No se analizara el empleado."
           GoTo prox_emp
        End If
        rsConsult.Close
        
        'Obtengo los datos de configuracion, la columna 1 la obtenemos despues ya que son los valores a calcular (acumuladores y conceptos)
        StrSql = " SELECT confnrocol, conftipo, confval2 " & _
                 " FROM confrep WHERE repnro = 411 AND confnrocol >= 2 ORDER BY confval "
        OpenRecordset StrSql, rs_confrep
        If Not rs_confrep.EOF Then

            Do While Not rs_confrep.EOF
                Select Case CLng(rs_confrep!confnrocol)
                    Case 2: 'tipo de codigo
                        tipocodigo = rs_confrep!confval2
                End Select
                rs_confrep.MoveNext
            Loop
                        
        Else
            Flog.writeline "No hay datos configurados en el reporte."
        End If
        
        'recupero el codigo asociado a la empresa
        StrSql = " SELECT nrocod FROM estr_cod WHERE tcodnro = " & tipocodigo & " AND estrnro = " & estrnroEmpresa
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            codempresa = Right("000000" & rsConsult!nrocod, 6)
        Else
            codempresa = ""
            Flog.writeline "La empresa no tiene codigo asociado."
        End If
                                
        'recupero la categoria del empleado
        StrSql = " SELECT his_estructura.estrnro, nrocod FROM his_estructura " & _
                 " INNER JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro AND tenro = 3 AND tcodnro = " & tipocodigo & _
                 " WHERE ternro = " & Ternro & _
                 " AND ((his_estructura.htetdesde <= " & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fecEstr2) & _
                 " or his_estructura.htethasta >= " & ConvFecha(fecEstr) & ")) OR(his_estructura.htetdesde >= " & ConvFecha(fecEstr) & " AND (his_estructura.htetdesde <= " & ConvFecha(fecEstr2) & "))) "
        
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            codCategoria = Right("000" & rsConsult!nrocod, 3)
        Else
            codCategoria = ""
            Flog.writeline "El empleado no posee categoria en el periodo o no tiene el codigo cargado el tipo de codigo " & tipocodigo & "."
        End If
                
        'recupero la sucursal del empleado
        StrSql = " SELECT estrdabr, codigopostal FROM his_estructura " & _
                 " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro AND his_estructura.tenro = 1 " & _
                 " INNER JOIN sucursal ON sucursal.estrnro = his_estructura.estrnro " & _
                 " INNER JOIN cabdom ON cabdom.ternro = sucursal.ternro AND cabdom.domdefault = -1 " & _
                 " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
                 " WHERE his_estructura.ternro = " & Ternro & _
                 " AND ((his_estructura.htetdesde <= " & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fecEstr2) & _
                 " or his_estructura.htethasta >= " & ConvFecha(fecEstr) & ")) OR(his_estructura.htetdesde >= " & ConvFecha(fecEstr) & " AND (his_estructura.htetdesde <= " & ConvFecha(fecEstr2) & "))) "
        
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            sucursal = Left(rsConsult!estrdabr, 20)
            cpsucursal = IIf(EsNulo(rsConsult!codigopostal), "", rsConsult!codigopostal)
        Else
            sucursal = ""
            cpsucursal = ""
            Flog.writeline "El empleado no posee sucursal, no tiene cargado el domicilio."
        End If
                      
        'busco la face activa del empleado
         StrSql = " SELECT altfec, bajfec, caucod FROM fases " & _
                  " LEFT JOIN causa ON causa.caunro = fases.caunro " & _
                  " WHERE empleado = " & Ternro & _
                  " AND ((altfec <= " & ConvFecha(fecEstr) & " AND (bajfec is null or bajfec >= " & ConvFecha(fecEstr2) & _
                  " or bajfec >= " & ConvFecha(fecEstr) & ")) OR(altfec >= " & ConvFecha(fecEstr) & " AND (altfec <= " & ConvFecha(fecEstr2) & "))) "
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            fechaIngreso = Replace(rsConsult!altfec, "/", "")
            If IsNull(rsConsult!bajfec) Then
                fechaegreso = "00000000"
                codbaja = "00"
            Else
                fechaegreso = Replace(rsConsult!bajfec, "/", "")
                codbaja = Right("00" & rsConsult!caucod, 2)
            End If
        Else
            Flog.writeline "El empleado no posee fase para el periodo."
            fechaIngreso = ""
            fechaegreso = ""
            codbaja = ""
        End If
        
        'El codigo de estado es un tipo de codigo 172 asociado al tipo de estructura situacion de revista
        StrSql = " SELECT his_estructura.estrnro, nrocod FROM his_estructura " & _
                 " INNER JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro AND tenro = 30 AND tcodnro = " & tipocodigo & _
                 " WHERE ternro = " & Ternro & _
                 " AND ((his_estructura.htetdesde <= " & ConvFecha(fecEstr) & " AND (his_estructura.htethasta is null or his_estructura.htethasta >= " & ConvFecha(fecEstr2) & _
                 " or his_estructura.htethasta >= " & ConvFecha(fecEstr) & ")) OR(his_estructura.htetdesde >= " & ConvFecha(fecEstr) & " AND (his_estructura.htetdesde <= " & ConvFecha(fecEstr2) & "))) "
        
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            codestado = Right("00" & rsConsult!nrocod, 2)
        Else
            codestado = "00"
            Flog.writeline "El empleado no posee categoria en el periodo o no tiene el codigo cargado el tipo de codigo " & tipocodigo & "."
        End If
            
        StrSql = " INSERT INTO rep_ddjj_sindicato_det (repddjjnro,bpronro,ternro,empleg,cuil,tipodocumento " & _
                 " ,documento,apellido,nombre,sucursal,cpsucursal,codCategoria " & _
                 " ,fechaIngreso,fechaegreso,codbaja,codestado) VALUES ( " & repnro & _
                 " ," & bpronro & "," & Ternro & "," & Legajo & " ,'" & Cuil & "','" & tipoDoc & "'" & _
                 " ,'" & nroDoc & "','" & apellido & "','" & nombre & "','" & sucursal & "','" & cpsucursal & "'" & _
                 " ,'" & codCategoria & "','" & fechaIngreso & "','" & fechaegreso & "','" & codbaja & "','" & codestado & "')"
                 
        objConn.Execute StrSql, , adExecuteNoRecords
        repnrodet = getLastIdentity(objConn, "rep_ddjj_sindicato_det")
        
        '--------------------------------------------- valores del confrep --------------------------------------------
        
        StrSql = " SELECT (select COUNT(confnrocol) from confrep  where repnro = 411 AND confnrocol = 1) cantCol, confnrocol, conftipo, confval2 " & _
                 " FROM confrep WHERE repnro = 411 AND confnrocol = 1 ORDER BY confval "
        OpenRecordset StrSql, rs_confrep
        If Not rs_confrep.EOF Then
            indice = 0
            ReDim arrConfrepCodigo(rs_confrep!cantCol - 1)
            ReDim arrConfrepTipo(rs_confrep!cantCol - 1)
            Do While Not rs_confrep.EOF
                arrConfrepCodigo(indice) = rs_confrep!confval2      'en la posicion 1 se guardan conceptos o acumuladores
                arrConfrepTipo(indice) = rs_confrep!conftipo        'en la posicion 1 se guardan conceptos o acumuladores
                rs_confrep.MoveNext
                indice = indice + 1
            Loop
                        
        Else
            Flog.writeline "No hay datos configurados en el reporte."
        End If
        
        montoRetenido = 0
        'obtengo los valores y los escribo en la tabla
        For indice = 0 To UBound(arrConfrepCodigo)
            Valor = buscarConceptoAcumPorEtiquetaEnProcesos(arrConfrepTipo(indice), Ternro, arrConfrepCodigo(indice), pliqmes, pliqanio, ListaProcesos)
            
            'en indice 12 (acumulador o concepto nro 13) estoy sobre el monto total del empleado
            If indice = 12 Then
                montoRetenido = Valor
                If Valor < 0 Then
                    codmovimiento = "1" 'retencion
                Else
                    codmovimiento = "0" 'reintegro
                End If
            End If
            
            Valor = FormatNumber(Abs(Valor), 2)
            Valor = Replace(Valor, ",", "")
            'Valor = Replace(Valor, ".", "")
            descripcionConAcu = buscarDescripcionAcuCon(arrConfrepTipo(indice), arrConfrepCodigo(indice))
            StrSql = " INSERT INTO rep_ddjj_sindicato_det_det (repddjjdetnro , bpronro, Ternro, Origen, descripcion, Monto) VALUES " & _
                     " (" & repnrodet & "," & bpronro & "," & Ternro & ",'" & arrConfrepCodigo(indice) & "','" & descripcionConAcu & "'," & Valor & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Insertado los datos para el concepto/acumulador con tipo: " & arrConfrepTipo(indice) & " codigo: " & arrConfrepCodigo(indice) & "."
        Next
        
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
    
        CantRegistros = CantRegistros - 1
    
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & prog
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(CantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords

      'Si se generaron todos los datos del empleado correctamente lo borro
      If Not EmpErrores Then
          StrSql = " DELETE FROM batch_empleado "
          StrSql = StrSql & " WHERE bpronro = " & NroProceso
          StrSql = StrSql & " AND ternro = " & Ternro
          objConn.Execute StrSql, , adExecuteNoRecords
      End If
      montoTotal = FormatNumber(CDbl(montoTotal) + CDbl(montoRetenido), 2)

    'actualizo el codigo del movimiento
    StrSql = " UPDATE rep_ddjj_sindicato_det SET " & _
         " codmovimiento = '" & codmovimiento & "'" & _
         " WHERE bpronro = " & NroProceso & "AND ternro =  " & Ternro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

prox_emp:
        rsEmpl.MoveNext
    Loop
    
  montoTotal = FormatNumber(Abs(montoTotal), 2)
    montoTotal = Replace(montoTotal, ",", "")
    'montoTotal = Replace(montoTotal, ".", "")
    'actualizo valores de la cabecera ya calculados
    StrSql = " UPDATE rep_ddjj_sindicato SET " & _
             " totalempleados = " & totalempleados & "," & _
             " aniopago = " & pliqanio & "," & _
             " mespago = " & pliqmes & "," & _
             " anioretencion = " & pliqanio & "," & _
             " mesretencion = " & pliqmes & "," & _
             " codempresa = '" & codempresa & "'," & _
             " montoTotal = '" & montoTotal & "'" & _
             " WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Se llama la funcion que exporta al archivo
    generarArchivo (NroProceso)

Exit Sub

MError:
    Flog.writeline "Error al generando el reporte. Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    MyRollbackTrans
    Exit Sub
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


Sub CargarEmpleados(NroProc, ByRef rsEmpl As ADODB.Recordset, procesos As String)
'--------------------------------------------------------------------
' Se encarga de generar un ResultSet de los empleados a cambiar
' si el RS es vacio significa que hay que aplicarlo sobre todos
'--------------------------------------------------------------------
Dim StrEmpl As String

    StrEmpl = " SELECT DISTINCT ternro  FROM batch_empleado " & _
              " INNER JOIN cabliq ON cabliq.empleado = ternro " & _
              " WHERE bpronro = " & NroProc & " AND pronro in (" & procesos & ") " & _
              " ORDER BY ternro "
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




Sub GenerarEncabezadoReporte(ByVal bpronro As Long, ByVal pliqnro As Long, ByVal Iduser As String, ByVal Fecha As String, ByVal hora As String, ByVal totalempleados As Long, ByVal tenro1 As Long, ByVal tenro2 As Long, ByVal tenro3 As Long, ByRef pliqmes, ByRef pliqanio)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : LED
' Fecha      : 19/02/2013
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim teNomb1 As String
Dim teNomb2 As String
Dim teNomb3 As String
Dim secuencia As Long

Dim I
Dim TituloRep As String

Dim rsConsult As New ADODB.Recordset


teNomb1 = ""
teNomb2 = ""
teNomb3 = ""

If tenro1 <> 0 Then
    StrSql = " SELECT tedabr "
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
    StrSql = " SELECT tedabr "
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
    StrSql = " SELECT tedabr "
    StrSql = StrSql & " FROM tipoestructura "
    StrSql = StrSql & "  WHERE tipoestructura.tenro = " & tenro3
    OpenRecordset StrSql, rsConsult
    
    If Not rsConsult.EOF Then
       teNomb3 = rsConsult!tedabr
    Else
       teNomb3 = ""
    End If
End If

'Descripcion del historico del reporte

    TituloRep = TituloRep & bpronro & " - "
    
    StrSql = " SELECT pliqdesc, pliqmes, pliqanio FROM periodo "
    StrSql = StrSql & "  WHERE pliqnro = " & pliqnro
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
       TituloRep = TituloRep & rsConsult!PliqDesc
       pliqmes = rsConsult!pliqmes
       pliqanio = rsConsult!pliqanio
    End If
    TituloRep = TituloRep & " - " & Fecha
    TituloRep = TituloRep & " " & hora

'busco la ultima secuencia creada para el dia
    StrSql = " SELECT max(secuencia) sec FROM rep_ddjj_sindicato "
    StrSql = StrSql & "  WHERE fecha = '" & Replace(Fecha, "/", "") & "'"
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        If IsNull(rsConsult!sec) Then
            secuencia = 0
        Else
            secuencia = CLng(rsConsult!sec) + 1
        End If
    Else
        secuencia = 0
    End If


    StrSql = " INSERT INTO rep_ddjj_sindicato (bpronro,repdesc,rep_user,fecha,hora,secuencia,pliqnro,pliqmes,pliqanio,totalempleados,tedabr1,tedabr2,tedabr3) VALUES ( "
    StrSql = StrSql & NroProceso
    StrSql = StrSql & ",'" & TituloRep & "'"
    StrSql = StrSql & ",'" & Iduser & "'"
    StrSql = StrSql & ",'" & Replace(Fecha, "/", "") & "'"
    StrSql = StrSql & ",'" & Replace(hora, ":", "") & "'"
    StrSql = StrSql & "," & secuencia
    StrSql = StrSql & "," & pliqnro
    StrSql = StrSql & "," & pliqmes
    StrSql = StrSql & "," & pliqanio
    StrSql = StrSql & "," & totalempleados
    StrSql = StrSql & ",'" & teNomb1 & "'"
    StrSql = StrSql & ",'" & teNomb2 & "'"
    StrSql = StrSql & ",'" & teNomb3 & "'"
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords



End Sub

Sub generarArchivo(ByVal bpronro As Long)
Dim Nombre_Arch As String
Dim dirsalidas As String
Dim secuencia As Long
Dim archsalida
Dim strlinea As String
Dim repddjjnro As Long
Dim repddjjdetnro As Long
Dim montoTotal As String
Dim codigoMontoTotal As Long
Dim totalempleados As Long
Dim rsConsult As New ADODB.Recordset
Dim rsConsultDet As New ADODB.Recordset

Set fs = CreateObject("Scripting.FileSystemObject")

' Directorio Salidas
StrSql = "SELECT sis_dirsalidas FROM sistema"
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    dirsalidas = objRs!sis_dirsalidas & "\DDJJ_Sindicato"
    Flog.writeline "Directorio de Salidas: " & dirsalidas
Else
    Flog.writeline "No se encuentra configurado sis_dirsalidas"
    Exit Sub
End If
If objRs.State = adStateOpen Then objRs.Close


  StrSql = " SELECT  repddjjnro, fecha, hora, aniopago, mespago, secuencia, anioretencion, mesretencion, codempresa, totalempleados, montototal " & _
            " FROM rep_ddjj_sindicato " & _
            " WHERE rep_ddjj_sindicato.bpronro = " & bpronro
   
   OpenRecordset StrSql, rsConsult
      
      
   If rsConsult.EOF Then
       Flog.writeline "No Hay datos para generar el archivo."
       Exit Sub
   End If
   
    secuencia = rsConsult!secuencia
    'mediante la secuencia q se calculo en el encabezado controlo si el archivo ya existe
    Nombre_Arch = "DDJJ_" & Right("00" & CStr(secuencia), 2) & "_" & Year(Date) & Right("00" & Month(Date), 2) & Right("00" & Day(Date), 2) & ".csv"
    
    If existe_archivo(Nombre_Arch, dirsalidas) Then
        Flog.writeline "Ya existe el archivo de salida."
        Exit Sub
    End If
    
    If Right(dirsalidas, 1) = "\" Then
        Nombre_Arch = dirsalidas & Nombre_Arch
    Else
        Nombre_Arch = dirsalidas & "\" & Nombre_Arch
    End If
    
    Set archsalida = fs.CreateTextFile(Nombre_Arch, True)
      
   repddjjnro = rsConsult!repddjjnro
   totalempleados = rsConsult!totalempleados 'recupero el total de empleados para imprmirlo en el pie
   montoTotal = rsConsult!montoTotal 'recupero el monto total para imprmirlo en el pie
   'Primer registro
   strlinea = "01"                                                              'cabezal
   strlinea = strlinea & "," & Right("00" & CStr(secuencia), 2)                 'secuencia
   strlinea = strlinea & "," & rsConsult!Fecha                                  'fecha de generacion
   strlinea = strlinea & "," & rsConsult!hora                                   'Hora de generacion
   strlinea = strlinea & "," & Right("0000" & CStr(rsConsult!aniopago), 4)        'Año de pago
   strlinea = strlinea & "," & Right("00" & CStr(rsConsult!mespago), 2)         'Mes de pago
   strlinea = strlinea & "," & Right("0000" & CStr(rsConsult!anioretencion), 4)   'Año de retencion
   strlinea = strlinea & "," & Right("00" & CStr(rsConsult!mesretencion), 2)    'Mes de retencion
   strlinea = strlinea & "," & rsConsult!codempresa                             'Codigo de la empresa
   strlinea = strlinea & "," & totalempleados                                   'Total de trabajadores de la actividad aseguradora
      
   archsalida.writeline strlinea
   Flog.writeline "Generada cabecera del archivo."
    
    StrSql = " SELECT repddjjdetnro, ternro, codmovimiento, cuil,tipodocumento, documento, apellido, nombre,sucursal, cpsucursal " & _
             " , empleg, codcategoria, fechaIngreso, fechaegreso, codbaja, codestado " & _
             " FROM rep_ddjj_sindicato_det " & _
             " WHERE repddjjnro = " & repddjjnro
   
   OpenRecordset StrSql, rsConsult

    Do While Not rsConsult.EOF
        Flog.writeline "Generando Registro de afiliados para el legajo: " & rsConsult!empleg & "."
        'Registro de afiliados
        repddjjdetnro = rsConsult!repddjjdetnro
        strlinea = "02"
        strlinea = strlinea & "," & rsConsult!codmovimiento                          'Codigo de movimiento
        strlinea = strlinea & "," & rsConsult!Cuil                                   'cuil
        strlinea = strlinea & "," & rsConsult!tipodocumento                          'Tipo de Documento
        strlinea = strlinea & "," & Right("00000000" & rsConsult!documento, 8)       'Nro de Documento
        strlinea = strlinea & "," & rsConsult!apellido                               'Apellidos
        strlinea = strlinea & "," & rsConsult!nombre                                 'Nombres
        strlinea = strlinea & "," & Left(rsConsult!sucursal, 20)                     'sucursal
        strlinea = strlinea & "," & Left(rsConsult!cpsucursal, 8)                    'codigo postal sucursal
        strlinea = strlinea & "," & Right("00000000" & rsConsult!empleg, 6)          'legajo del empleado
        strlinea = strlinea & "," & rsConsult!codCategoria                           'codigo de la categoria del empleado
        
        strlinea = strlinea & "," & rsConsult!fechaIngreso                           'Fecha de ingreso del empleado
        strlinea = strlinea & "," & rsConsult!fechaegreso                            'Fecha de egreso del empleado
        strlinea = strlinea & "," & rsConsult!codbaja                                'codigo de la categoria del empleado
        strlinea = strlinea & "," & rsConsult!codestado                           'codigo de la categoria del empleado
        
        StrSql = " SELECT origen, monto " & _
                 " FROM rep_ddjj_sindicato_det_det " & _
                 " WHERE repddjjdetnro = " & repddjjdetnro & " AND ternro = " & rsConsult!Ternro
        OpenRecordset StrSql, rsConsultDet
        Flog.writeline "Generando montos para el legajo: " & rsConsult!empleg & "."
        Do While Not rsConsultDet.EOF

            strlinea = strlinea & "," & Replace(Replace(FormatNumber(rsConsultDet!Monto, 2), ",", ""), ".", "")
            'Montos configurados en la columna 1
            rsConsultDet.MoveNext
        Loop
        archsalida.writeline strlinea
        Flog.writeline "Montos para el legajo: " & rsConsult!empleg & " generados con exito."
        Flog.writeline "Registro de afiliados para el legajo: " & rsConsult!empleg & " generado con exito."
        rsConsult.MoveNext
    Loop
    
    Flog.writeline "Generando ultimo registro."
    strlinea = "03"
    strlinea = strlinea & "," & CStr(CLng(totalempleados) + 2)
    strlinea = strlinea & "," & Replace(Replace(FormatNumber(montoTotal, 2), ",", ""), ".", "")
    archsalida.writeline strlinea
    Flog.writeline "Ultimo registro generado por exito."
archsalida.Close

End Sub

'Funcion que Valida si existe un archivo.
Function existe_archivo(ByVal Filename As String, ByVal ubioriginal As String)
    Dim sFilename
    Dim oFSO
    Dim pos
    Dim directorio
    Dim fso
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set directorio = fso.GetFolder(ubioriginal)
    
    If Right(directorio, 1) = "\" Then
        sFilename = directorio & Filename
    Else
        sFilename = directorio & "\" & Filename
    End If

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    If oFSO.FileExists(sFilename) Then
        'El Archivo Existe
        existe_archivo = True
    Else
        'El Archivo NO Existe
        existe_archivo = False
    End If
End Function 'existe_archivo(Filename,l_ubioriginal)
