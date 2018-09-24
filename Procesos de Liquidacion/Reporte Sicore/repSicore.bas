Attribute VB_Name = "repSicore"
Option Explicit

'Const Version = 1.04 'Version Inicial
'Const FechaVersion = "10/03/2006" - FGZ - Sacaba mal la parte decimal cuando tenia un solo decimal

'Const Version = 1.05 'Agregado de log cuando calcula la suma de los detalles de ganancias y el acumulador imponible de ganancia
'Const FechaVersion = "14/03/2006"

'Const Version = 1.06 'Agregado de log de la configuracion regional, cuando hace el insert, etc
'Const FechaVersion = "21/03/2006"

'Const Version = 1.07 'Se saco def de la variable Flog y ahora se tienen en cta los domicilios que son de empleados que fueron postulantes
'Const FechaVersion = "25/03/2009"

'Const Version = 1.08 'Se agrego la lectura de nuevos parametros de CONFREP para recuperar informacion de Operaciones con Beneficiarios del Exterior
'Const FechaVersion = "02/02/2011"

'Const Version = 1.09 'Se agrego la lectura de nuevos parametros de CONFREP para recuperar informacion de Operaciones con Beneficiarios del Exterior
'Const FechaVersion = "04/02/2014"  ' MDF -CAS-23749 - H&A - Cambio legal - Ganancias - SICORE v8 r13- nuevo calculo para ganancias imponibles, si se pone -1 en columna 5 de confrep trae el valor de traza_gan

'Const Version = 1.1   'Se agrego la lectura de nuevos parametros de CONFREP para recuperar informacion de Operaciones con Beneficiarios del Exterior
'Const FechaVersion = "11/02/2014"  ' MDF -CAS-23749 - H&A - Cambio legal - Ganancias - SICORE v8 r13- en caso de traer dos registros, se queda con el el que tiene maximo valor.

'Const Version = 1.11   'Se agrego la lectura de nuevos parametros de CONFREP para recuperar informacion de Operaciones con Beneficiarios del Exterior
'Const FechaVersion = "17/02/2014"  ' MDF -CAS-23749 - H&A - Cambio legal - Ganancias - SICORE v8 r13- cuando el signo es 8 cambian las ganancias imponibles

'Const Version = 1.12   'Se cambio el codigo del impuesto al valor 787
'Const FechaVersion = "03/02/2015"  ' Sebastian Stremel - CAS-29236 - HYA - RG 3731 SICORE

'Const Version = 1.13   'Se cambio de NroProcesoBatch a NroProceso cuando cambia el estado del proceso a procesando
'Const FechaVersion = "13/03/2015"  ' Borrelli Facundo - CAS-29635 - G.Compartida Edenor Convenio - Error en Task Manager

'Const Version = 1.14   'Se desglozan los conceptos por Retenciones y Devoluciones
'Const FechaVersion = "05/05/2015"  ' Dimatz Rafael - CAS-30536 - RHPRO - Modificacion en SICORE

Const Version = 1.15   'Se desglozan los conceptos por Retenciones y Devoluciones y se corrige insertando fec_ret_CAA
Const FechaVersion = "20/05/2015"  ' Dimatz Rafael - CAS-30536 - RHPRO - Modificacion en SICORE

'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

Private Type TConfrepCO
    ConcCod As String
    Conccod_txt As String
    ConcNro As Long
End Type

'FGZ - 16/11/2004
Dim Concepto1() As TConfrepCO
Dim Concepto2() As TConfrepCO
Dim Concepto3() As TConfrepCO
Dim Indice1 As Integer
Dim Indice2 As Integer
Dim Indice3 As Integer

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

Dim codobe1
Dim codobe2
Dim codobe3
Dim codtcod ' Codigo del tipo de codigo asociado a las estructuras para OBE!


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
Dim fechadesde
Dim fechahasta
Dim arr
Dim empresa
Dim incoperben
Dim tipoDepuracion
Dim param
Dim Ternro
Dim arrpronro
Dim rsEmpl As New ADODB.Recordset
Dim rsleg  As New ADODB.Recordset
Dim Sql_leg As String
Dim legemp As Long

Dim concod1
Dim concod2
Dim concod3
Dim concod1_txt As String
Dim concod2_txt As String
Dim concod3_txt As String
Dim nro_acu_neto
Dim nro_acu_imponible

Dim auxcon1
Dim auxcon2
Dim auxcon3

Dim cliqnro

Dim Retencion
Dim EmpAnt
Dim profecpago
Dim cantRegistros
Dim actualReg

Dim TiempoInicialProceso
Dim TiempoAcumulado
Dim PID As String
Dim bprcparam As String

Dim I As Integer
Dim ArrParametros
   
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
    
    'Obtiene los datos de como esta configurado el servidor actualmente
    Call ObtenerConfiguracionRegional
    
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
    
    On Error GoTo CE
    
    TiempoInicialProceso = GetTickCount
    
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "Sicore" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
'    Flog.Writeline "Inicio Sicore : " & Now
'    Flog.Writeline "Cambio el estado del proceso a Procesando"
    
    concod1 = "0"
    concod2 = "0"
    concod3 = "0"
    nro_acu_neto = "0"
    nro_acu_imponible = "0"
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Numero, separador decimal    : " & NumeroSeparadorDecimal
    Flog.writeline "Numero, separador de miles   : " & NumeroSeparadorMiles
    Flog.writeline "Moneda, separador decimal    : " & MonedaSeparadorDecimal
    Flog.writeline "Moneda, separador de miles   : " & MonedaSeparadorMiles
    Flog.writeline "Formato de Fecha del Servidor: " & FormatoDeFechaCorto
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso 'NroProcesoBatch  LM
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 28 AND bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
 
    If Not objRs.EOF Then
       'Obtengo los parametros del proceso
       fechadesde = objRs!bprcfecdesde
       fechahasta = objRs!bprcfechasta
       Flog.writeline
       Flog.writeline "Fecha desde " & fechadesde
       Flog.writeline "Fecha hasta " & fechahasta
       arr = Split(objRs!bprcparam, ".")
       empresa = arr(0)
        ' GdeCos - Si el estrnro de la empresa es 0, entonces se buscan todas las empresas
        If empresa = 0 Then
            Flog.writeline "Parametro Empresa = Todas"
        Else
            Flog.writeline "Parametro Empresa = " & empresa
        End If
       incoperben = arr(1)
       
       'EMPIEZA EL PROCESO
       
        'FGZ - 16/11/2004
        'Cuento la cantidad de conceptos 1
'        StrSql = " SELECT * FROM confrep "
'        StrSql = StrSql & " WHERE repnro = 14 "
'        StrSql = StrSql & " AND confnrocol = 1 "
'        If objRs2.State = adStateOpen Then objRs2.Close
'        OpenRecordset StrSql, objRs2
'        ReDim Preserve Concepto1(objRs2.RecordCount) As TConfrepCO
        
'        'Cuento la cantidad de conceptos 2
'        StrSql = " SELECT * FROM confrep "
'        StrSql = StrSql & " WHERE repnro = 14 "
'        StrSql = StrSql & " AND confnrocol = 2 "
'        If objRs2.State = adStateOpen Then objRs2.Close
'        OpenRecordset StrSql, objRs2
'        ReDim Preserve Concepto2(objRs2.RecordCount) As TConfrepCO
        
'        'Cuento la cantidad de conceptos 3
'        StrSql = " SELECT * FROM confrep "
'        StrSql = StrSql & " WHERE repnro = 14 "
'        StrSql = StrSql & " AND confnrocol = 3 "
'        If objRs2.State = adStateOpen Then objRs2.Close
'        OpenRecordset StrSql, objRs2
'        ReDim Preserve Concepto3(objRs2.RecordCount) As TConfrepCO
        
        StrSql = " SELECT * FROM confrep "
        StrSql = StrSql & " WHERE repnro = 14 "
        StrSql = StrSql & " ORDER BY confnrocol"
        If objRs2.State = adStateOpen Then objRs2.Close
        OpenRecordset StrSql, objRs2
        Indice1 = -1
        Indice2 = -1
        Indice3 = -1
        codobe1 = 0
        codobe2 = 0
        codobe3 = 0
        codtcod = 0
        Flog.writeline "Obtengo los datos del confrep"
        If Not objRs2.EOF Then
            Do Until objRs2.EOF
                Select Case objRs2!confnrocol
'                Case 1
'                    Indice1 = Indice1 + 1
'                    Concepto1(Indice1).ConcCod = CStr(objRs2!confval)
'                    Concepto1(Indice1).Conccod_txt = IIf(Not EsNulo(objRs2!confval2), objRs2!confval2, "")
'                    Concepto1(Indice1).ConcNro = 0
'                Case 2
'                    Indice2 = Indice2 + 1
'                    Concepto2(Indice2).ConcCod = CStr(objRs2!confval)
'                    Concepto2(Indice2).Conccod_txt = IIf(Not EsNulo(objRs2!confval2), objRs2!confval2, "")
'                    Concepto2(Indice2).ConcNro = 0
'                Case 3
'                    Indice3 = Indice3 + 1
'                    Concepto3(Indice3).ConcCod = CStr(objRs2!confval)
'                    Concepto3(Indice3).Conccod_txt = IIf(Not EsNulo(objRs2!confval2), objRs2!confval2, "")
'                    Concepto3(Indice3).ConcNro = 0
                Case 4
                    nro_acu_neto = objRs2!confval
                Case 5
                    nro_acu_imponible = objRs2!confval
                Case 50
                    codobe1 = objRs2!confval
                Case 51
                    codobe2 = objRs2!confval
                Case 52
                    codobe3 = objRs2!confval
                Case 53
                    codtcod = objRs2!confval
                End Select
               objRs2.MoveNext
            Loop
        Else
           Flog.writeline "No esta configurado el ConfRep"
           Exit Sub
        End If
        If objRs2.State = adStateOpen Then objRs2.Close

        Call comprobarCodigosOBE(codobe1, codobe2, codobe3, codtcod)
        

       '----------------------------------------------------------------------
       'Busco los conceptos definidos en el confrep
       '----------------------------------------------------------------------
       Flog.writeline "Buscando los conceptos del confrep"
           
'--------------- Asigna el Concnro a Concepto1() Concepto2() Concepto3() ----------------------
        'FGZ - 08/11/2004 - comparo por el valor numerico o alfanumerico
'        For I = 0 To Indice1
'            StrSql = "SELECT * FROM concepto "
'            StrSql = StrSql & " WHERE (concepto.conccod = " & Concepto1(I).ConcCod
'            StrSql = StrSql & " OR concepto.conccod = '" & Concepto1(I).Conccod_txt & "')"
'            If objRs2.State = adStateOpen Then objRs2.Close
'            OpenRecordset StrSql, objRs2
'
'            If Not objRs2.EOF Then
'                'auxcon1 = objRs2!Concnro
'                Concepto1(I).ConcNro = objRs2!ConcNro
'            Else
'                Flog.writeline "Conceptos No encontrado. " & Concepto1(I).ConcCod
'            End If
'        Next I
'        If objRs2.State = adStateOpen Then objRs2.Close
'
'       '----------------------------------------------------------------------
'        'FGZ - 08/11/2004 - comparo por el valor numerico o alfanumerico
'        For I = 0 To Indice2
'            StrSql = "SELECT * FROM concepto "
'            StrSql = StrSql & " WHERE (concepto.conccod = " & Concepto2(I).ConcCod
'            StrSql = StrSql & " OR concepto.conccod = '" & Concepto2(I).Conccod_txt & "')"
'            If objRs2.State = adStateOpen Then objRs2.Close
'            OpenRecordset StrSql, objRs2
'
'            If Not objRs2.EOF Then
'                'auxcon2 = objRs2!Concnro
'                Concepto2(I).ConcNro = objRs2!ConcNro
'            Else
'                Flog.writeline "Conceptos No encontrado. " & Concepto2(I).ConcCod
'            End If
'        Next I
'        If objRs2.State = adStateOpen Then objRs2.Close
'
'       '----------------------------------------------------------------------
'
'        'FGZ - 08/11/2004 - comparo por el valor numerico o alfanumerico
'        For I = 0 To Indice3
'             StrSql = "SELECT * FROM concepto "
'             StrSql = StrSql & " WHERE (concepto.conccod = " & Concepto3(I).ConcCod
'             StrSql = StrSql & " OR concepto.conccod = '" & Concepto3(I).Conccod_txt & "')"
'             If objRs2.State = adStateOpen Then objRs2.Close
'             OpenRecordset StrSql, objRs2
'
'             If Not objRs2.EOF Then
'                 Concepto3(I).ConcNro = objRs2!ConcNro
'             Else
'                Flog.writeline "Conceptos No encontrado. " & Concepto3(I).ConcCod
'             End If
'        Next I
'        If objRs2.State = adStateOpen Then objRs2.Close
''--------------- Asigna el Concnro a Concepto1() Concepto2() Concepto3() ----------------------
'
'       'Borro de la table sicore los datos que existan para la fecha ingresada
'       ' 30/05/2005 - GdeCos - Al borrar, No se filtra por empresa si empresa=0
       StrSql = "DELETE FROM sicore WHERE fec_desde = " & ConvFecha(fechadesde) & " AND fec_hasta = " & ConvFecha(fechahasta) & _
                " AND inc_oper_ben=" & incoperben
        If empresa <> 0 Then
            StrSql = StrSql & " AND empresa = " & empresa
        End If

       objConn.Execute StrSql, , adExecuteNoRecords
       
       Flog.writeline "Borro los datos del Historico para el rango de fechas"
       
       'Obtengo la lista de datos sobre la que tengo que procesar
       
        StrSql = " SELECT cabliq.cliqnro,cabliq.empleado,profecpago,his_estructura.estrnro "
        StrSql = StrSql & " FROM proceso "
        StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
'       StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= " & empresa & " AND htetdesde <= profecpago AND (htethasta IS NULL OR htethasta >= profecpago) AND his_estructura.tenro=10 AND his_estructura.ternro= cabliq.empleado"
       ' 30/05/2005 - GdeCos - Si empresa=0 se consideran todas las empresas
        If empresa <> 0 Then
            StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.estrnro= " & empresa & " AND htetdesde <= " & ConvFecha(fechahasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fechahasta) & ") AND his_estructura.tenro=10 AND his_estructura.ternro= cabliq.empleado"
        Else
            StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro= cabliq.empleado AND htetdesde <= " & ConvFecha(fechahasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fechahasta) & ") AND his_estructura.tenro=10"
        End If
        StrSql = StrSql & " WHERE proceso.profecpago >= " & ConvFecha(fechadesde) & " AND profecpago <= " & ConvFecha(fechahasta)
        StrSql = StrSql & " ORDER BY cabliq.empleado,cabliq.cliqnro "
        Flog.writeline "Obtengo la lista de datos sobre la que tengo que procesar"
        Flog.writeline StrSql
        OpenRecordset StrSql, rsEmpl

       'Genero por cada registro una linea

       cantRegistros = rsEmpl.RecordCount
       TiempoAcumulado = GetTickCount
       
       StrSql = "UPDATE batch_proceso SET bprcprogreso = 0 " & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
       'objConn.Execute StrSql, , adExecuteNoRecords
       objconnProgreso.Execute StrSql, , adExecuteNoRecords
       
       actualReg = 1

       Do Until rsEmpl.EOF
          EmpErrores = False
          Ternro = rsEmpl!Empleado
          cliqnro = rsEmpl!cliqnro
          profecpago = rsEmpl!profecpago
       ' 30/05/2005 - GdeCos
          empresa = rsEmpl!Estrnro
          
       ' 17/08/2005 - M Ferraro - Se busca el legajo para el log
          legemp = 0
          Sql_leg = "select empleg from empleado where empleado.ternro = " & Ternro
          OpenRecordset Sql_leg, rsleg
          If Not rsleg.EOF Then
            legemp = rsleg!empleg
          End If
          rsleg.Close
          
          Flog.writeline "Generando linea para el empleado empleado " & legemp & " y cliqnro " & cliqnro
          'Call generarDatos(empresa, fechadesde, fechahasta, cliqnro, ternro, auxcon1, auxcon2, auxcon3, nro_acu_neto, nro_acu_imponible, profecpago, incoperben)
          Call generarDatos(empresa, fechadesde, fechahasta, cliqnro, Ternro, nro_acu_neto, nro_acu_imponible, profecpago, incoperben)
          
          'Actualizo el estado del proceso
          TiempoAcumulado = GetTickCount

          StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix((actualReg * 100) / cantRegistros) & _
                   ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'" & _
                   ", bprcempleados ='" & CStr((cantRegistros - actualReg)) & "' WHERE bpronro = " & NroProceso
          'objConn.Execute StrSql, , adExecuteNoRecords
          objconnProgreso.Execute StrSql, , adExecuteNoRecords
          
          actualReg = actualReg + 1

          rsEmpl.MoveNext
       Loop
    
    Else
        Exit Sub
    End If
   
    If Not HuboErrores Then
        'Actualizo el estado del proceso
        StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprcempleados ='0', bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprcempleados ='0', bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Flog.writeline "Fin :" & Now
    Flog.Close
    objConn.Close
    Exit Sub
    
CE:
    HuboErrores = True
    Flog.writeline " Error: " & Err.Description & Now

End Sub

Sub comprobarCodigosOBE(ByVal cod1, ByVal cod2, ByVal cod3, ByVal cod4)

Dim problemas

    problemas = 0

    If (cod1 = 0) Then
        problemas = problemas + 1
        Flog.writeline "No se encontro la columna 50 en el CONFREP referida a Denominación del Ordenante."
    End If
    If (cod2 = 0) Then
        problemas = problemas + 1
        Flog.writeline "No se encontro la columna 51 en el CONFREP referida a Acrecentamiento."
    End If
    If (cod3 = 0) Then
        problemas = problemas + 1
        Flog.writeline "No se encontro la columna 52 en el CONFREP referida a CUIT del país del retenido."
    End If
    If (cod4 = 0) Then
        problemas = problemas + 1
        Flog.writeline "No se encontro la columna 53 en el CONFREP referida al Codigo del tipo de codigo de CUIT."
    End If
    If (problemas = 0) Then
        Flog.writeline "Configuracion para Operaciones con Beneficiarios del Exterior cargada correctamente."
    End If
    
End Sub

Sub Desgloce(empresa, rs_auxcon, Ternro, codobe1, codobe2, codobe3, fec_hasta, fec_desde, incoperben, parte_ent_neto, parte_dec_neto, nro_acu_imponible, impo_ent, impo_dec, signo, cliqnro, fec_ret, ConcTipo)
Dim sumaMonto
Dim ret_lin
Dim aux_lin
Dim parte_ent
Dim parte_dec
Dim parte_ent_impo
Dim parte_dec_impo
Dim cod_liq
Dim cod_impuesto
Dim cod_regimen
Dim neto_ent
Dim neto_dec
Dim cod_condicion
Dim gan_ent
Dim gan_dec
Dim fec_bol
Dim porcexl
Dim tipo_doc
Dim Monto
Dim pos
Dim apenom
Dim empleg
Dim dom_fiscal
Dim dom_localidad
Dim cod_provincia
Dim dom_cp
Dim extranjero
Dim deno_orden
Dim acrecent
Dim codtcod
Dim cuit_pais
Dim cuit_orden
Dim Cuil
Dim fec_ret_ano
Dim fec_retencion
Dim fec_ret_CAA

Dim rsConsultEmp As New ADODB.Recordset
Dim rsConsultDir As New ADODB.Recordset
Dim rsconsult As New ADODB.Recordset
Dim rsconsultconfrep As New ADODB.Recordset
Dim rsdetliq As New ADODB.Recordset

If signo = 7 Then
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = 14 And confnrocol = 1"
Else
    If signo = 8 Then
        StrSql = "SELECT * FROM confrep "
        StrSql = StrSql & " WHERE repnro = 14 And confnrocol = 2"
    End If
End If
OpenRecordset StrSql, rsconsultconfrep
Do Until rsconsultconfrep.EOF
    StrSql = "SELECT * FROM detliq "
    StrSql = StrSql & "INNER JOIN concepto on concepto.concnro=detliq.concnro "
    StrSql = StrSql & "WHERE detliq.cliqnro = " & cliqnro & " AND concepto.conccod = " & rsconsultconfrep!confval
    OpenRecordset StrSql, rsdetliq
    sumaMonto = 0
    fec_retencion = fec_ret
    fec_ret_CAA = fec_ret
    If Not rsdetliq.EOF Then
        If (signo = 7 And rsdetliq!dlimonto < 0) Or (signo = 8 And rsdetliq!dlimonto > 0) Then
        sumaMonto = sumaMonto + rsdetliq!dlimonto
        If sumaMonto <> 0 Then
                'Armar el valor absoluto del movimiento
                ret_lin = (-1) * sumaMonto
                If ret_lin < 0 Then
                   aux_lin = (-1) * ret_lin
                Else
                   aux_lin = ret_lin
                End If
                'FGZ - 10/03/2006 - Sacaba mal la parte decimal cuando tenia un solo decimal
                parte_ent = Fix(aux_lin)
                pos = InStr(1, aux_lin, ".")
                If pos > 1 Then
                    parte_dec = Mid(aux_lin & "0", pos + 1, 2)
                Else
                    parte_dec = 0
                End If

                ' Si es un recibo de sueldo devolucion (signo = 8 o devolucion de ganancias) entonces
                '    asigna a la base imponible el valor de la devolucion.
                ' Si es recibo de sueldo (signo = 7 o retencion de ganancias) busca la base imponible del recibo de sueldo ,columan 5
                
                If signo = "8" Then
                   parte_ent_impo = Fix(aux_lin)
                   'parte_dec_impo = Fix((aux_lin - parte_ent_impo) * 100)
                    pos = InStr(1, aux_lin, ".")
                    If pos > 1 Then
                        parte_dec_impo = Mid(aux_lin & "0", pos + 1, 2)
                    Else
                        parte_dec_impo = 0
                    End If
                End If
                
                cod_liq = cliqnro
                'cod_impuesto = "217"
                cod_impuesto = "787"
                cod_regimen = "160"
                neto_ent = parte_ent_neto
                neto_dec = parte_dec_neto
                If CInt(nro_acu_imponible) <> -1 Then
                 impo_ent = parte_ent_impo
                 impo_dec = parte_dec_impo
                End If
                cod_condicion = "01"
                gan_ent = parte_ent
                gan_dec = parte_dec
                fec_bol = "        "
                porcexl = "000,00"
                tipo_doc = "86"
                Monto = "0"
                
                '------------------------------------------------------------------
                'Busco los datos del empleado
                '------------------------------------------------------------------
                StrSql = " SELECT empleg,terape,terape2,ternom,ternom2,empfecalta,empremu "
                StrSql = StrSql & " FROM empleado "
                StrSql = StrSql & " WHERE ternro= " & Ternro
                
                OpenRecordset StrSql, rsConsultEmp
                
                If Not rsConsultEmp.EOF Then
                   apenom = rsConsultEmp!terape & ", " & rsConsultEmp!ternom
                   empleg = rsConsultEmp!empleg
                Else
                   Flog.writeline "Error al obtener los datos del empleado"
                   'GoTo MError
                End If
                
                '------------------------------------------------------------------
                'Busco el valor de la direccion y localidad
                '------------------------------------------------------------------
                StrSql = " SELECT detdom.calle,detdom.nro,localidad.locdesc, detdom.codigopostal, provincia.provcodext "
                StrSql = StrSql & " FROM cabdom "
                StrSql = StrSql & " INNER JOIN detdom ON detdom.domnro = cabdom.domnro AND cabdom.ternro = " & Ternro & " AND (cabdom.tipnro=1 OR cabdom.tipnro=14) AND cabdom.domdefault = -1 "
                StrSql = StrSql & " LEFT JOIN localidad ON detdom.locnro = localidad.locnro"
                StrSql = StrSql & " LEFT JOIN provincia ON detdom.provnro = provincia.provnro"
                OpenRecordset StrSql, rsConsultDir
                
                If Not rsConsultDir.EOF Then
                   dom_fiscal = rsConsultDir!calle & " " & rsConsultDir!nro
                   dom_localidad = rsConsultDir!locdesc
                   cod_provincia = rsConsultDir!provcodext
                   dom_cp = rsConsultDir!codigopostal
                Else
                   Flog.writeline "Error al obtener los datos de la direccion"
                   'GoTo MError
                End If
    
                '------------------------------------------------------------------
                'Busco si el empleado es extranjero o no
                '------------------------------------------------------------------
                extranjero = 0
                StrSql = " SELECT empresext FROM empleado WHERE ternro = " & Ternro
                OpenRecordset StrSql, rsConsultDir
                extranjero = rsConsultDir!empresext
                If (incoperben = -1) And (extranjero = -1) Then
                    
                    '------------------------------------------------------------------
                    ' Busco Denominación del Ordenante
                    '------------------------------------------------------------------
                    If (codobe1 <> 0) Then
                        StrSql = " SELECT estructura.estrdabr detalle FROM his_estructura INNER JOIN estructura ON "
                        StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro WHERE "
                        StrSql = StrSql & " his_estructura.tenro = " & codobe1
                        StrSql = StrSql & " AND his_estructura.ternro = " & Ternro
                        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fec_hasta)
                        StrSql = StrSql & " AND ( his_estructura.htethasta >= " & ConvFecha(fec_hasta)
                        StrSql = StrSql & " OR his_estructura.htethasta IS NULL ) "
                        OpenRecordset StrSql, rsConsultDir
                    
                        If Not rsConsultDir.EOF Then
                            deno_orden = rsConsultDir!detalle
                        Else
                           Flog.writeline Espacios(Tabulador * 2) & "No se encontraton datos de Denominacion del ordenante"
                           deno_orden = ""
                        End If
                    Else
                        deno_orden = ""
                    End If
                    
                    '--------------------------------------------------------------------------
                    ' Busco si el empleado tiene asociado el tipo de estructura Acrecentamiento
                    '--------------------------------------------------------------------------
                    acrecent = 0
                    If (codobe2 <> 0) Then
                        StrSql = " SELECT estructura.estrdabr detalle FROM his_estructura INNER JOIN estructura ON "
                        StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro WHERE "
                        StrSql = StrSql & " his_estructura.tenro = " & codobe2
                        StrSql = StrSql & " AND his_estructura.ternro = " & Ternro
                        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fec_hasta)
                        StrSql = StrSql & " AND ( his_estructura.htethasta >= " & ConvFecha(fec_hasta)
                        StrSql = StrSql & " OR his_estructura.htethasta IS NULL ) "
                                 
                        OpenRecordset StrSql, rsConsultDir
                    
                        If Not rsConsultDir.EOF Then
                            acrecent = 1
                        Else
                           Flog.writeline Espacios(Tabulador * 2) & "No se encontraron datos de Acrecentamiento"
                           acrecent = 0
                        End If
                    Else
                        acrecent = 0
                    End If
                    
                    '--------------------------------------------------------------------------
                    ' Busco CUIT del pais del retenido
                    '--------------------------------------------------------------------------
                    If (codobe3 <> 0) And (codtcod <> 0) Then
                        StrSql = " SELECT nrocod FROM his_estructura INNER JOIN estructura ON "
                        StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro "
                        StrSql = StrSql & " INNER JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro WHERE "
                        StrSql = StrSql & " his_estructura.tenro = " & codobe3
                        StrSql = StrSql & " AND his_estructura.ternro = " & Ternro
                        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fec_hasta)
                        StrSql = StrSql & " AND ( his_estructura.htethasta >= " & ConvFecha(fec_hasta)
                        StrSql = StrSql & " OR his_estructura.htethasta IS NULL ) "
                        StrSql = StrSql & " AND estr_cod.tcodnro = " & codtcod
                                 
                        OpenRecordset StrSql, rsConsultDir
                    
                        If Not rsConsultDir.EOF Then
                            cuit_pais = rsConsultDir!nrocod
                        Else
                           Flog.writeline Espacios(Tabulador * 2) & "No se encontron datos sobre el CUIT del país del retenido"
                           cuit_pais = ""
                        End If
                    Else
                        cuit_pais = ""
                    End If
                    
                    '--------------------------------------------------------------------------
                    ' Busco CUIT del Ordenante
                    '--------------------------------------------------------------------------
                    If (codobe1 <> 0) And (codtcod <> 0) Then
                        StrSql = " SELECT nrocod FROM his_estructura INNER JOIN estructura ON "
                        StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro "
                        StrSql = StrSql & " INNER JOIN estr_cod ON estr_cod.estrnro = his_estructura.estrnro WHERE "
                        StrSql = StrSql & " his_estructura.tenro = " & codobe1
                        StrSql = StrSql & " AND his_estructura.ternro = " & Ternro
                        StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(fec_hasta)
                        StrSql = StrSql & " AND ( his_estructura.htethasta >= " & ConvFecha(fec_hasta)
                        StrSql = StrSql & " OR his_estructura.htethasta IS NULL ) "
                        StrSql = StrSql & " AND estr_cod.tcodnro = " & codtcod
                                 
                        OpenRecordset StrSql, rsConsultDir
                    
                        If Not rsConsultDir.EOF Then
                            cuit_orden = rsConsultDir!nrocod
                        Else
                           Flog.writeline Espacios(Tabulador * 2) & "No se encontron datos sobre el CUIT del ordenante"
                           cuit_orden = ""
                        End If
                    Else
                        cuit_orden = ""
                    End If
                    
                Else 'si no ex extranjero o no esta acitvado el incluye extranjeros
                
                    cuit_orden = ""
                    cuit_pais = ""
                    acrecent = 0
                    deno_orden = ""
                    Flog.writeline Espacios(Tabulador * 2) & "El empleado no es extranjero"
                
                End If
                
                '------------------------------------------------------------------
                'Armo la SQL para guardar los datos
                '------------------------------------------------------------------
                
                If CInt(nro_acu_imponible) = -1 Then    ' voy a comparar si el signo es 8 y la ganancia imponible es <> 0 voy a poner gan_ent y gan_dec
                   
                   If (impo_ent <> 0 Or impo_dec <> 0) And signo = 8 Then
                      
                     impo_ent = gan_ent
                     impo_dec = gan_dec
                   Flog.writeline "el signo es 8, pongo en ganancias imponibles las ganancias"
                   End If
                End If
                
                '------------------------------------------------------------------
                'Busco el valor del cuil
                '------------------------------------------------------------------
                StrSql = " SELECT cuil.nrodoc "
                StrSql = StrSql & " FROM tercero LEFT JOIN ter_doc cuil ON (tercero.ternro=cuil.ternro and cuil.tidnro=10) "
                StrSql = StrSql & " WHERE tercero.ternro= " & Ternro
                       
                OpenRecordset StrSql, rsconsult
                
                If Not rsconsult.EOF Then
                   Cuil = rsconsult!NroDoc
                Else
                   Flog.writeline "Error al obtener los datos del cuil"
                   GoTo MError
                End If
                
                'Le quito al cuil los -
                If InStr(Cuil, "-") Then
                   Cuil = Replace(Cuil, "-", "")
                End If
                
                StrSql = "SELECT * FROM confrep "
                StrSql = StrSql & "WHERE conftipo= 'CAA' AND repnro = 14 "
                StrSql = StrSql & " AND confval IN(" & rsconsultconfrep!confval & ")"
                OpenRecordset StrSql, rsconsult
                If Not rsconsult.EOF Then
                    fec_ret_ano = Year(fec_retencion)
                    fec_ret_CAA = "31/12/" & (fec_ret_ano - 1)
                End If
                
                StrSql = " INSERT INTO sicore "
                StrSql = StrSql & "(empresa, fec_desde ,fec_hasta    ,empleg       ,fec_ret, inc_oper_ben, "
                StrSql = StrSql & " cod_liq   , cod_impuesto ,cod_regimen  ,neto_ent, "
                StrSql = StrSql & " neto_dec  , impo_ent     ,impo_dec     ,cod_condicion,gan_ent, "
                StrSql = StrSql & " gan_dec   , fec_bol      ,porcexl      ,tipo_doc, "
                StrSql = StrSql & " cuil      , monto        ,signo        ,apenom, "
                StrSql = StrSql & " dom_fiscal, dom_localidad,cod_provincia,dom_cp, "
                StrSql = StrSql & " deno_orden, acrecent, cuit_pais, fec_ret_CAA, cuit_orden )"
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "(" & empresa
                StrSql = StrSql & "," & ConvFecha(fec_desde)
                StrSql = StrSql & "," & ConvFecha(fec_hasta)
                StrSql = StrSql & "," & empleg
                StrSql = StrSql & ",'" & Mid(fec_ret, 1, 10) & "'"
                StrSql = StrSql & "," & incoperben & ""
                StrSql = StrSql & ",'" & Mid(cod_liq, 1, 12) & "'"
                StrSql = StrSql & ",'" & Mid(cod_impuesto, 1, 3) & "'"
                StrSql = StrSql & ",'" & Mid(cod_regimen, 1, 3) & "'"
                StrSql = StrSql & ",'" & Mid(neto_ent, 1, 13) & "'"
                StrSql = StrSql & ",'" & Mid(neto_dec, 1, 2) & "'"
                StrSql = StrSql & ",'" & Mid(impo_ent, 1, 13) & "'"
                StrSql = StrSql & ",'" & Mid(impo_dec, 1, 2) & "'"
                StrSql = StrSql & ",'" & Mid(cod_condicion, 1, 2) & "'"
                StrSql = StrSql & ",'" & Mid(gan_ent, 1, 13) & "'"
                StrSql = StrSql & ",'" & Mid(gan_dec, 1, 2) & "'"
                StrSql = StrSql & ",'" & Mid(fec_bol, 1, 10) & "'"
                StrSql = StrSql & ",'" & Mid(porcexl, 1, 10) & "'"
                StrSql = StrSql & ",'" & Mid(tipo_doc, 1, 2) & "'"
                StrSql = StrSql & ",'" & Mid(Cuil, 1, 20) & "'"
                StrSql = StrSql & ",'" & Mid(Monto, 1, 20) & "'"
                StrSql = StrSql & ",'" & Mid(signo, 1, 1) & "'"
                StrSql = StrSql & ",'" & Mid(apenom, 1, 20) & "'"
                StrSql = StrSql & ",'" & Mid(dom_fiscal, 1, 20) & "'"
                StrSql = StrSql & ",'" & Mid(dom_localidad, 1, 20) & "'"
                StrSql = StrSql & ",'" & Mid(cod_provincia, 1, 2) & "'"
                StrSql = StrSql & ",'" & Mid(dom_cp, 1, 8) & "'"
                StrSql = StrSql & ",'" & Mid(deno_orden, 1, 30) & "'"
                StrSql = StrSql & ",'" & acrecent & "'"
                StrSql = StrSql & ",'" & Mid(cuit_pais, 1, 11) & "'"
                StrSql = StrSql & ",'" & Mid(fec_ret_CAA, 1, 10) & "'"
                StrSql = StrSql & ",'" & Mid(cuit_orden, 1, 11) & "')"
                
                
                '------------------------------------------------------------------
                'Guardo los datos en la BD
                '------------------------------------------------------------------
                objConn.Execute StrSql, , adExecuteNoRecords
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 3) & "Inserté los importes -----------------> "
                Flog.writeline Espacios(Tabulador * 3) & "        Imponible " & impo_ent & "." & impo_dec
                Flog.writeline Espacios(Tabulador * 3) & "        Ganancia  " & gan_ent & "." & gan_dec
                
                'Exit Sub
        Else
            Flog.writeline "La suma del Monto es 0. SQL. " & StrSql
            Flog.writeline "Siguiente legajo"
            Flog.writeline
        End If
    End If
    End If
    rsconsultconfrep.MoveNext
Loop
    Exit Sub
MError:
    Flog.writeline "Error en empleado: " & empleg & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

'--------------------------------------------------------------------
' Se encarga de generar los datos
'--------------------------------------------------------------------
Sub generarDatos(empresa, Desde, Hasta, cliqnro, Ternro, nro_acu_neto, nro_acu_imponible, profecpago, incoperben)

Dim StrSql As String
Dim rsconsult As New ADODB.Recordset
Dim rsConsultEmp As New ADODB.Recordset
Dim rsConsultDir As New ADODB.Recordset
Dim rs_auxcon As New ADODB.Recordset
Dim rs_Acu_liq As New ADODB.Recordset

'Variables donde se guardan los datos del INSERT final

' NAM - 01/02/2011

Dim deno_orden As String
Dim acrecent
Dim cuit_pais As String
Dim cuit_orden As String
Dim extranjero

Dim sumaMonto

Dim fec_desde
Dim fec_hasta
Dim empleg
Dim fec_ret
Dim cod_liq
Dim cod_impuesto
Dim cod_regimen
Dim neto_ent
Dim neto_dec
Dim impo_ent
Dim impo_dec
Dim cod_condicion
Dim gan_ent
Dim gan_dec
Dim fec_bol
Dim porcexl
Dim tipo_doc
Dim Cuil
Dim Monto
Dim signo
Dim apenom
Dim dom_fiscal
Dim dom_localidad
Dim cod_provincia
Dim dom_cp

Dim parte_ent_neto
Dim parte_dec_neto
Dim ret_lin
Dim aux_lin
Dim parte_ent As Double
Dim parte_dec As Double
Dim parte_ent_impo
Dim parte_dec_impo
'FGZ - 10/03/2006
Dim pos
Dim ganimpon 'mdf

Dim I As Integer
Dim PrimeraVez As Boolean
Dim StrSql2 As String 'mdf

Dim ConcTipo
fec_desde = Desde
fec_hasta = Hasta

On Error GoTo MError

'------------------------------------------------------------------
'Fecha
'------------------------------------------------------------------
fec_ret = profecpago

'------------------------------------------------------------------
' Busco el acumulador NETO de liquidacion definidos en el confrep
'------------------------------------------------------------------
Flog.writeline
Flog.writeline Espacios(Tabulador * 1) & "Busco el acumulador NETO de liquidacion definidos en el confrep"
parte_ent_neto = 0
parte_dec_neto = 0
sumaMonto = 0
    
StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & cliqnro & _
         " AND acunro =" & nro_acu_neto

OpenRecordset StrSql, rs_Acu_liq

Do Until rs_Acu_liq.EOF
    Flog.writeline Espacios(Tabulador * 2) & "Monto: " & rs_Acu_liq!almonto
    sumaMonto = sumaMonto + rs_Acu_liq!almonto
    rs_Acu_liq.MoveNext
Loop

parte_ent_neto = Fix(sumaMonto)
parte_dec_neto = Fix((sumaMonto - parte_ent_neto) * 100)

'------------------------------------------------------------------
' Busco el acumulador GAN IMPONIBLE de liquidacion definidos en el confrep
'------------------------------------------------------------------
'--------------------mdf inicio-------------------------
If CInt(nro_acu_imponible) <> -1 Then 'mdf --> cuando ponen -1 se trae el valor de traza_gan para el proceso de liq (por si cambia a lo de antes)
Flog.writeline
Flog.writeline Espacios(Tabulador * 1) & "Busco el acumulador GAN IMPONIBLE de liquidacion definidos en el confrep"

parte_ent_impo = 0
parte_dec_impo = 0
sumaMonto = 0

        StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & cliqnro & _
                 " AND acunro =" & nro_acu_imponible
        
        OpenRecordset StrSql, rs_Acu_liq
        
        Do Until rs_Acu_liq.EOF
            Flog.writeline Espacios(Tabulador * 2) & "Monto: " & rs_Acu_liq!almonto
            sumaMonto = sumaMonto + rs_Acu_liq!almonto
            rs_Acu_liq.MoveNext
        Loop
        
        parte_ent_impo = Fix(sumaMonto)
        parte_dec_impo = Fix((sumaMonto - parte_ent_impo) * 100)
 End If 'mdf
 
 If CInt(nro_acu_imponible) = -1 Then
        Flog.writeline "-----------inicio nuevo calculo ganancia imponibles------------------------"
        ganimpon = "0"
        StrSql2 = StrSql2 & "Select max(ganimpo)ganimpo from traza_gan "
        StrSql2 = StrSql2 & "INNER JOIN proceso ON proceso.pronro = traza_gan.pronro "
        StrSql2 = StrSql2 & "INNER JOIN tipoproc ON proceso.tprocnro = tipoproc.tprocnro "
        StrSql2 = StrSql2 & "WHERE traza_gan.ternro = " & Ternro
        StrSql2 = StrSql2 & " and (traza_gan.fecha_pago <= '" & fec_hasta & "' AND fecha_pago >= '" & fec_desde & "')"
        StrSql2 = StrSql2 & " group by fecha_pago, traza_gan.pronro "
        StrSql2 = StrSql2 & " ORDER BY fecha_pago DESC, traza_gan.pronro DESC"
        OpenRecordset StrSql2, rsConsultDir
        If Not rsConsultDir.EOF Then
             ganimpon = rsConsultDir!ganimpo
        End If
        impo_ent = Fix(ganimpon)
        impo_dec = Fix((ganimpon - impo_ent) * 100)
        Flog.writeline "consulta trazagan: " & StrSql2
        Flog.writeline "parte entera: " & impo_ent
        Flog.writeline "parte decimal: " & impo_dec
        Flog.writeline "-----------fin nuevo calculo ganancia imponible------------------------"
 End If
'--------------------mdf fin -------------------------
       
        
'------------------------------------------------------------------
' Busco los detalles de liquidacion definidos en el confrep
'------------------------------------------------------------------
Flog.writeline
Flog.writeline Espacios(Tabulador * 1) & "Busco los detalles de liquidacion definidos en el confrep"

PrimeraVez = True
'FGZ - 16/11/2004
'------------------------------ Calcula si corresponde al año anterior ----------------------------------------------
'ConcTipo = "0"
'StrSql = " SELECT * "
'StrSql = StrSql & " FROM detliq "
'StrSql = StrSql & " WHERE cliqnro= " & cliqnro
'StrSql = StrSql & "   AND ( "
'For I = 0 To Indice1
'    If PrimeraVez Then
'        StrSql = StrSql & " concnro = " & Concepto1(I).ConcNro
'        PrimeraVez = False
'    Else
'        StrSql = StrSql & " OR concnro = " & Concepto1(I).ConcNro
'    End If
'    ConcTipo = ConcTipo & "," & Concepto1(I).ConcCod
'Next I
'StrSql = StrSql & "   ) "
'OpenRecordset StrSql, rs_auxcon
'If Not rs_auxcon.EOF Then
    signo = 7
    Call Desgloce(empresa, rs_auxcon, Ternro, codobe1, codobe2, codobe3, fec_hasta, fec_desde, incoperben, parte_ent_neto, parte_dec_neto, nro_acu_imponible, impo_ent, impo_dec, signo, cliqnro, fec_ret, ConcTipo)
'End If
'ConcTipo = "0"
'StrSql = " SELECT * "
'StrSql = StrSql & " FROM detliq "
'StrSql = StrSql & " WHERE cliqnro= " & cliqnro
'StrSql = StrSql & "   AND ( "
'PrimeraVez = True
'For I = 0 To Indice2
'    If PrimeraVez Then
'        StrSql = StrSql & " concnro = " & Concepto2(I).ConcNro
'        PrimeraVez = False
'    Else
'        StrSql = StrSql & " OR concnro = " & Concepto2(I).ConcNro
'    End If
'    ConcTipo = ConcTipo & "," & Concepto2(I).ConcCod
'Next I
'StrSql = StrSql & "   ) "
'OpenRecordset StrSql, rs_auxcon
'If Not rs_auxcon.EOF Then
    signo = 8
    Call Desgloce(empresa, rs_auxcon, Ternro, codobe1, codobe2, codobe3, fec_hasta, fec_desde, incoperben, parte_ent_neto, parte_dec_neto, nro_acu_imponible, impo_ent, impo_dec, signo, cliqnro, fec_ret, ConcTipo)
'End If
'ConcTipo = "0"
'StrSql = " SELECT * "
'StrSql = StrSql & " FROM detliq "
'StrSql = StrSql & " WHERE cliqnro= " & cliqnro
'StrSql = StrSql & "   AND ( "
'PrimeraVez = True
'For I = 0 To Indice3
'    If PrimeraVez Then
'        StrSql = StrSql & " concnro = " & Concepto3(I).ConcNro
'        PrimeraVez = False
'    Else
'        StrSql = StrSql & " OR concnro = " & Concepto3(I).ConcNro
'    End If
'    ConcTipo = ConcTipo & "," & Concepto3(I).ConcCod
'Next I
'StrSql = StrSql & "   ) "
'OpenRecordset StrSql, rs_auxcon
'If Not rs_auxcon.EOF Then
'    Call Desgloce(empresa, rs_auxcon, Ternro, codobe1, codobe2, codobe3, fec_hasta, fec_desde, incoperben, parte_ent_neto, parte_dec_neto, nro_acu_imponible, impo_ent, impo_dec, signo, cliqnro, fec_ret, ConcTipo)
'End If
Exit Sub

MError:
    Flog.writeline "Error en empleado: " & empleg & " Error: " & Err.Description
    HuboErrores = True
    EmpErrores = True
    Exit Sub
End Sub

