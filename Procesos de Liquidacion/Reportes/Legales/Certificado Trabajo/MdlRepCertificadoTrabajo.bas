Attribute VB_Name = "MdlRepCertificadoTrabajo"
Option Explicit

'Version: 1.01
' Correccion en el rango del calculo de los anios
'Const Version = 1.01
'Const FechaVersion = "17/07/2006" ' FAF - 17/07/2006 - Para el Certificado de empleo, el en cuil mostraba en Nro Documento.
        
'Global Const Version = "1.02" ' Cesar Stankunas
'Global Const FechaVersion = "06/08/2009"
'Global Const UltimaModificacion = ""    'Encriptacion de string connection

'Global Const Version = "1.03" ' Matias Dallegro
'Global Const FechaVersion = "22/09/2011"
'Global Const UltimaModificacion = "" ' Si empleado esta inactivo arrojaba error en la fase del empleado
 'Global Const Version = "1.04" ' Matias Dallegro
 'Global Const FechaVersion = "03/10/2011"
 'Global Const UltimaModificacion = "" 'No funciona con mas de un proceso
 
 'Global Const Version = "1.05" ' Manterola Maria Magdalena
 'Global Const FechaVersion = "06/10/2011"
 'Global Const UltimaModificacion = "Se agrego el Reporte de Certificado Salarial"
 
' Global Const Version = "1.06" ' Manterola Maria Magdalena
' Global Const FechaVersion = "31/10/2011"
' Global Const UltimaModificacion = "Se agrego la condicion de que si el empleado tiene ternom2 = '' o terape2 = '' elimine el espacio"


'Global Const Version = "1.07" ' Gonzalez Nicolás
'Global Const FechaVersion = "02/11/2012"
'Global Const UltimaModificacion = "Se modifica DatosCertifSalarial "

'Global Const Version = "1.08" ' Fernando Favre
'Global Const FechaVersion = "11/12/2012"
'Global Const UltimaModificacion = "Se modifica Admcer02_Empleo, agregando detdom. a los campos zonanro. Se corrigio avance de proceso. Se agrego manejador de errores. CAS - 17520 - OSDOP - Error certificado de Empleo "

'Global Const Version = "1.09" ' FGZ
'Global Const FechaVersion = "03/01/2013"
'Global Const UltimaModificacion = " "
''                                               CAS-17520 - OSDOP - Error certificado de Empleo
''                                                   Se agregó manejador de error en sub LevantarParamteros

'Global Const Version = "1.10" ' FGZ
'Global Const FechaVersion = "23/05/2013"
'Global Const UltimaModificacion = " "
'                                           CAS-18674 - Sykes - CUSTOM REPORTE DE CERTIFICADO POR EMPLEADO
'                                           Cuando busca los embargos no tenia en cuenta el estado de los mismos. Solo Debe buscar Activos (embest = 'A')

'Global Const Version = "1.11" ' Carmen Quintero
'Global Const FechaVersion = "02/08/2013"
'Global Const UltimaModificacion = " "
'                                           CAS-20675 - SYKES CR - Error en certificado salarial en multiples fases
'                                           Se modificó la funcion DatosCertifSalarial por agregarse validacion para el caso
'                                           cuando un empleado tiene multiples fases y solo considere la fase activa.

'Global Const Version = "1.12" ' Borrelli Facundo
'Global Const FechaVersion = "17/02/2014"
'Global Const UltimaModificacion = " "
'                                           CAS-23926 - LA CAPITAL - ERROR EN CERTIFICADO DE EMPLEO
'                                           Se corrige la consulta que busca los procesos a evaluar,
'                                           se cambia el * por los campos necesarios en la consulta.

'Global Const Version = "1.13" ' Carmen Quintero
'Global Const FechaVersion = "23/06/2014"
'Global Const UltimaModificacion = " "
'                                           CAS-26046 - ZARCAM - Error en Certificado de Empleo
'                                           Se corrige la consulta que busca los procesos a evaluar,
'                                           se cambia el * por los campos necesarios en la consulta.
'                                           para el caso del Certificado Salarial

'Global Const Version = "1.14" ' Miriam Ruiz
'Global Const FechaVersion = "17/12/2015"
'Global Const UltimaModificacion = " "
'                                           CAS-34511 - ASM - Reporte Certificado de Prestacion de Servicios
'                                           Se corrige el cuil en el certificado 38,

Global Const Version = "1.15" ' Miriam Ruiz
Global Const FechaVersion = "17/12/2015"
Global Const UltimaModificacion = " "
'                                           CAS-34511 - ASM - Reporte Certificado de Prestacion de Servicios - entrega 2
'                                           Se dejan los guiones en el cuil

Global IdUser As String
Global Fecha As Date
Global Hora As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reportes.
' Autor      : FGZ
' Fecha      : 17/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
Dim ArrParametros

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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    Nombre_Arch = PathFLog & "Certificados_Empleo" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Fecha   = " & FechaVersion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE (btprcnro = 37 OR btprcnro = 38 OR btprcnro = 39 OR btprcnro = 314) AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    Flog.Close
    objConn.Close
    objconnProgreso.Close
    
    Set objconnProgreso = Nothing
    Set objConn = Nothing
    
End Sub


'Public Sub Admcer02_Prestacion(ByVal TipoReporte As Integer, ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Todos_Empleados As Boolean)
'' --------------------------------------------------------------------------------------------
'' Descripcion: Procedimiento de generacion del Reporte de Certificado de Trabajo
'' Autor      : FGZ
'' Fecha      : 05/03/2004
'' Ult. Mod   :
'' Fecha      :
'' --------------------------------------------------------------------------------------------
'Dim fechadesde As Date
'Dim fechahasta As Date
'Dim fec_ini As Date
'Dim fec_fin As Date
'
'Dim selector  As Integer
'Dim selector2 As Integer
'Dim selector3 As Integer
'Dim X As Boolean
'Dim Par As Integer
'Dim Aux_CUIL As String
'
'
'Dim rs_Reporte As New ADODB.Recordset
'Dim rs_Periodo As New ADODB.Recordset
'Dim rs_Procesos As New ADODB.Recordset
'Dim rs_Confrep As New ADODB.Recordset
'Dim rs_Detliq As New ADODB.Recordset
'Dim rs_aculiq As New ADODB.Recordset
'Dim rs_rep67 As New ADODB.Recordset
'Dim rs_Empleados As New ADODB.Recordset
'Dim rs_Fases As New ADODB.Recordset
'Dim rs_Cuil As New ADODB.Recordset
'
''Configuracion del Reporte
'Select Case TipoReporte
'Case 37: 'Certificado de Empleo
'    StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 37"
'Case 38: 'Certificado de Prestacion de Servicio
'    StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 38"
'Case 39: 'Certificado de DNRP
'    StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 39"
'End Select
'OpenRecordset StrSql, rs_Confrep
'If rs_Confrep.EOF Then
'    Flog.writeline "No se encontró la configuración del Reporte"
'    Exit Sub
'End If
'
''cargo el periodo
'StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
'OpenRecordset StrSql, rs_Periodo
'If rs_Periodo.EOF Then
'    Flog.writeline "No se encontró el Periodo"
'    Exit Sub
'End If
'If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
'
'' Comienzo la transaccion
'MyBeginTrans
'
''Depuracion del Temporario
'StrSql = "DELETE FROM rep67 "
'StrSql = StrSql & " WHERE pliqnro = " & Nroliq
'If Not Todos_Pro Then
'    StrSql = StrSql & " AND pronro = " & NroProc
'Else
'    StrSql = StrSql & " AND pronro = 0"
'    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
'End If
'StrSql = StrSql & " AND empresa = " & Empresa
'objConn.Execute StrSql, , adExecuteNoRecords
'
''Busco los procesos a evaluar
'StrSql = "SELECT * FROM  empleado "
'If Not Todos_Empleados Then
'    StrSql = StrSql & " INNER JOIN batch_empleado ON batch_empleado.ternro = empleado.ternro "
'End If
'StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
'StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
'StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
'StrSql = StrSql & " WHERE periodo.pliqnro =" & Nroliq
''StrSql = StrSql & " AND periodo.empnro =" & Empresa
'If Not Todos_Pro Then
'    StrSql = StrSql & " AND proceso.pronro =" & NroProc
'Else
'    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
'End If
'StrSql = StrSql & " ORDER BY empleado.ternro"
'OpenRecordset StrSql, rs_Procesos
'
'If Not rs_Procesos.EOF Then
'    fechadesde = rs_Procesos!pliqdesde
'    fechahasta = rs_Procesos!pliqhasta
'End If
'
'Do While Not rs_Procesos.EOF
'    StrSql = "SELECT * FROM batch_empleado " & _
'             " INNER JOIN cabliq ON cabliq.empleado = batch_empleado.ternro " & _
'             " WHERE cabliq.pronro =" & rs_Procesos!pronro
'    OpenRecordset StrSql, rs_Empleados
'
'    'seteo de las variables de progreso
'    Progreso = 0
'    CConceptosAProc = rs_Procesos.RecordCount
'    CEmpleadosAProc = rs_Empleados.RecordCount
'    If CEmpleadosAProc = 0 Then
'       CEmpleadosAProc = 1
'    End If
'    IncPorc = ((100 / CEmpleadosAProc) * (100 / CConceptosAProc)) / 100
'
'    Do While Not rs_Empleados.EOF
'
'        StrSql = "SELECT * FROM fases WHERE empleado = " & rs_Empleados!ternro
'        OpenRecordset StrSql, rs_Fases
'
'        If Not rs_Fases.EOF Then
'            rs_Fases.MoveFirst
'            fec_ini = rs_Fases!altfec
'
'            rs_Fases.MoveLast
'            fec_fin = rs_Fases!bajfec
'        End If
'
'        If TipoReporte <> 1 And (IsNull(fec_ini) Or (IsNull(fec_fin) Or Not rs_Fases!estado) Or (fec_ini > fec_fin)) Then
'            Flog.writeline "Error en las fases del empleado"
'            Exit Sub
'        End If
'
'        'busco la categoria del empleado
'        StrSql = " SELECT * FROM his_estructura " & _
'                 " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
'                 " WHERE his_estructura.ternro = " & rs_Empleados!ternro & " AND " & _
'                 " his_estructura.tenro = 3 AND " & _
'                 " (his_estructura.htetdesde <= " & ConvFecha(fechahasta) & ") AND " & _
'                 " ((" & ConvFecha(fechahasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))" & _
'                 " ORDER BY his_estructura.htetdesde"
'        If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
'        OpenRecordset StrSql, rs_Estructura
'        If Not rs_Estructura.EOF Then
'            Aux_Catdesc = rs_Estructura!estrdabr
'        Else
'            Flog.writeline "No se encontro la Categoria del empleado"
'            Aux_Catdesc = " "
'        End If
'
'
'        'Si no existe el rep67
'        StrSql = "SELECT * FROM rep67 "
'        StrSql = StrSql & " WHERE bpronro = " & bpronro
'        StrSql = StrSql & " AND pliqnro = " & Nroliq
'        StrSql = StrSql & " AND empresa = " & Empresa
'        StrSql = StrSql & " AND pronro = " & NroProc
'        StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'        OpenRecordset StrSql, rs_rep67
'
'        If rs_rep67.EOF Then
'            'Inserto
'            StrSql = "INSERT INTO rep67 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
'            StrSql = StrSql & "ternro,fechaalta,fechabaja,valor,nrodoc2) VALUES ("
'            StrSql = StrSql & bpronro & ","
'            StrSql = StrSql & Nroliq & ","
'            If Not Todos_Pro Then
'                StrSql = StrSql & rs_Procesos!pronro & ","
'                StrSql = StrSql & rs_Procesos!procaprob & ","
'            Else
'                StrSql = StrSql & "0" & ","
'                StrSql = StrSql & CInt(Proc_Aprob) & ","
'            End If
'            StrSql = StrSql & Empresa & ","
'            StrSql = StrSql & "'" & IdUser & "',"
'            StrSql = StrSql & ConvFecha(Fecha) & ","
'            StrSql = StrSql & "'" & Hora & "',"
'            StrSql = StrSql & rs_Empleados!ternro & ","
'            StrSql = StrSql & ConvFecha(fec_ini) & ","
'            StrSql = StrSql & ConvFecha(fec_fin) & ","
'            StrSql = StrSql & "0,' ')"
'            objConn.Execute StrSql, , adExecuteNoRecords
'        End If
'
'
'        Do While Not rs_Confrep.EOF
'            Select Case UCase(rs_Confrep!conftipo)
'            Case "AC": 'ACUMULADORES
'                StrSql = "SELECT * FROM aculiq " & _
'                         " INNER JOIN cabliq ON aculiq.cliqnro = " & rs_Empleados!cliqnro & _
'                         " WHERE aculiq.acunro =" & Par
'                OpenRecordset StrSql, rs_aculiq
'
'                If Not rs_aculiq.EOF Then
'                    'Actualizo
'                    StrSql = "UPDATE rep67 SET valor = valor + " & rs_aculiq!almonto
'                    StrSql = StrSql & " WHERE bpronro = " & bpronro
'                    StrSql = StrSql & " AND pliqnro = " & Nroliq
'                    StrSql = StrSql & " AND empresa = " & Empresa
'                    StrSql = StrSql & " AND pronro = " & NroProc
'                    StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'                    objConn.Execute StrSql, , adExecuteNoRecords
'
'
'                    StrSql = " SELECT cuil.nrodoc FROM tercero " & _
'                             " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) " & _
'                             " WHERE tercero.ternro= " & rs_Empleados!ternro
'                    OpenRecordset StrSql, rs_Cuil
'                    If Not rs_Cuil.EOF Then
'                        Aux_CUIL = "CUIL " & Left(CStr(rs_Cuil!nrodoc), 13)
'                        Aux_CUIL = Replace(CStr(Aux_CUIL), "-", "")
'
'                        StrSql = "UPDATE rep67 SET nrodoc2 = " & Aux_CUIL
'                        StrSql = StrSql & " WHERE bpronro = " & bpronro
'                        StrSql = StrSql & " AND pliqnro = " & Nroliq
'                        StrSql = StrSql & " AND empresa = " & Empresa
'                        StrSql = StrSql & " AND pronro = " & NroProc
'                        StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'                        objConn.Execute StrSql, , adExecuteNoRecords
'                    End If
'
'                    If TipoReporte <> 1 Then
'                        StrSql = " SELECT cuil.nrodoc FROM tercero " & _
'                                 " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro <= 4 AND cuail.nrodoc <> "") " & _
'                                 " WHERE tercero.ternro= " & rs_Empleados!ternro
'                        OpenRecordset StrSql, rs_Cuil
'                        If Not rs_Cuil.EOF Then
'                            Aux_CUIL = "CUIL " & Left(CStr(rs_Cuil!nrodoc), 13)
'                            Aux_CUIL = Replace(CStr(Aux_CUIL), "-", "")
'
'                            StrSql = "UPDATE rep67 SET nrodoc2 = " & Aux_CUIL
'                            StrSql = StrSql & " WHERE bpronro = " & bpronro
'                            StrSql = StrSql & " AND pliqnro = " & Nroliq
'                            StrSql = StrSql & " AND empresa = " & Empresa
'                            StrSql = StrSql & " AND pronro = " & NroProc
'                            StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'                            objConn.Execute StrSql, , adExecuteNoRecords
'                        End If
'                    End If
'                End If
'
'            Case "CO": 'CONCEPTOS
'                StrSql = "SELECT * FROM detliq " & _
'                         " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
'                         " WHERE concepto.concpuente = -1 " & _
'                         " AND concepto.tconnro =" & rs_Confrep!confval
'                OpenRecordset StrSql, rs_Detliq
'
'                If Not rs_Detliq.EOF Then
'                    'Actualizo
'                    StrSql = "UPDATE rep67 SET total_liquidado = total_liquidado + " & rs_Detliq!dlimonto
'                    StrSql = StrSql & ", cant_liquidado = cant_liquidado + " & rs_Detliq!dlicant
'                    StrSql = StrSql & ", emp_liquidado = emp_liquidado + 1 "
'                    StrSql = StrSql & " WHERE concnro = " & rs_Detliq!concnro
'                    StrSql = StrSql & " AND agrupacion = " & selector
'                    StrSql = StrSql & " AND agrupacion2 = " & selector2
'                    StrSql = StrSql & " AND agrupacion3 = " & selector3
'                    StrSql = StrSql & " AND bpronro = " & bpronro
'                    StrSql = StrSql & " AND pliqnro = " & Nroliq
'                    StrSql = StrSql & " AND empresa = " & Empresa
'                    StrSql = StrSql & " AND pronro = " & NroProc
'                    StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'                    objConn.Execute StrSql, , adExecuteNoRecords
'                End If
'
'            End Select
'
'                'Actualizo el progreso del Proceso
'                Progreso = Progreso + IncPorc
'                TiempoAcumulado = GetTickCount
'                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
'                         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
'                         "' WHERE bpronro = " & NroProcesoBatch
'                objConn.Execute StrSql, , adExecuteNoRecords
'
'            'Siguiente confrep
'            rs_Confrep.MoveNext
'        Loop
'
'        'Siguiente empleado
'        rs_Empleados.MoveNext
'    Loop
'
'    'Siguiente proceso
'    rs_Procesos.MoveNext
'Loop
'
''Fin de la transaccion
'MyCommitTrans
'
'
'If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
'If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
'If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
'If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
'If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
'If rs_rep67.State = adStateOpen Then rs_rep67.Close
'If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
'
'Set rs_Empleados = Nothing
'Set rs_Procesos = Nothing
'Set rs_Confrep = Nothing
'Set rs_Detliq = Nothing
'Set rs_rep67 = Nothing
'Set rs_Periodo = Nothing
'Set rs_Reporte = Nothing
'
'
'Exit Sub
'CE:
'    HuboError = True
'    MyRollbackTrans
'
'End Sub

Public Sub Admcer02_Empleo(ByVal TipoReporte As Integer, ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Todos_Empleados As Boolean, ByVal Fecha_Certificado As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Certificado de Trabajo
' Autor      : FGZ
' Fecha      : 05/03/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim fechadesde As Date
Dim fechahasta As Date
Dim fec_ini As Date
Dim fec_fin As Date

Dim selector  As Integer
Dim selector2 As Integer
Dim selector3 As Integer
Dim X As Boolean
Dim Par As Integer
Dim Aux_CUIL As String
Dim Aux_DNI As String

Dim Arreglo(20) As Single
Dim Total_liquidado As Single
Dim Cant_liquidado As Long
Dim Emp_liquidado As Integer
Dim I As Integer
Dim Opcion As Integer
Dim Aux_Localidad As String
Dim Aux_Catdesc As String
Dim TipoEstr As Integer
Dim TieneFechaBaja As Integer

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_aculiq As New ADODB.Recordset
Dim rs_rep67 As New ADODB.Recordset
Dim rs_Empleados As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_Cuil As New ADODB.Recordset
Dim rs_DNI As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_Zona As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

On Error GoTo CE

'Configuracion del Reporte
Select Case TipoReporte
Case 37: 'Certificado de Empleo
    StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 37"
Case 38: 'Certificado de Prestacion de Servicio
    StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 38"
Case 39: 'Certificado de DNRP
    StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 39"
End Select
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Período de Liquidación"
    Exit Sub
End If
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close

'Inicializo
Emp_liquidado = 0

For I = 1 To 20
    Arreglo(I) = 0
Next I

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep67 "
StrSql = StrSql & " WHERE tiporep = " & TipoReporte
'If Not Todos_Pro Then
'    StrSql = StrSql & " AND pronro = " & NroProc
'Else
'    StrSql = StrSql & " AND pronro = 0"
'    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
'End If
'StrSql = StrSql & " AND empresa = " & Empresa
objConn.Execute StrSql, , adExecuteNoRecords

'Busco los procesos a evaluar
'StrSql = "SELECT * FROM  empleado "
'FB - 17/02/2014- Se agregan los campos que utiliza la consulta, porque no traia bien los campos.
StrSql = " SELECT empleado.ternro, periodo.pliqdesde, periodo.pliqhasta, proceso.pronro, proceso.proaprob, empleado.empleg, empleado.ternom,"
StrSql = StrSql & " empleado.ternom2, empleado.terape, empleado.terape2, proceso.profecfin, cabliq.cliqnro FROM empleado"
If Not Todos_Empleados Then
    StrSql = StrSql & " INNER JOIN batch_empleado ON batch_empleado.ternro = empleado.ternro "
End If
StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = empleado.ternro and empresa.tenro = 10 "
StrSql = StrSql & " AND (empresa.htetdesde <= " & ConvFecha(Date) & ")"
StrSql = StrSql & " AND ((" & ConvFecha(Date) & " <= empresa.htethasta) or (empresa.htethasta is null)) "
'FGZ -  23/05/2013 ---------------------------------------------------------------
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN empresa emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
'FGZ -  23/05/2013 ---------------------------------------------------------------
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
StrSql = StrSql & " WHERE periodo.pliqnro =" & Nroliq
If Not Todos_Empleados Then
    StrSql = StrSql & " AND batch_empleado.bpronro =" & NroProcesoBatch
End If
If Not Todos_Pro Then
    'StrSql = StrSql & " AND proceso.pronro =" & NroProc
    StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")"
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
End If
StrSql = StrSql & " ORDER BY empleado.ternro"

Flog.writeline "SQL:" & StrSql

OpenRecordset StrSql, rs_Procesos

If Not rs_Procesos.EOF Then
    fechadesde = rs_Procesos!pliqdesde
    fechahasta = rs_Procesos!pliqhasta
Else
    Flog.writeline " No Encontro fechas desde y hasta del Periodo de Liquidacion de la SQL Anterior"
End If

Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
'CEmpleadosAProc = rs_Empleados.RecordCount
If CConceptosAProc = 0 Then
   CConceptosAProc = 1
End If
IncPorc = (100 / CConceptosAProc)

Do While Not rs_Procesos.EOF
    Total_liquidado = 0
    Cant_liquidado = 0
    Flog.writeline "   - Empleado " & rs_Procesos!empleg & " - " & rs_Procesos!terape & " " & rs_Procesos!terape & ", " & rs_Procesos!ternom & " " & rs_Procesos!ternom2
    
    StrSql = "SELECT * FROM fases WHERE empleado = " & rs_Procesos!Ternro
    StrSql = StrSql & " ORDER BY altfec "
    Flog.writeline " SQL: " & StrSql
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveFirst
        fec_ini = rs_Fases!altfec
        Flog.writeline " Fecha Inicio fase Empleado: " & fec_ini
        rs_Fases.MoveLast
        fec_fin = IIf(Not IsNull(rs_Fases!bajfec), rs_Fases!bajfec, rs_Procesos!profecfin)
        TieneFechaBaja = IIf(Not IsNull(rs_Fases!bajfec), -1, 0)
        Flog.writeline " Si fecha baja de la base es vacia, entonces Fecha fin igual fecha fin del proceso liq"
        Flog.writeline " Fecha Fin Fase Empleado  : " & fec_fin
    End If
        
    'If Not (TipoReporte <> 37 And (IsNull(fec_ini) Or (IsNull(fec_fin) Or Not rs_Fases!estado) Or (fec_ini > fec_fin))) Then
    'If Not ((IsNull(fec_ini) Or (IsNull(fec_fin) Or Not CBool(rs_Fases!estado)) Or (fec_ini > fec_fin))) Then
     
     ' Se valida que la fecha inicio de la fecha no sea nula o fecha fin sea nula o la fecha fin mayor a la fecha inicio
     
     If Not ((IsNull(fec_ini) Or (IsNull(fec_fin)) Or (fec_ini > fec_fin))) Then
      
        Flog.writeline "****************************************** "
'        Flog.writeline
        Flog.writeline " Comienzo del Proceso "
        Flog.writeline " Chequeo de fases Empleado Correcto"
'        Flog.writeline
        Flog.writeline "*********************************************"
         
         
        'busco la categoria del empleado
        StrSql = " SELECT * FROM his_estructura " & _
                 " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
                 " WHERE his_estructura.ternro = " & rs_Procesos!Ternro & " AND " & _
                 " his_estructura.tenro = 3 AND " & _
                 " (his_estructura.htetdesde <= " & ConvFecha(fec_fin) & ") AND " & _
                 " ((" & ConvFecha(fec_fin) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))" & _
                 " ORDER BY his_estructura.htetdesde"
        If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
       Flog.writeline "SQL: " & StrSql
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            Aux_Catdesc = rs_Estructura!estrdabr
        Else
            Flog.writeline "No se encontro la Categoria del empleado"
            Aux_Catdesc = " "
        End If
        
        
        'Si no existe el rep67
        StrSql = "SELECT * FROM rep67 "
        StrSql = StrSql & " WHERE bpronro = " & bpronro
        StrSql = StrSql & " AND ternro    = " & rs_Procesos!Ternro
        StrSql = StrSql & " AND pliqnro   = " & Nroliq
        StrSql = StrSql & " AND empresa   = " & Empresa
        'StrSql = StrSql & " AND pronro = '" & NroProc & "'"
        StrSql = StrSql & " AND pronro    = " & rs_Procesos!pronro
        StrSql = StrSql & " AND proaprob  = " & CInt(Proc_Aprob)
        OpenRecordset StrSql, rs_rep67
    
        If rs_rep67.EOF Then
            'Inserto
            StrSql = "INSERT INTO rep67 (apenom,bpronro,tiporep,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,catdesc,"
            StrSql = StrSql & "ternro,fechaalta,fechabaja,fecha_certificado,valor,nrodoc2,total_liquidado,cant_liquidado,emp_liquidado "
            StrSql = StrSql & ", col1,col2,col3,col4,col5,col6,col7,col8,col9,col10 "
            StrSql = StrSql & ", col11,col12,col13,col14,col15,col16,col17,col18,col19,col20) VALUES ("
            StrSql = StrSql & "'" & rs_Procesos!terape
            If Not IsNull(rs_Procesos!terape2) Then
                StrSql = StrSql & " " & rs_Procesos!terape2
            End If
            StrSql = StrSql & ", " & rs_Procesos!ternom
            If Not IsNull(rs_Procesos!ternom2) Then
                StrSql = StrSql & " " & rs_Procesos!ternom2
            End If
            StrSql = StrSql & "',"
            StrSql = StrSql & bpronro & ","
            StrSql = StrSql & TipoReporte & ","
            StrSql = StrSql & Nroliq & ","
            If Not Todos_Pro Then
                StrSql = StrSql & rs_Procesos!pronro & ","
                'StrSql = StrSql & "'" & NroProc & "',"
                StrSql = StrSql & rs_Procesos!proaprob & ","
            Else
                'StrSql = StrSql & "'0'" & ","
                StrSql = StrSql & rs_Procesos!pronro & ","
                StrSql = StrSql & CInt(Proc_Aprob) & ","
            End If
            StrSql = StrSql & Empresa & ","
            StrSql = StrSql & "'" & IdUser & "',"
            StrSql = StrSql & ConvFecha(Fecha) & ","
            StrSql = StrSql & "'" & Hora & "',"
            StrSql = StrSql & "'" & Aux_Catdesc & "',"
            StrSql = StrSql & rs_Procesos!Ternro & ","
            StrSql = StrSql & ConvFecha(fec_ini) & ","
            StrSql = StrSql & ConvFecha(fec_fin) & ","
            StrSql = StrSql & ConvFecha(Fecha_Certificado) & ","
            StrSql = StrSql & "0,' ',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0," & TieneFechaBaja & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
                
        Aux_Localidad = ""
        rs_Confrep.MoveFirst
        Do While Not rs_Confrep.EOF
            Select Case UCase(rs_Confrep!conftipo)
            Case "VAL": 'configracion de la zona
                Opcion = rs_Confrep!confval
                Flog.writeline " De acuerdo a la opcion busco la zona definida en el domicilio (1-Sucursal del Empleado, 2-Empresa del Empleado, 3-Empleado)"
                Select Case Opcion
                Case 1: 'Sucursal
                  Flog.writeline "  Busca la zona del domicio de la sucursal del Empleado"
                    TipoEstr = 1
                
                    StrSql = " SELECT estrnro FROM his_estructura " & _
                             " WHERE ternro = " & rs_Procesos!Ternro & " AND " & _
                             " tenro = 1 AND " & _
                             " (htetdesde <= " & ConvFecha(fec_fin) & ") AND " & _
                             " ((" & ConvFecha(fec_fin) & " <= htethasta) or (htethasta is null))"
                     Flog.writeline "SQL: " & StrSql
                    OpenRecordset StrSql, rs_Estructura
                
                    If Not rs_Estructura.EOF Then
                        StrSql = " SELECT * FROM sucursal " & _
                                 " WHERE estrnro =" & rs_Estructura!Estrnro
                        OpenRecordset StrSql, rs_Sucursal
                        
                        If Not rs_Sucursal.EOF Then
                            StrSql = " SELECT detdom.zonanro,localidad.locdesc FROM detdom " & _
                                     " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                                     " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
                                     " WHERE cabdom.ternro = " & rs_Sucursal!Ternro '& " AND " & _
                                     '" cabdom.tipnro =" & TipoEstr
                            OpenRecordset StrSql, rs_Zona
                            If Not rs_Zona.EOF Then
                                Aux_Localidad = rs_Zona!locdesc
                            End If
                        End If  ' If Not rs_Sucursal.EOF Then
                    End If ' If Not rs_Estructura.EOF Then
                Case 2: 'Empresa
                    Flog.writeline "  Busca la zona del domicio de la empresa del Empleado"
                    'Cargo el tipo de estructura segun sea sucursal o Empresa
                    TipoEstr = 12
                    StrSql = " SELECT estrnro FROM his_estructura " & _
                             " WHERE ternro = " & rs_Procesos!Ternro & " AND " & _
                             " tenro = 10 AND " & _
                             " (htetdesde <= " & ConvFecha(fec_fin) & ") AND " & _
                             " ((" & ConvFecha(fec_fin) & " <= htethasta) or (htethasta is null))"
                    OpenRecordset StrSql, rs_Estructura
                
                    If Not rs_Estructura.EOF Then
                        StrSql = " SELECT * FROM empresa " & _
                                 " WHERE estrnro =" & rs_Estructura!Estrnro
                        OpenRecordset StrSql, rs_Sucursal
                        
                        If Not rs_Sucursal.EOF Then
                            StrSql = " SELECT detdom.zonanro,localidad.locdesc FROM detdom " & _
                                     " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                                     " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
                                     " WHERE cabdom.ternro = " & rs_Sucursal!Ternro '& " AND " & _
                                     '" cabdom.tipnro =" & TipoEstr
                            OpenRecordset StrSql, rs_Zona
                            If Not rs_Zona.EOF Then
                                Aux_Localidad = rs_Zona!locdesc
                            End If
                        End If  ' If Not rs_Sucursal.EOF Then
                    End If ' If Not rs_Estructura.EOF Then
                Case 3: 'Empleado
                    Flog.writeline "  Busca la zona del domicio del Empleado"
                    TipoEstr = 1
                    StrSql = " SELECT localidad.locdesc FROM detdom " & _
                             " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                             " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
                             " WHERE cabdom.ternro = " & rs_Procesos!Ternro '& " AND " & _
                             '" cabdom.tipnro =" & TipoEstr
                    OpenRecordset StrSql, rs_Zona
                    If Not rs_Zona.EOF Then
                        Aux_Localidad = rs_Zona!locdesc
                    End If
                End Select
            Case "AC": 'ACUMULADORES
                StrSql = "SELECT * FROM acu_liq " & _
                         " WHERE acu_liq.acunro =" & rs_Confrep!confval & _
                         " AND acu_liq.cliqnro = " & rs_Procesos!cliqnro
                OpenRecordset StrSql, rs_aculiq
            
                Do While Not rs_aculiq.EOF
                    Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_aculiq!almonto
                    
                    rs_aculiq.MoveNext
                Loop
            
            Case "CO": 'CONCEPTOS
                StrSql = "SELECT * FROM detliq " & _
                         " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
                         " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
                         " WHERE concepto.conccod =" & rs_Confrep!confval & _
                         " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
                OpenRecordset StrSql, rs_Detliq
            
                Do While Not rs_Detliq.EOF
                    Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlimonto
                    
                    Total_liquidado = Total_liquidado + rs_Detliq!dlimonto
                    Cant_liquidado = Cant_liquidado + rs_Detliq!dlicant
                    
                    rs_Detliq.MoveNext
                Loop
                    
            End Select
                
            'Siguiente confrep
            rs_Confrep.MoveNext
        Loop
                        
        Emp_liquidado = Emp_liquidado + 1
                
        'busco el dni
        StrSql = " SELECT * FROM tercero " & _
                 " INNER JOIN ter_doc dni ON (tercero.ternro = dni.ternro AND dni.tidnro <= 4) " & _
                 " INNER JOIN tipodocu ON dni.tidnro = tipodocu.tidnro " & _
                 " WHERE tercero.ternro= " & rs_Procesos!Ternro
        OpenRecordset StrSql, rs_DNI
        Flog.writeline "   - Busco el Documento del empleado "
        If Not rs_DNI.EOF Then
            Aux_DNI = rs_DNI!tidsigla & " " & CStr(rs_DNI!NroDoc)
        Else
            Aux_DNI = "N.D."
        End If
                
                
        'Actualizo
        StrSql = "UPDATE rep67 SET "
        StrSql = StrSql & " total_liquidado = total_liquidado + " & Total_liquidado
        StrSql = StrSql & " ,nrodoc = '" & Aux_DNI & "'"
        StrSql = StrSql & " ,cant_liquidado = cant_liquidado + " & Cant_liquidado
        StrSql = StrSql & " ,emp_liquidado = " & Emp_liquidado
        StrSql = StrSql & " ,locdesc = '" & Aux_Localidad & "'"
        For I = 1 To 20
            If Arreglo(I) <> 0 Then
                StrSql = StrSql & " ,col" & I & " = col" & I & " + " & Arreglo(I)
            End If
        Next I
        StrSql = StrSql & " WHERE bpronro = " & bpronro
        StrSql = StrSql & " AND ternro = " & rs_Procesos!Ternro
        StrSql = StrSql & " AND pliqnro = " & Nroliq
        StrSql = StrSql & " AND empresa = " & Empresa
        'StrSql = StrSql & " AND pronro = '" & NroProc & "'"
        StrSql = StrSql & " AND pronro = " & rs_Procesos!pronro
        StrSql = StrSql & " AND proaprob= " & Proc_Aprob
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        StrSql = " SELECT cuil.nrodoc FROM tercero " & _
                 " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) " & _
                 " WHERE tercero.ternro= " & rs_Procesos!Ternro
        OpenRecordset StrSql, rs_Cuil
        Flog.writeline "   - Busco el CUIL. SQL ==> " & StrSql
        If Not rs_Cuil.EOF Then
            Aux_CUIL = Left(CStr(rs_Cuil!NroDoc), 13)
           ' Aux_CUIL = Replace(CStr(Aux_CUIL), "-", "")
            
            StrSql = "UPDATE rep67 SET nrodoc2 = '" & Aux_CUIL & "'"
            StrSql = StrSql & " WHERE bpronro = " & bpronro
            StrSql = StrSql & " AND ternro = " & rs_Procesos!Ternro
            StrSql = StrSql & " AND pliqnro = " & Nroliq
            StrSql = StrSql & " AND empresa = " & Empresa
            'StrSql = StrSql & " AND pronro = " & NroProc
            StrSql = StrSql & " AND pronro = " & rs_Procesos!pronro
            StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        If TipoReporte <> 37 And TipoReporte <> 38 Then
            StrSql = " SELECT cuil.nrodoc FROM tercero " & _
                     " INNER JOIN ter_doc cuil ON tercero.ternro = cuil.ternro" & _
                     " WHERE tercero.ternro= " & rs_Procesos!Ternro & _
                     " AND cuil.tidnro <= 4 AND cuil.nrodoc <> ''"
            OpenRecordset StrSql, rs_Cuil
            If Not rs_Cuil.EOF Then
                Flog.writeline "   - En el CUIL cargo el Documento. SQL ==> " & StrSql
                Aux_CUIL = Left(CStr(rs_Cuil!NroDoc), 13)
                Aux_CUIL = Replace(CStr(Aux_CUIL), "-", "")
                
                StrSql = "UPDATE rep67 SET nrodoc2 = '" & Aux_CUIL & "'"
                StrSql = StrSql & " WHERE bpronro = " & bpronro
                StrSql = StrSql & " AND ternro  = " & rs_Procesos!Ternro
                StrSql = StrSql & " AND pliqnro = " & Nroliq
                StrSql = StrSql & " AND empresa = " & Empresa
                'StrSql = StrSql & " AND pronro = " & NroProc
                StrSql = StrSql & " AND pronro   =  " & rs_Procesos!pronro
                StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    Else
    
        Flog.writeline "     Error en las fases del empleado. SQL ==> " & StrSql
    End If
    
    'Actualizo el progreso del Proceso
    Progreso = Progreso + IncPorc
    'Flog.writeline "------------------------------------------------"
    'Flog.writeline "Progreso: " & Progreso & " %"
    'Flog.writeline "------------------------------------------------"
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_rep67.State = adStateOpen Then rs_rep67.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Sucursal.State = adStateOpen Then rs_Sucursal.Close
If rs_Zona.State = adStateOpen Then rs_Zona.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_DNI.State = adStateOpen Then rs_DNI.Close

Set rs_Empleados = Nothing
Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_rep67 = Nothing
Set rs_Periodo = Nothing
Set rs_Reporte = Nothing
Set rs_Sucursal = Nothing
Set rs_Zona = Nothing
Set rs_Estructura = Nothing
Set rs_DNI = Nothing
Exit Sub
CE:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboError = True
    MyRollbackTrans
End Sub
Public Sub DatosCertifSalarial(ByVal TipoReporte As Integer, ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Todos_Empleados As Boolean, ByVal Fecha_Certificado As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Certificado Salarial
' Autor      : Manterola Maria Magdalena
' Fecha      : 05/10/2011
' Ult. Mod   : 02/11/2012 - Gonzalez Nicolás -
' --------------------------------------------------------------------------------------------
Dim fechadesde As Date
Dim fechahasta As Date
Dim fec_ini As Date
Dim fec_fin As Date

Dim Aux_DNI As String

Dim bruto As Double
Dim neto As Double

Dim Emp_liquidado As Integer
Dim I As Integer
Dim Ternro As Integer
Dim Aux_Localidad As String
Dim Aux_Catdesc As String
Dim Tiene_Embargo As String

Dim rs_Embargo As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Acu As New ADODB.Recordset
Dim rs_rep67 As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_DNI As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rsConsult As New ADODB.Recordset

Dim UltAC As String
StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro "
StrSql = StrSql & " WHERE reporte.repnro = 358"
StrSql = StrSql & " AND UPPER(conftipo) = 'AC'"
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub

Else 'A partir de la V 1.07
    If Left(ListaNroProc, 2) = "0," Then
        ListaNroProc = Mid(ListaNroProc, 3, Len(ListaNroProc))
    End If
    
    UltAC = UCase(EsNulo(rs_Confrep!confval2))
    Do While Not rs_Confrep.EOF
        
        'Valido si usa acu_liq ó acu_mes
        If (UCase(EsNulo(rs_Confrep!confval2)) = "ACL" And Todos_Pro = True) Or InStr(1, ListaNroProc, ",") > 0 Then
            Flog.writeline "Solo se debe seleccionar un proceso cuando los AC estan configurados como ACL"
            Exit Sub
        End If
        
        'Valido que la columna AC este bien configurada
        If UltAC <> UCase(EsNulo(rs_Confrep!confval2)) Then
            Flog.writeline "Error de configuración en los AC. Se buscan valores de los acumuladores mensuales"
            Flog.writeline ""
        End If
        UltAC = UCase(EsNulo(rs_Confrep!confval2))
        
        rs_Confrep.MoveNext
    Loop
End If

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close

'Inicializo
Emp_liquidado = 0

' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
StrSql = "DELETE FROM rep67 "
StrSql = StrSql & " WHERE tiporep = " & TipoReporte

objConn.Execute StrSql, , adExecuteNoRecords

'Busco los procesos a evaluar
'StrSql = "SELECT * FROM  empleado "
'Agregado - 23/06/2014 - Se agregan los campos que utiliza la consulta, porque no traia bien los campos.
StrSql = " SELECT empleado.ternro, periodo.pliqdesde, periodo.pliqhasta, proceso.pronro, proceso.proaprob, empleado.empleg, empleado.ternom,"
StrSql = StrSql & " empleado.ternom2, empleado.terape, empleado.terape2, proceso.profecfin, cabliq.cliqnro FROM empleado"
If Not Todos_Empleados Then
    StrSql = StrSql & " INNER JOIN batch_empleado ON batch_empleado.ternro = empleado.ternro "
End If
StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = empleado.ternro and empresa.tenro = 10 "
StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro " 'AND emp.empnro =" & Empresa
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
StrSql = StrSql & " WHERE periodo.pliqnro =" & Nroliq
If Not Todos_Empleados Then
    StrSql = StrSql & " AND batch_empleado.bpronro =" & NroProcesoBatch
End If
If Not Todos_Pro Then
    StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")"
Else
    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
    StrSql = StrSql & " AND proceso.empnro = " & Empresa
End If
StrSql = StrSql & " ORDER BY empleado.ternro"

OpenRecordset StrSql, rs_Procesos

If Not rs_Procesos.EOF Then
    fechadesde = rs_Procesos!pliqdesde
    fechahasta = rs_Procesos!pliqhasta
End If

Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount


        
        If CConceptosAProc = 0 Then
           Flog.writeline "no hay procesos"
           CConceptosAProc = 1
           IncPorc = 100
        Else
            'IncPorc = Int(100 / CEmpleadosAProc)
            IncPorc = 99 / CConceptosAProc
            Flog.writeline "INCPORC:" & IncPorc
        End If
'If CConceptosAProc = 0 Then
'   CConceptosAProc = 1
'End If
'IncPorc = (100 / CConceptosAProc)
Ternro = 0
Do While Not rs_Procesos.EOF
    If Ternro <> rs_Procesos!Ternro Then
        neto = 0
        bruto = 0
        
        Ternro = rs_Procesos!Ternro

        Flog.writeline "   - Empleado " & rs_Procesos!empleg & " - " & rs_Procesos!terape & " " & rs_Procesos!terape2 & ", " & rs_Procesos!ternom & " " & rs_Procesos!ternom2
        
        StrSql = "SELECT * FROM fases WHERE empleado = " & rs_Procesos!Ternro
        StrSql = StrSql & " ORDER BY altfec DESC"
        OpenRecordset StrSql, rs_Fases
        
        If Not rs_Fases.EOF Then
            
            'fec_ini = rs_Fases!altfec
           rs_Fases.MoveFirst
        fec_ini = rs_Fases!altfec
        Flog.writeline " Fecha Inicio fase Empleado: " & fec_ini
        'Comentado el 02/08/2013
        'rs_Fases.MoveLast
        'fin
        fec_fin = IIf(Not IsNull(rs_Fases!bajfec), rs_Fases!bajfec, rs_Procesos!profecfin)
        Flog.writeline " Si fecha baja de la base es vacia, entonces Fecha fin igual fecha fin del proceso liq"
        Flog.writeline " Fecha Fin Fase Empleado  : " & fec_fin
        End If
            
         If Not ((IsNull(fec_ini) Or (IsNull(fec_fin)) Or (fec_ini > fec_fin))) Then
        'If Not ((IsNull(fec_ini) Or (IsNull(fec_fin) Or Not CBool(rs_Fases!estado)) Or (fec_ini > fec_fin))) Then
        'If Not (IsNull(fec_ini) Or Not CBool(rs_Fases!estado)) Then
                    
            'busco el puesto del empleado y lo guardo en la categoria
            StrSql = " SELECT * FROM his_estructura " & _
                     " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro " & _
                     " WHERE his_estructura.ternro = " & rs_Procesos!Ternro & " AND " & _
                     " his_estructura.tenro = 4 AND " & _
                     " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Certificado) & ") AND " & _
                     " ((" & ConvFecha(Fecha_Certificado) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))" & _
                     " ORDER BY his_estructura.htetdesde"
            If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
            Flog.writeline "   - Busco el puesto del empleado. SQL ==> " & StrSql
            OpenRecordset StrSql, rs_Estructura
            
            If Not rs_Estructura.EOF Then
                Aux_Catdesc = rs_Estructura!estrdabr
            Else
                Flog.writeline "No se encontro el Puesto del empleado"
                Aux_Catdesc = " "
            End If
                            
            'Si no existe el rep67
            StrSql = "SELECT * FROM rep67 "
            StrSql = StrSql & " WHERE bpronro = " & bpronro
            StrSql = StrSql & " AND ternro = " & rs_Procesos!Ternro
            StrSql = StrSql & " AND pliqnro = " & Nroliq
            StrSql = StrSql & " AND empresa = " & Empresa
            StrSql = StrSql & " AND pronro = '" & NroProc & "'"
            StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
            OpenRecordset StrSql, rs_rep67
        
            If rs_rep67.EOF Then
                'Inserto
                StrSql = "INSERT INTO rep67 (apenom,bpronro,tiporep,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,catdesc,"
                StrSql = StrSql & "ternro,fechaalta,fechabaja,fecha_certificado,valor,nrodoc2,total_liquidado,cant_liquidado,emp_liquidado "
                StrSql = StrSql & ", col1,col2,col3,col4,col5,col6,col7,col8,col9,col10 "
                StrSql = StrSql & ", col11,col12,col13,col14,col15,col16,col17,col18,col19,col20) VALUES ("
                StrSql = StrSql & "'" & rs_Procesos!terape
                Dim teraux As String
                Dim teraux2 As String
                teraux = IIf(IsNull(rs_Procesos!terape2), "", rs_Procesos!terape2)
                If (teraux <> "") Then 'Manterola Maria Magdalena 31/10/2011
                    StrSql = StrSql & " " & teraux
                End If
                StrSql = StrSql & ", " & rs_Procesos!ternom
                
                teraux2 = IIf(IsNull(rs_Procesos!ternom2), "", rs_Procesos!ternom2)
                If (teraux2 <> "") Then 'Manterola Maria Magdalena 31/10/2011
                    StrSql = StrSql & " " & teraux2
                End If
                
                StrSql = StrSql & "',"
                StrSql = StrSql & bpronro & ","
                StrSql = StrSql & TipoReporte & ","
                StrSql = StrSql & Nroliq & ","
                If Not Todos_Pro Then
                    'StrSql = StrSql & rs_Procesos!pronro & ","
                    StrSql = StrSql & "'" & NroProc & "',"
                    StrSql = StrSql & rs_Procesos!proaprob & ","
                Else
                    StrSql = StrSql & "'0'" & ","
                    StrSql = StrSql & CInt(Proc_Aprob) & ","
                End If
                StrSql = StrSql & Empresa & ","
                StrSql = StrSql & "'" & IdUser & "',"
                StrSql = StrSql & ConvFecha(Fecha) & ","
                StrSql = StrSql & "'" & Hora & "',"
                StrSql = StrSql & "'" & Aux_Catdesc & "',"
                StrSql = StrSql & rs_Procesos!Ternro & ","
                StrSql = StrSql & ConvFecha(fec_ini) & ","
                If Not IsNull(rs_Fases!bajfec) Then
                    StrSql = StrSql & ConvFecha(rs_Fases!bajfec) & ","
                Else
                    StrSql = StrSql & "'" & Null & "',"
                End If
                StrSql = StrSql & ConvFecha(Fecha_Certificado) & ","
                StrSql = StrSql & "0,' ',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)"
                
                objConn.Execute StrSql, , adExecuteNoRecords
    
            End If
                    
            Aux_Localidad = ""
            
            
            StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro "
            StrSql = StrSql & " WHERE reporte.repnro = 358"
            StrSql = StrSql & " AND UPPER(conftipo) = 'AC'"
            'rs_Confrep.MoveFirst
            OpenRecordset StrSql, rs_Confrep
            Do While Not rs_Confrep.EOF
                
                If UCase(rs_Confrep!conftipo) = "AC" Then
                    StrSql = " SELECT * FROM acumulador "
                    StrSql = StrSql & " WHERE acunro = " & rs_Confrep!confval
                    
                    OpenRecordset StrSql, rs_Acu
                                
                    If Not rs_Acu.EOF Then
                    
                        If UCase(rs_Confrep!confval2) = "ACL" Then
                            'SI LOS AC SON DE TIPO ACL BUSCO ALMONTO EN ACU_LIQ
                            StrSql = "SELECT * FROM acu_liq WHERE acunro = " & rs_Confrep!confval
                            StrSql = StrSql & " AND cliqnro = " & rs_Procesos!cliqnro
                            OpenRecordset StrSql, rsConsult
                                                                            
                            If rs_Confrep!confnrocol = 1 Then
                                'Do Until rsConsult.EOF
                                 If Not rsConsult.EOF Then
                                    'BRUTO
                                    bruto = rsConsult!almonto
                                    'rsConsult.MoveNext
                                End If
                                'Loop
                            Else
                                'Do Until rsConsult.EOF
                                If Not rsConsult.EOF Then
                                    'NETO
                                    neto = rsConsult!almonto
                                    'rsConsult.MoveNext
                                End If
                                'Loop
                            End If

                        
                        Else
                            StrSql = "SELECT ammonto FROM acu_mes WHERE ternro = " & rs_Procesos!Ternro & " AND acunro = " & rs_Confrep!confval
                            StrSql = StrSql & " AND amanio = " & Year(rs_Procesos!profecfin) & " AND ammes = " & Month(rs_Procesos!profecfin)
                        
                            OpenRecordset StrSql, rsConsult
                                                                            
                            If rs_Confrep!confnrocol = 1 Then
                                'Do Until rsConsult.EOF
                                 If Not rsConsult.EOF Then
                                    'BRUTO
                                    bruto = rsConsult!ammonto
                                    'rsConsult.MoveNext
                                End If
                                'Loop
                            Else
                                'Do Until rsConsult.EOF
                                If Not rsConsult.EOF Then
                                    'NETO
                                    neto = rsConsult!ammonto
                                    'rsConsult.MoveNext
                                End If
                                'Loop
                            End If
                        
                        End If
                        
                        
                    Else
                        Flog.writeline "No se encontró el acumulador"
                    End If
                    
                    
                Else 'CIERRA IF
                    Flog.writeline "Error en la configuración del acumulador. El tipo debe ser AC"
                End If
                    
                'Siguiente confrep
                rs_Confrep.MoveNext
            Loop
    
            'Ahora busco la configuracion de la forma de liquidacion activa del empleado,
            'Segun lo configurado en el Confrep
            StrSql = " SELECT confetiq,conftipo FROM his_estructura "
            StrSql = StrSql & " INNER JOIN reporte ON reporte.repnro = 358 "
            StrSql = StrSql & " INNER JOIN confrep ON reporte.repnro = confrep.repnro "
            StrSql = StrSql & " AND (UPPER(conftipo) = 'SM' OR UPPER(conftipo) = 'NSM')"
            StrSql = StrSql & " WHERE ternro = " & rs_Procesos!Ternro
            StrSql = StrSql & " AND tenro = 22 "
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Fecha_Certificado) & ")"
            StrSql = StrSql & " AND estrnro = confrep.confval "
            StrSql = StrSql & " AND ((" & ConvFecha(Fecha_Certificado) & " <= htethasta) or (htethasta is null))"
            StrSql = StrSql & " ORDER BY htetdesde DESC,confnrocol DESC"
            OpenRecordset StrSql, rs_Estructura
            Flog.writeline "   - Busco la forma de liquidacion activa del empleado. SQL ==> " & StrSql
            If Not rs_Estructura.EOF Then
                If UCase(rs_Estructura!conftipo) = "NSM" Then
                    Aux_Localidad = "NO_TIENE"
                Else
                    Aux_Localidad = rs_Estructura!confetiq
                End If
            End If
                
            Emp_liquidado = Emp_liquidado + 1
                    
            'busco el dni
            StrSql = " SELECT * FROM tercero " & _
                     " INNER JOIN ter_doc dni ON (tercero.ternro = dni.ternro AND dni.tidnro <= 4) " & _
                     " INNER JOIN tipodocu ON dni.tidnro = tipodocu.tidnro " & _
                     " WHERE tercero.ternro= " & rs_Procesos!Ternro
            OpenRecordset StrSql, rs_DNI
            Flog.writeline "   - Busco el Documento. SQL ==> " & StrSql
            If Not rs_DNI.EOF Then
                Aux_DNI = CStr(rs_DNI!NroDoc)
            Else
                Aux_DNI = "N.D."
            End If
                    
            'busco si el empleado tiene embargos
            StrSql = " SELECT * FROM embargo "
            StrSql = StrSql & " WHERE embargo.ternro= " & rs_Procesos!Ternro
            'FGZ - 23/05/2013 ---------------
            StrSql = StrSql & " AND embest = 'A'"
            'FGZ - 23/05/2013 ---------------
            OpenRecordset StrSql, rs_Embargo
            Flog.writeline "   - Busco si el empleado tiene embargos. SQL ==> " & StrSql
            If Not rs_Embargo.EOF Then
                Tiene_Embargo = "SI"
                Flog.writeline "      - SI TIENE EMBARGOS!"
            Else
                Tiene_Embargo = "NO"
                Flog.writeline "      - NO TIENE EMBARGOS!"
            End If
                    
            'Actualizo
            StrSql = "UPDATE rep67 SET "
            'EN TOTAL LIQUIDADO ---> PONGO EL BRUTO!!!
            StrSql = StrSql & " total_liquidado = " & bruto
            StrSql = StrSql & " ,nrodoc = '" & Aux_DNI & "'"
            'EN CANT LIQUIDADO ---> PONGO EL NETO!!!
            StrSql = StrSql & " ,cant_liquidado = " & neto
            StrSql = StrSql & " ,emp_liquidado = " & Emp_liquidado
            'EN AUX LOCALIDAD ---> PONGO LA FORMA DE LIQUIDACION DEL EMPLEADO!!!
            StrSql = StrSql & " ,locdesc = '" & Aux_Localidad & "'"
            'EN NRODOC ---> PONGO SI EL EMPLEADO TIENE EMBARGOS O NO!!!
            StrSql = StrSql & " ,nrodoc2 = '" & Tiene_Embargo & "'"
            StrSql = StrSql & " WHERE bpronro = " & bpronro
            StrSql = StrSql & " AND ternro = " & rs_Procesos!Ternro
            StrSql = StrSql & " AND pliqnro = " & Nroliq
            StrSql = StrSql & " AND empresa = " & Empresa
            StrSql = StrSql & " AND pronro = '" & NroProc & "'"
            StrSql = StrSql & " AND proaprob= " & Proc_Aprob
            objConn.Execute StrSql, , adExecuteNoRecords
            
        Else
            Flog.writeline "     Error en las fases del empleado. SQL ==> " & StrSql
        End If
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    End If
    
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans


If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_rep67.State = adStateOpen Then rs_rep67.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_DNI.State = adStateOpen Then rs_DNI.Close
If rs_Embargo.State = adStateOpen Then rs_Embargo.Close
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rsConsult.State = adStateOpen Then rsConsult.Close

Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_rep67 = Nothing
Set rs_Periodo = Nothing
Set rs_Estructura = Nothing
Set rs_DNI = Nothing
Set rs_Embargo = Nothing
Set rs_Acu = Nothing
Set rs_Fases = Nothing
Set rsConsult = Nothing

Exit Sub
CE:
    HuboError = True
    MyRollbackTrans
End Sub


Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer

Dim Nroliq As Long
Dim Todos_Pro As Boolean
Dim Todos_Empleados As Boolean
Dim Proc_Aprob As Integer
Dim Empresa As Long
Dim Todos_Sindicatos As Boolean
Dim Nro_Sindicato As Long
Dim TipoReporte As Integer
Dim Fecha_Certificado As Date

Dim Tenro1 As Long
Dim Tenro2 As Long
Dim Tenro3 As Long

Dim Estrnro1 As Long
Dim Estrnro2 As Long
Dim Estrnro3 As Long

Dim AgrupaTE1 As Boolean
Dim AgrupaTE2 As Boolean
Dim AgrupaTE3 As Boolean

Dim Agrupado As Boolean

'Inicializacion
Agrupado = False
Tenro1 = 0
Tenro2 = 0
Tenro3 = 0
AgrupaTE1 = False
AgrupaTE2 = False
AgrupaTE3 = False

On Error GoTo CE

' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, ".") - 1
        Nroliq = CLng(Mid(parametros, pos1, pos2))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Todos_Pro = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        If Not Todos_Pro Then
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, ".") - 1
            NroProc = Mid(parametros, pos1, pos2 - pos1 + 1)
            ListaNroProc = Replace(NroProc, "-", ",")
        Else
            NroProc = "0"
            ListaNroProc = "0"
        End If
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Proc_Aprob = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Empresa = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        Todos_Empleados = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, ".") - 1
        TipoReporte = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        Fecha_Certificado = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
    End If
End If

Select Case TipoReporte
Case 37, 38, 39: 'Certificado de Empleo
    Call Admcer02_Empleo(TipoReporte, bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Todos_Empleados, Fecha_Certificado)
'Case 38: 'Certificado de Prestacion de Servicio
'    Call Admcer02_Prestacion(TipoReporte, bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Todos_Empleados)
'Case 39: 'Certificado de DNRP
'    Call Admcer02_DNRP(TipoReporte, bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Todos_Empleados)
Case 358: 'Certificado Salarial
    Call DatosCertifSalarial(TipoReporte, bpronro, Nroliq, Todos_Pro, Proc_Aprob, Empresa, Todos_Empleados, Fecha_Certificado)
End Select


Exit Sub
CE:
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    Flog.writeline Espacios(Tabulador * 1) & "Error levantando parametros: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Última SQL ejecutada: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    HuboError = True
End Sub


Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para saber si es el ultimo empleado de la secuencia
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
    
    rs.MoveNext
    If rs.EOF Then
        EsElUltimoEmpleado = True
    Else
        If rs!Empleado <> Anterior Then
            EsElUltimoEmpleado = True
        Else
            EsElUltimoEmpleado = False
        End If
    End If
    rs.MovePrevious
End Function

'Public Sub Admcer02_DNRP(ByVal TipoReporte As Integer, ByVal bpronro As Long, ByVal Nroliq As Long, ByVal Todos_Pro As Boolean, ByVal Proc_Aprob As Integer, ByVal Empresa As Long, ByVal Todos_Empleados As Boolean)
'' --------------------------------------------------------------------------------------------
'' Descripcion: Procedimiento de generacion del Reporte de Certificado de Trabajo
'' Autor      : FGZ
'' Fecha      : 05/03/2004
'' Ult. Mod   :
'' Fecha      :
'' --------------------------------------------------------------------------------------------
'Dim fechadesde As Date
'Dim fechahasta As Date
'Dim fec_ini As Date
'Dim fec_fin As Date
'
'Dim selector  As Integer
'Dim selector2 As Integer
'Dim selector3 As Integer
'Dim X As Boolean
'Dim Par As Integer
'Dim Aux_CUIL As String
'
'
'Dim rs_Reporte As New ADODB.Recordset
'Dim rs_Periodo As New ADODB.Recordset
'Dim rs_Procesos As New ADODB.Recordset
'Dim rs_Confrep As New ADODB.Recordset
'Dim rs_Detliq As New ADODB.Recordset
'Dim rs_aculiq As New ADODB.Recordset
'Dim rs_rep67 As New ADODB.Recordset
'Dim rs_Empleados As New ADODB.Recordset
'Dim rs_Fases As New ADODB.Recordset
'Dim rs_Cuil As New ADODB.Recordset
'
''Configuracion del Reporte
'Select Case TipoReporte
'Case 37: 'Certificado de Empleo
'    StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 37"
'Case 38: 'Certificado de Prestacion de Servicio
'    StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 38"
'Case 39: 'Certificado de DNRP
'    StrSql = "SELECT * FROM reporte INNER JOIN confrep ON  reporte.repnro = confrep.repnro WHERE reporte.repnro = 39"
'End Select
'OpenRecordset StrSql, rs_Confrep
'If rs_Confrep.EOF Then
'    Flog.writeline "No se encontró la configuración del Reporte"
'    Exit Sub
'End If
'
''cargo el periodo
'StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
'OpenRecordset StrSql, rs_Periodo
'If rs_Periodo.EOF Then
'    Flog.writeline "No se encontró el Periodo"
'    Exit Sub
'End If
'If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
'
'' Comienzo la transaccion
'MyBeginTrans
'
''Depuracion del Temporario
'StrSql = "DELETE FROM rep67 "
'StrSql = StrSql & " WHERE pliqnro = " & Nroliq
'If Not Todos_Pro Then
'    StrSql = StrSql & " AND pronro = " & NroProc
'Else
'    StrSql = StrSql & " AND pronro = 0"
'    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
'End If
'StrSql = StrSql & " AND empresa = " & Empresa
'objConn.Execute StrSql, , adExecuteNoRecords
'
''Busco los procesos a evaluar
'StrSql = "SELECT * FROM  empleado "
'If Not Todos_Empleados Then
'    StrSql = StrSql & " INNER JOIN batch_empleado ON batch_empleado.ternro = empleado.ternro "
'End If
'StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
'StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
'StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
'StrSql = StrSql & " WHERE periodo.pliqnro =" & Nroliq
''StrSql = StrSql & " AND periodo.empnro =" & Empresa
'If Not Todos_Pro Then
'    StrSql = StrSql & " AND proceso.pronro =" & NroProc
'Else
'    StrSql = StrSql & " AND proaprob = " & CInt(Proc_Aprob)
'End If
'StrSql = StrSql & " ORDER BY empleado.ternro"
'OpenRecordset StrSql, rs_Procesos
'
'If Not rs_Procesos.EOF Then
'    fechadesde = rs_Procesos!pliqdesde
'    fechahasta = rs_Procesos!pliqhasta
'End If
'
'Do While Not rs_Procesos.EOF
'    StrSql = "SELECT * FROM batch_empleado " & _
'             " INNER JOIN cabliq ON cabliq.empleado = batch_empleado.ternro " & _
'             " WHERE cabliq.pronro =" & rs_Procesos!pronro
'    OpenRecordset StrSql, rs_Empleados
'
'    'seteo de las variables de progreso
'    Progreso = 0
'    CConceptosAProc = rs_Procesos.RecordCount
'    CEmpleadosAProc = rs_Empleados.RecordCount
'    If CEmpleadosAProc = 0 Then
'       CEmpleadosAProc = 1
'    End If
'    IncPorc = ((100 / CEmpleadosAProc) * (100 / CConceptosAProc)) / 100
'
'    Do While Not rs_Empleados.EOF
'
'        StrSql = "SELECT * FROM fases WHERE empleado = " & rs_Empleados!ternro
'        OpenRecordset StrSql, rs_Fases
'
'        If Not rs_Fases.EOF Then
'            rs_Fases.MoveFirst
'            fec_ini = rs_Fases!altfec
'
'            rs_Fases.MoveLast
'            fec_fin = rs_Fases!bajfec
'        End If
'
'        If TipoReporte <> 1 And (IsNull(fec_ini) Or (IsNull(fec_fin) Or Not rs_Fases!estado) Or (fec_ini > fec_fin)) Then
'            Flog.writeline "Error en las fases del empleado"
'            Exit Sub
'        End If
'
'        'Si no existe el rep67
'        StrSql = "SELECT * FROM rep67 "
'        StrSql = StrSql & " WHERE bpronro = " & bpronro
'        StrSql = StrSql & " AND pliqnro = " & Nroliq
'        StrSql = StrSql & " AND empresa = " & Empresa
'        StrSql = StrSql & " AND pronro = " & NroProc
'        StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'        OpenRecordset StrSql, rs_rep67
'
'        If rs_rep67.EOF Then
'            'Inserto
'            StrSql = "INSERT INTO rep67 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
'            StrSql = StrSql & "ternro,fechaalta,fechabaja,valor,nrodoc2) VALUES ("
'            StrSql = StrSql & bpronro & ","
'            StrSql = StrSql & Nroliq & ","
'            If Not Todos_Pro Then
'                StrSql = StrSql & rs_Procesos!pronro & ","
'                StrSql = StrSql & rs_Procesos!procaprob & ","
'            Else
'                StrSql = StrSql & "0" & ","
'                StrSql = StrSql & CInt(Proc_Aprob) & ","
'            End If
'            StrSql = StrSql & Empresa & ","
'            StrSql = StrSql & "'" & IdUser & "',"
'            StrSql = StrSql & ConvFecha(Fecha) & ","
'            StrSql = StrSql & "'" & Hora & "',"
'            StrSql = StrSql & rs_Empleados!ternro & ","
'            StrSql = StrSql & ConvFecha(fec_ini) & ","
'            StrSql = StrSql & ConvFecha(fec_fin) & ","
'            StrSql = StrSql & "0,' ')"
'            objConn.Execute StrSql, , adExecuteNoRecords
'        End If
'
'
'        Do While Not rs_Confrep.EOF
'            Select Case UCase(rs_Confrep!conftipo)
'            Case "AC": 'ACUMULADORES
'                StrSql = "SELECT * FROM aculiq " & _
'                         " INNER JOIN cabliq ON aculiq.cliqnro = " & rs_Empleados!cliqnro & _
'                         " WHERE aculiq.acunro =" & Par
'                OpenRecordset StrSql, rs_aculiq
'
'                If Not rs_aculiq.EOF Then
'                    'Actualizo
'                    StrSql = "UPDATE rep67 SET valor = valor + " & rs_aculiq!almonto
'                    StrSql = StrSql & " WHERE bpronro = " & bpronro
'                    StrSql = StrSql & " AND pliqnro = " & Nroliq
'                    StrSql = StrSql & " AND empresa = " & Empresa
'                    StrSql = StrSql & " AND pronro = " & NroProc
'                    StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'                    objConn.Execute StrSql, , adExecuteNoRecords
'
'
'                    StrSql = " SELECT cuil.nrodoc FROM tercero " & _
'                             " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) " & _
'                             " WHERE tercero.ternro= " & rs_Empleados!ternro
'                    OpenRecordset StrSql, rs_Cuil
'                    If Not rs_Cuil.EOF Then
'                        Aux_CUIL = "CUIL " & Left(CStr(rs_Cuil!nrodoc), 13)
'                        Aux_CUIL = Replace(CStr(Aux_CUIL), "-", "")
'
'                        StrSql = "UPDATE rep67 SET nrodoc2 = " & Aux_CUIL
'                        StrSql = StrSql & " WHERE bpronro = " & bpronro
'                        StrSql = StrSql & " AND pliqnro = " & Nroliq
'                        StrSql = StrSql & " AND empresa = " & Empresa
'                        StrSql = StrSql & " AND pronro = " & NroProc
'                        StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'                        objConn.Execute StrSql, , adExecuteNoRecords
'                    End If
'
'                    If TipoReporte <> 1 Then
'                        StrSql = " SELECT cuil.nrodoc FROM tercero " & _
'                                 " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro <= 4 AND cuail.nrodoc <> "") " & _
'                                 " WHERE tercero.ternro= " & rs_Empleados!ternro
'                        OpenRecordset StrSql, rs_Cuil
'                        If Not rs_Cuil.EOF Then
'                            Aux_CUIL = "CUIL " & Left(CStr(rs_Cuil!nrodoc), 13)
'                            Aux_CUIL = Replace(CStr(Aux_CUIL), "-", "")
'
'                            StrSql = "UPDATE rep67 SET nrodoc2 = " & Aux_CUIL
'                            StrSql = StrSql & " WHERE bpronro = " & bpronro
'                            StrSql = StrSql & " AND pliqnro = " & Nroliq
'                            StrSql = StrSql & " AND empresa = " & Empresa
'                            StrSql = StrSql & " AND pronro = " & NroProc
'                            StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'                            objConn.Execute StrSql, , adExecuteNoRecords
'                        End If
'                    End If
'                End If
'
'            Case "CO": 'CONCEPTOS
'                StrSql = "SELECT * FROM detliq " & _
'                         " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
'                         " WHERE concepto.concpuente = -1 " & _
'                         " AND concepto.tconnro =" & rs_Confrep!confval
'                OpenRecordset StrSql, rs_Detliq
'
'                If Not rs_Detliq.EOF Then
'                    'Actualizo
'                    StrSql = "UPDATE rep67 SET total_liquidado = total_liquidado + " & rs_Detliq!dlimonto
'                    StrSql = StrSql & ", cant_liquidado = cant_liquidado + " & rs_Detliq!dlicant
'                    StrSql = StrSql & ", emp_liquidado = emp_liquidado + 1 "
'                    StrSql = StrSql & " WHERE concnro = " & rs_Detliq!concnro
'                    StrSql = StrSql & " AND agrupacion = " & selector
'                    StrSql = StrSql & " AND agrupacion2 = " & selector2
'                    StrSql = StrSql & " AND agrupacion3 = " & selector3
'                    StrSql = StrSql & " AND bpronro = " & bpronro
'                    StrSql = StrSql & " AND pliqnro = " & Nroliq
'                    StrSql = StrSql & " AND empresa = " & Empresa
'                    StrSql = StrSql & " AND pronro = " & NroProc
'                    StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
'                    objConn.Execute StrSql, , adExecuteNoRecords
'                End If
'
'            End Select
'
'                'Actualizo el progreso del Proceso
'                Progreso = Progreso + IncPorc
'                TiempoAcumulado = GetTickCount
'                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
'                         ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
'                         "' WHERE bpronro = " & NroProcesoBatch
'                objConn.Execute StrSql, , adExecuteNoRecords
'
'            'Siguiente confrep
'            rs_Confrep.MoveNext
'        Loop
'
'        'Siguiente empleado
'        rs_Empleados.MoveNext
'    Loop
'
'    'Siguiente proceso
'    rs_Procesos.MoveNext
'Loop
'
''Fin de la transaccion
'MyCommitTrans
'
'
'If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
'If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
'If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
'If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
'If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
'If rs_rep67.State = adStateOpen Then rs_rep67.Close
'If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
'
'Set rs_Empleados = Nothing
'Set rs_Procesos = Nothing
'Set rs_Confrep = Nothing
'Set rs_Detliq = Nothing
'Set rs_rep67 = Nothing
'Set rs_Periodo = Nothing
'Set rs_Reporte = Nothing
'
'
'Exit Sub
'CE:
'    HuboError = True
'    MyRollbackTrans
'
'End Sub


