Attribute VB_Name = "MdlRHProToPayroll"
Option Explicit

'Const Version = "1.0"
'Const FechaVersion = "14/04/2015" ' Sebastian Stremel - CAS-28350 - Salto Grande - Custom ADP - Web Service Adp/Payrrol

'Const Version = "1.1"
'Const FechaVersion = "04/06/2015" ' Sebastian Stremel - Modificacion en campos y update - CAS-28350 - Salto Grande - Custom ADP - Web Service Adp/Payrrol


'Const Version = "1.2"
'Const FechaVersion = "14/01/2016"  ' Sebastian Stremel - Modificacion en campos y update - CAS-28350 - Salto Grande - Custom ADP - Web Service Adp/Payrrol [Entrega 3]
                                   ' Cambios en las columnas de la tabla famila que no aceptan valores nulos
                                   ' Se actualizan los valores nulos
                                   ' Modificacion en categoria  categoria subrogante
                                   ' correccion cuando busca si existe el familiar
                                   ' Se busca el familiar por numero de renglon


'Const Version = "1.3"
'Const FechaVersion = "05/02/2016"  ' Sebastian Stremel - Modificacion en asig familiar, salario fam, estruc subrogante, caja jub, se agrega campo funflghc - CAS-28350 - Salto Grande - Custom ADP - Web Service Adp/Payrrol [Entrega 4]

Const Version = "1.4"
Const FechaVersion = "10/04/2016"  ' Sebastian Stremel - Los siguientes campos no se modifican en el update (uniestId2,LugpagIdNr,FunNroCtab) - CAS-28350 - Salto Grande - Custom ADP - Web Service Adp/Payrrol [Entrega 5]


Dim rs_datos As New ADODB.Recordset
'Dim dirsalidas As String
Dim usuario As String
Dim Incompleto As Boolean
Dim objconnInsert As New ADODB.Connection
'variables del confrep
Dim tipoDoc As Integer
Dim TeUnidadGerencia As Integer
Dim TeCargo As Integer
Dim TeCategoria As Integer
Dim TeCategoria2 As Integer
Dim TeMutual As Integer
Dim TeAmbienteContable As Integer
Dim TeCentroCosto As Integer
Dim TeLugarFisico As Integer
Dim TeGrupoEmpleado As Integer
Dim TeLugarPago As Integer
Dim TeBanco As Integer
Dim TeRelFuncional As Integer
Dim TeIdFuncion As Integer
Dim TeCategoriaSub As Integer
Dim TeLugarFisicoTrabajo As Integer
Dim TeRegimenHorario As Integer
Dim TeInstitucionAporte As Integer
Dim teContrato As Integer
Dim TDMutual As Integer
Dim TipoLicMaternidad As Integer
Dim TeEmeM As Integer
Dim TeFunId2Cat As Integer
Dim TeFunIdMon As Integer


Dim strconexionAux As String
Dim nroConexion As Integer
Dim empCantHoras As String
Dim modeloInterfaz

Dim dirsalidas As String
Global cn_externa  As New ADODB.Connection

Public Sub Main()

Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
'Dim cn_externa  As New ADODB.Connection
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

Dim param
Dim ternroAux As Integer


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

    Nombre_Arch = PathFLog & "Interfaz_payroll_" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    
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
    
    On Error GoTo ME_Main
    
    StrSql = " SELECT * FROM confrepAdv "
    StrSql = StrSql & " WHERE repnro= 474 and confnrocol=22 "
    StrSql = StrSql & " ORDER BY confnrocol ASC"
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        nroConexion = IIf(EsNulo(rs_datos!confval), 0, rs_datos!confval)
    End If
    rs_datos.Close
    
    'busco los datos de la BD
    StrSql = "SELECT cnstring FROM conexion "
    StrSql = StrSql & " WHERE cnnro=" & nroConexion
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        strconexionAux = rs_datos("cnstring")
    Else
        strconexionAux = ""
        Flog.writeline "la conexion a payroll no esta configurada se aborta el proceso."
        Exit Sub
    End If
    rs_datos.Close
    
    'abro la conexion a la base de payrrol
    OpenConnection strconexionAux, objconnInsert
    
    cn_externa.ConnectionString = strconexionAux
    cn_externa.Open
    
    'CONTROLO EL VALOR DE LA TABLA FILQIN DE PAYROLL QUE INDICA SI PUEDO PROCESAR EN ESTE MOMENTO
    StrSql = "SELECT PerFlgMae FROM Filqin "
    'StrSql = "SELECT PerFlgMae FROM Filqion "
    OpenRecordsetExt StrSql, rs_datos, cn_externa
    If Not rs_datos.EOF Then
        If rs_datos!perflgmae = "N" Then
            Flog.writeline " El proceso queda en pendiente ya que la tabla Filqin tiene valor N "
            'Flog.writeline "Acutaliza el estado a pendiente"
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 100 ,bprcestado = 'Procesado', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            StrSql = "DELETE FROM batch_proceso WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Exit Sub
        End If
    End If
    rs_datos.Close
    'HASTA ACA
    
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Acutaliza el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 443 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = IIf(EsNulo(rs_batch_proceso!bprcparam), "", rs_batch_proceso!bprcparam)
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Flog.writeline "Levanto los datos del confrep"
        StrSql = " SELECT * FROM confrepAdv "
        StrSql = StrSql & " WHERE repnro= 474 "
        StrSql = StrSql & " ORDER BY confnrocol ASC"
        OpenRecordset StrSql, rs_datos
        If Not rs_datos.EOF Then
            Do While Not rs_datos.EOF
                Select Case rs_datos("confnrocol")
                    Case 1:
                        tipoDoc = rs_datos("confval")
                        Flog.writeline "tipo de documento configurado:" & tipoDoc
                    Case 2:
                        TeUnidadGerencia = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeUnidadGerencia:" & TeUnidadGerencia
                    Case 3:
                        TeCargo = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeCargo:" & TeCargo
                    Case 4:
                        TeCategoria = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        TeCategoria2 = IIf(EsNulo(rs_datos("confval2")), 0, rs_datos("confval2"))
                        Flog.writeline "TeCategoria:" & TeCategoria & " --- " & TeCategoria2
                    Case 5:
                        TeMutual = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeMutual:" & TeMutual
                    Case 6:
                        TeAmbienteContable = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeAmbienteContable:" & TeAmbienteContable
                    Case 7:
                        TeCentroCosto = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeCentroCosto:" & TeCentroCosto
                    Case 8:
                        TeLugarFisicoTrabajo = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeLugarFisicoTrabajo:" & TeLugarFisicoTrabajo
                    Case 9:
                        TeGrupoEmpleado = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeGrupoEmpleado:" & TeGrupoEmpleado
                    Case 10:
                        TeLugarPago = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeLugarPago:" & TeLugarPago
                    Case 11:
                        TeBanco = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeBanco:" & TeBanco
                    Case 12:
                        TeRelFuncional = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeRelFuncional:" & TeRelFuncional
                    Case 13:
                        TeIdFuncion = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeIdFuncion:" & TeIdFuncion
                    Case 14:
                        TeCategoriaSub = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeCategoriaSub:" & TeCategoriaSub
                    Case 15:
                        TeLugarFisico = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeLugarFisico:" & TeLugarFisico
                    Case 16:
                        TeRegimenHorario = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeRegimenHorario:" & TeRegimenHorario
                    Case 17:
                        TeInstitucionAporte = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeInstitucionAporte:" & TeInstitucionAporte
                    Case 18:
                        teContrato = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "TeContrato:" & teContrato
                    Case 19:
                        empCantHoras = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "Cantidad de horas diarias:" & empCantHoras
                    Case 20:
                        TDMutual = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "Tipo de documento de mutual:" & TDMutual
                    Case 21:
                        TipoLicMaternidad = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "Tipo de licencia maternidad:" & TipoLicMaternidad
                    Case 22:
                        nroConexion = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                    Case 23:
                        TeEmeM = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "Tipo de estructura Emergencia Med:" & TeEmeM
                    Case 24:
                        TeFunId2Cat = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "Tipo de estructura FunId2Cat:" & TeFunId2Cat
                    Case 25:
                        TeFunIdMon = IIf(EsNulo(rs_datos("confval")), 0, rs_datos("confval"))
                        Flog.writeline "Tipo de estructura FunIdMon:" & TeFunIdMon
                End Select
            rs_datos.MoveNext
            Loop
        End If
        rs_datos.Close
        
        'verifico los parametros, si tiene 2 parametros viene de un trigger
        param = Split(bprcparam, "@")
        If UBound(param) > 0 Then
            modeloInterfaz = param(0)
            ternroAux = param(1)
        Else
            modeloInterfaz = bprcparam
            ternroAux = 0
        End If
        
        
        'busco los datos de la BD
        StrSql = "SELECT cnstring FROM conexion "
        StrSql = StrSql & " WHERE cnnro=" & nroConexion
        OpenRecordset StrSql, rs_datos
        If Not rs_datos.EOF Then
            strconexionAux = rs_datos("cnstring")
        Else
            strconexionAux = ""
            Flog.writeline "la conexion a payroll no esta configurada se aborta el proceso."
            Exit Sub
        End If
        rs_datos.Close
        
        'abro la conexion a la base de payrrol
        OpenConnection strconexionAux, objconnInsert
        
        Call borrarEmpleados(ternroAux)
        Select Case modeloInterfaz
            Case -1: 'Familiares y Empleados
                Call interfazADP(ternroAux)
                Call interfazFamiliares(ternroAux)
            
            Case 0: 'Empleados
                Call interfazADP(ternroAux)
                
            Case 1: 'Familiares
                Call interfazFamiliares(ternroAux)
        End Select
        
        
    Else
        Flog.writeline "No se encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline "--------------------------------------------------------------"
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Else
        If Incompleto Then
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado', bprcprogreso=100 WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
            'Borro de batch_proceso
            StrSql = "DELETE FROM batch_proceso WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    
Fin:
    Flog.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline
        
    'Actualizo el progreso
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub
Public Sub OpenRecordsetExt(strSQLQuery As String, ByRef objRs As ADODB.Recordset, ByVal objConnE As ADODB.Connection, Optional lockType As LockTypeEnum = adLockReadOnly)
' ---------------------------------------------------------------------------------------------
' Descripcion: Abre recordset de conexion externa
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim pos1 As Long
Dim pos2 As Long
Dim aux As String

    'Abre un recordset con la consulta strSQLQuery
    If objRs.State <> adStateClosed Then
        If objRs.lockType <> adLockReadOnly Then objRs.UpdateBatch
        objRs.Close
    End If
    
    objRs.CacheSize = 500

    objRs.Open strSQLQuery, objConnE, adOpenDynamic, lockType, adCmdText
    
    Cantidad_de_OpenRecordset = Cantidad_de_OpenRecordset + 1
    
End Sub
Public Sub interfazADP(ByVal ternro As Integer)

'-------------------------------------------------
'ternro = 0 todos los empleados
'ternro <> 0  un empleado en particular
'-------------------------------------------------
Dim actualizar As Integer
Dim listaEmpleados As String
Dim empleado
Dim j As Integer
Dim Objeto As New datosPersonales

'variables para el insert
Dim empnrodoc As String
Dim empnrodocAux As String
Dim empNombreAbreviado As String
Dim empEstado As String
Dim empMonedaCobro As Integer
Dim empNombre As String
Dim empSegundoNombre As String
Dim empApellido As String
Dim empSegundoApellido As String
Dim empUnidad As Integer
Dim empCargo As Integer
Dim empDescCargo As String

Dim empLegajo As Long
Dim empSexo As String
Dim empOficina As String
Dim empReporta As String
Dim empFNacimiento As String
Dim empLugarPago As String
Dim empCategoria As String

Dim empFechaBaja As String
Dim empFechaAlta As String
Dim empEstadoCivil As String
Dim empSector As String
Dim empMoneda As String
Dim empPais As String
Dim empGrupo As String
Dim empDescPais As String
Dim empCausaBaja As String
Dim empFechaReingreso As String
Dim empCuil As String
Dim empTurno As String
Dim empLugarNac As String
Dim empTitulo As String

Dim empCedulaIdentidad As String
Dim empEdad As Integer
Dim empLugarFisico As String
Dim empInstAporte As String


Dim empNroAfiliadoMutual As String
Dim empSalarioFamiliar As String
Dim empMutual As String
Dim empFechaEstadoCivil As String
Dim empAmbienteContable As String
Dim empCentroCosto As String
Dim empBanco As String
Dim empNroCuenta As String
Dim empTipoCuenta As String
Dim empIdFuncion As String
Dim empCategoriaSub As String
Dim empNacionalidad As String
Dim empEmail As String
Dim empRelFuncional As String
Dim empLugarFisicoTrabajo As String
Dim empInstitucion As String
Dim empRegimenHorario As String
Dim empTipoDoc As String
Dim empTelInterno As String

Dim empContrato As String

Dim EmpFamiFlg1Co As String
Dim empMaternidad As String

Dim cantEmpleados As Long
Dim HayError As Boolean

Dim aux As String
Dim aux1 As String

Dim empFunId2Cat As String

Dim listaUpdate As String
Dim empresa As String
'empresa = "CTM"
Dim str_error As String     ' arma la tabla con errores que se enviara por mail
ReDim arr_Errores(71) As String
listaUpdate = 0
listaEmpleados = 0

Dim hubo_error As Boolean
Dim hubo_warning As Boolean


Dim grupoLiq
Dim ok As Boolean
Dim empNroContr As String
Dim empVencCont

Dim FunFlgHC

Dim I
'-------------------------OBTENGO LOS TERCEROS A SINCRONIZAR--------------------------
If ternro = 0 Then
    StrSql = " SELECT distinct ternro,tipo,tipnro FROM empsincPayroll "
    StrSql = StrSql & " WHERE tipnro in(1,2) AND sinc=0 "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        Do While Not rs_datos.EOF
            listaEmpleados = listaEmpleados & "," & rs_datos("ternro")
        rs_datos.MoveNext
        Loop
    End If
    rs_datos.Close
Else
    listaEmpleados = listaEmpleados & "," & ternro
End If

'-------------------------------------------------------------------------------------
empleado = Split(listaEmpleados, ",")
If UBound(empleado) < 1 Then
    Flog.writeline "No hay empleados para procesar. Se aborta el proceso."
    Exit Sub
End If

    If modeloInterfaz = -1 Then
        cantEmpleados = UBound(empleado)
        If cantEmpleados > 0 Then
            IncPorc = 50 / cantEmpleados
        Else
            IncPorc = 50
        End If
    Else
        cantEmpleados = UBound(empleado)
        If cantEmpleados > 0 Then
            IncPorc = 99 / cantEmpleados
        Else
            IncPorc = 99
        End If
    End If
    Dim arr_i
    Dim datos
    
    For j = 1 To UBound(empleado)

        ok = True
        For arr_i = 0 To UBound(arr_Errores)
            arr_Errores(arr_i) = ""
        Next
        hubo_error = False
        hubo_warning = False
        arr_Errores(0) = empleado(j)
        
        Objeto.buscarDatosPersonales (empleado(j))
        Flog.writeline "-------Comienza busqueda datos del empleado-------"
        Flog.writeline "Numero de tercero a buscar: " & empleado(j)
        
        '----------------------------DNI---------------------------------
        Objeto.buscarNroDoc empleado(j), tipoDoc
        Call checkError(Objeto.obtenerNroDoc, True, 0, "EmpDni", arr_Errores, 1, "INT", hubo_error, hubo_warning, empnrodoc)
        Flog.writeline "DNI: " & empnrodoc
        '---------------------------------------------------------------------------------
        
        
        '-------------------------Estado del empleado-------------------------------------
        Objeto.buscarEstado empleado(j)
        Call checkError(Objeto.obtenerEstado, False, 1, "FunStsEsta", arr_Errores, 3, "STRING", hubo_error, hubo_warning, empEstado)
        Flog.writeline "Estado: " & Objeto.obtenerEstado
        '---------------------------------------------------------------------------------

        
        '---------------------------Moneda de cobro---------------------------------------
        Objeto.buscarEstructuras empleado(j), TeFunIdMon, Date, "estrcodext", False
        Call checkError(Objeto.obtenerEstructura, False, 0, "FunIdMon", arr_Errores, 4, "INT", hubo_error, hubo_warning, empMonedaCobro)
        Flog.writeline "Moneda de cobro: " & Objeto.obtenerMoneda
        '---------------------------------------------------------------------------------
        
        '----------------------------Unidad Gerencia--------------------------------------
        'POR AHORA NO SE INFORMA YA QUE NO SABEN DE DONDE SALIO
        'Objeto.buscarEstructuras Empleado(j), TeUnidadGerencia, Date, "estrcodext"
        'Call checkError(Objeto.obtenerEstructura, False, 0, "UniestId2", arr_Errores, 5, "INT", hubo_error, hubo_warning, empUnidad)
        'Flog.writeline "Unidad de gerencia: " & Objeto.obtenerEstructura
        empUnidad = "1"
        '---------------------------------------------------------------------------------
        
        '----------------------------Legajo del empleado----------------------------------
        Call checkError(Objeto.obtenerLegajo, False, 0, "FunNroTarj", arr_Errores, 25, "INT", hubo_error, hubo_warning, empLegajo)
        Flog.writeline "Legajo del empleado: " & Objeto.obtenerLegajo
        '---------------------------------------------------------------------------------
        
        '----------------------------Nombre del empleado----------------------------------
        Call checkError(Objeto.obtenerNombreApellido("nombre"), False, 20, "FunNom1", arr_Errores, 6, "CHAR", hubo_error, hubo_warning, empNombre)
        Flog.writeline "Nombre del empleado: " & Objeto.obtenerNombreApellido("nombre")
        '---------------------------------------------------------------------------------
        
        '-----------------------------Segundo Nombre--------------------------------------
        Call checkError(Objeto.obtenerNombreApellido("nombre2"), False, 20, "FunNom2", arr_Errores, 7, "CHAR", hubo_error, hubo_warning, empSegundoNombre)
        Flog.writeline "Segundo Nombre del empleado: " & Objeto.obtenerNombreApellido("nombre2")
        '---------------------------------------------------------------------------------
        
        '-------------------------------Apellido------------------------------------------
        Call checkError(Objeto.obtenerNombreApellido("apellido"), False, 20, "FunApe1", arr_Errores, 8, "CHAR", hubo_error, hubo_warning, empApellido)
        Flog.writeline "Apellido del empleado: " & Objeto.obtenerNombreApellido("apellido")
        '---------------------------------------------------------------------------------
        
        '-----------------------------Segundo apellido------------------------------------
        Call checkError(Objeto.obtenerNombreApellido("apellido2"), False, 20, "FunApe2", arr_Errores, 9, "CHAR", hubo_error, hubo_warning, empSegundoApellido)
        Flog.writeline "Segundo Apellido del empleado: " & Left(empSegundoApellido, 20)
        '---------------------------------------------------------------------------------
        
        '-----------------------------Nombre Abreviado------------------------------------
        empNombreAbreviado = Left(empApellido & "," & empNombre & " " & empSegundoNombre, 24)
        Flog.writeline "Nombre Abreviado : " & empNombreAbreviado
        '---------------------------------------------------------------------------------
        
        '---------------------------------cargo-------------------------------------------
        Objeto.buscarEstructuras empleado(j), TeCargo, Date, "estrcodext", False
        Call checkError(Objeto.obtenerEstructura, False, 0, "FunIdCodca", arr_Errores, 10, "INT", hubo_error, hubo_warning, empCargo)
        Flog.writeline "El cargo del empleado es: " & Objeto.obtenerEstructura
        '---------------------------------------------------------------------------------
        
        '--------------------------------categoria----------------------------------------
        
        'primero busco si tiene categoria subrogante, sino le guardo la categoria
        Objeto.buscarEstructuras empleado(j), TeCategoriaSub, Date, "estrdabr"
        Call checkError(Objeto.obtenerEstructura, False, 5, "FunIdCat", arr_Errores, 35, "CHAR", hubo_error, hubo_warning, empCategoria)
        Flog.writeline "FUNIDCAT: " & Objeto.obtenerEstructura
        If EsNulo(Objeto.obtenerEstructura) Then
            Objeto.buscarEstructuras empleado(j), TeCategoria, Date, "estrdext"
            aux = Objeto.obtenerEstructura
            
            Objeto.buscarEstructuras empleado(j), TeCategoria2, Date, "estrcodext"
            aux1 = Objeto.obtenerEstructura
            empCategoria = aux & "-" & aux1
            Flog.writeline "La categoria del empleado es: " & empCategoria
        Else
            empCategoria = Objeto.obtenerEstructura
        End If

        '---------------------------------------------------------------------------------
        
        '------------------------------Estado Civil---------------------------------------
        Call checkError(Objeto.obtenerEstadoCivil, False, 1, "FunStsCivi", arr_Errores, 12, "CHAR", hubo_error, hubo_warning, empEstadoCivil)
        Flog.writeline "Estado Civil del empleado es: " & Objeto.obtenerEstadoCivil
        '---------------------------------------------------------------------------------
        
        '-----------------------------Cobra salario familiar------------------------------
        Objeto.buscarAsignacionesFam (empleado(j))
        Call checkError(Objeto.obtenerAsignacionFam, False, 1, "FamiFlgSal", arr_Errores, 13, "CHAR", hubo_error, hubo_warning, empSalarioFamiliar)
        Flog.writeline "Cobra salario familiar: " & empSalarioFamiliar
        
        If empSalarioFamiliar = "S" Then  'N=0
            FunFlgHC = "F"
        Else
            FunFlgHC = "N"
        End If
        Flog.writeline "FunFlgHC: " & FunFlgHC
        
        '-----------------------------------Mutual----------------------------------------
        'Objeto.buscarEstructuras Empleado(j), TeMutual, Date, "estrcodext"
        'Call checkError(Objeto.obtenerEstructura, False, 0, "FunIdMut", arr_Errores, 14, "INT", hubo_error, hubo_warning, empMutual)
        'Flog.writeline "La mutual del empleado es: " & Objeto.obtenerEstructura
        empMutual = ""
        '---------------------------------------------------------------------------------
        
        '------------------------Nro de afiliado de la mutual-----------------------------
        'Objeto.buscarNroDoc Empleado(j), TDMutual
        'Call checkError(Objeto.obtenerNroDoc, False, 12, "FunNroMutu", arr_Errores, 15, "CHAR", hubo_error, hubo_warning, empNroAfiliadoMutual)
        'Flog.writeline "Nro de afiliado de mutual del empleado es: " & Objeto.obtenerNroDoc
        empNroAfiliadoMutual = ""
        
        '--------------Se informa vacio porque no es necesario para payroll---------------
        EmpFamiFlg1Co = ""
        '---------------------------------------------------------------------------------
        
        '------------------------------fecha de baja--------------------------------------
        Objeto.buscarFechaBaja empleado(j), Date
        Call checkError(Objeto.obtenerFechaBaja, False, 0, "FunFchBaja", arr_Errores, 16, "DATETIME", hubo_error, hubo_warning, empFechaBaja)
        Flog.writeline "Fecha de baja del empleado: " & Objeto.obtenerFechaBaja
        '---------------------------------------------------------------------------------
        
        '------------------------------fecha de alta--------------------------------------
        Objeto.buscarFechaAlta empleado(j), Date
        Call checkError(Objeto.obtenerFechaAlta, False, 0, "FunFchIngr", arr_Errores, 17, "DATETIME", hubo_error, hubo_warning, empFechaAlta)
        Flog.writeline "Fecha de alta del empleado: " & Objeto.obtenerFechaAlta
        '---------------------------------------------------------------------------------
        
        '----------------------------Fecha de reingreso-----------------------------------
        Objeto.buscarFechaReingreso empleado(j), Date
        Call checkError(Objeto.obtenerFechaReingreso, False, 0, "FunFchRein", arr_Errores, 18, "DATETIME", hubo_error, hubo_warning, empFechaReingreso)
        Flog.writeline "Fecha de reingreso: " & empFechaReingreso
        '---------------------------------------------------------------------------------
        
        '---------------------------Contrato del empleado---------------------------------
        Objeto.buscarEstructuras empleado(j), teContrato, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 1, "FunFlgCont", arr_Errores, 19, "CHAR", hubo_error, hubo_warning, empContrato)
        Flog.writeline "Contrato del empleado: " & Objeto.obtenerEstructura
        '---------------------------------------------------------------------------------
        
        '----------------------------Vencimiento del contrato-----------------------------
        Objeto.buscarFechaVencContr empleado(j), teContrato
        empVencCont = Objeto.obtenerFechaVencContr
        Flog.writeline "fecha de vencimiento del contrato: " & Objeto.obtenerFechaVencContr
        '---------------------------------------------------------------------------------
        
        '-----------------------------Nro de contrato del emp-----------------------------
        empNroContr = "99999"
        '---------------------------------------------------------------------------------
        
        '----------------------------Fecha de nacimiento----------------------------------
        Call checkError(Objeto.obtenerFNacimiento, False, 0, "FunFchNaci", arr_Errores, 20, "CHAR", hubo_error, hubo_warning, empFNacimiento)
        Flog.writeline "Fecha de nacimiento: " & empFNacimiento
        '---------------------------------------------------------------------------------
        
        '-----------------------------lugar de nacimiento---------------------------------
        Call checkError(Objeto.obtenerLugarNac, False, 20, "FunTxtLugn", arr_Errores, 21, "CHAR", hubo_error, hubo_warning, empLugarNac)
        Flog.writeline "Lugar de nacimiento: " & Objeto.obtenerLugarNac
        '---------------------------------------------------------------------------------
        
        '-----------------------------Sexo del empleado-----------------------------------
        empSexo = Left(Objeto.obtenerSexo, 1)
        Flog.writeline "Sexo del empleado: " & Left(empSexo, 1)
        '---------------------------------------------------------------------------------
        
        '---------------------Fecha de casamiento del empleado----------------------------
        Call checkError(Objeto.obtenerFechaEstadoCivil, False, 0, "FunFchCasa", arr_Errores, 22, "DATETIME", hubo_error, hubo_warning, empFechaEstadoCivil)
        Flog.writeline "Fecha de casamiento: " & Objeto.obtenerFechaEstadoCivil
        '---------------------------------------------------------------------------------
        
        '----------------------------Ambiente contable------------------------------------
        Objeto.buscarEstructuras empleado(j), TeAmbienteContable, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 0, "FunIdAc", arr_Errores, 23, "INT", hubo_error, hubo_warning, empAmbienteContable)
        Flog.writeline "Ambiente contable: " & Objeto.obtenerEstructura
        '---------------------------------------------------------------------------------
        
        '----------------------------Centro de costo--------------------------------------
        Objeto.buscarEstructuras empleado(j), TeCentroCosto, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 0, "FunIdCC", arr_Errores, 24, "INT", hubo_error, hubo_warning, empCentroCosto)
        Flog.writeline "Centro de costo: " & Objeto.obtenerEstructura
        '----------------------------------------------------------------------------------
        
        '--------------------------Lugar Fisico Trabajo------------------------------------
        Objeto.buscarEstructuras empleado(j), TeLugarFisicoTrabajo, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 0, "OfiperIdNr", arr_Errores, 25, "INT", hubo_error, hubo_warning, empLugarFisicoTrabajo)
        Flog.writeline "Lugar fisico de trabajo: " & empLugarFisicoTrabajo
        '----------------------------------------------------------------------------------
        
        '---------------------------grupo del empleado-------------------------------------
        Objeto.buscarEstructuras empleado(j), TeGrupoEmpleado, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 0, "GrupoIdPag", arr_Errores, 26, "CHAR", hubo_error, hubo_warning, empGrupo)
        grupoLiq = Split(empGrupo, "@")
        If UBound(grupoLiq) >= 1 Then
            empGrupo = grupoLiq(0)
            empLugarPago = grupoLiq(1)
        Else
            If UBound(grupoLiq) = 0 Then
                empGrupo = grupoLiq(0)
                empLugarPago = 99
            Else
                empGrupo = ""
                empLugarPago = 99
            End If
        End If
        'Flog.writeline "Grupo del empleado: " & Objeto.obtenerEstructura
        Flog.writeline "Grupo del empleado: " & empGrupo
        '----------------------------------------------------------------------------------
        
        '-----------------------------lugar de pago----------------------------------------
        'Objeto.buscarEstructuras Empleado(j), TeLugarPago, Date, "estrcodext"
        'Call checkError(Objeto.obtenerEstructura, False, 0, "LugpagIdNr", arr_Errores, 27, "INT", hubo_error, hubo_warning, empLugarPago)
        Flog.writeline "Lugar de pago: " & empLugarPago
        '----------------------------------------------------------------------------------
        
        '---------------------------------Emp banco----------------------------------------
        'Objeto.buscarEstructuras Empleado(j), TeBanco, Date, "estrcodext"
        'Call checkError(Objeto.obtenerEstructura, False, 0, "BanId", arr_Errores, 28, "INT", hubo_error, hubo_warning, empBanco)
        'Flog.writeline "Banco: " & Objeto.obtenerEstructura
        empBanco = "14"
        '----------------------------------------------------------------------------------
        
        '-------------------------------Emp nro cuenta-------------------------------------
        'Objeto.buscarCtaBancaria Empleado(j)
        'Call checkError(Objeto.obtenerCtaBancaria, False, 26, "FunNroCtab", arr_Errores, 29, "CHAR", hubo_error, hubo_warning, empNroCuenta)
        'Flog.writeline "Nro de cuenta: " & Objeto.obtenerCtaBancaria
        empNroCuenta = "99999"
        '----------------------------------------------------------------------------------
        
        '-------------------------------causa de baja--------------------------------------
        Call checkError(Objeto.obtenerCausaBajaEmpleado, False, 5, "Cauegrid", arr_Errores, 30, "CHAR", hubo_error, hubo_warning, empCausaBaja)
        Flog.writeline "Causa de baja: " & Objeto.obtenerCausaBajaEmpleado
        '----------------------------------------------------------------------------------
        
        '---------------------------Tipo de cta del banco----------------------------------
        'Call checkError(Objeto.obtenerTipoCta, False, 0, "TipctaId", arr_Errores, 31, "INT", hubo_error, hubo_warning, empTipoCuenta)
        'Flog.writeline "Tipo de cuenta del banco: " & empTipoCuenta
        empTipoCuenta = "9"
        '----------------------------------------------------------------------------------
        
        '---------------------------Esta de maternidad-------------------------------------
        Objeto.buscarLicencia empleado(j), TipoLicMaternidad
        empMaternidad = Objeto.obtenerLicencia
        If empMaternidad = "" Then
            empMaternidad = "N"
        Else
            empMaternidad = "S"
        End If
        '----------------------------------------------------------------------------------
        
        '---------------------------Email del empleado-------------------------------------
        Call checkError(Objeto.obtenerEmail, False, 40, "FunImail", arr_Errores, 32, "CHAR", hubo_error, hubo_warning, empEmail)
        Flog.writeline "Email del empleado: " & Objeto.obtenerEmail
        '----------------------------------------------------------------------------------
        
        '---------------------------Relacion funcional-------------------------------------
        Objeto.buscarEstructuras empleado(j), TeRelFuncional, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 0, "RelFunId", arr_Errores, 33, "INT", hubo_error, hubo_warning, empRelFuncional)
        Flog.writeline "Relacion funcional: " & Objeto.obtenerEstructura
        empRelFuncional = "99"
        '----------------------------------------------------------------------------------
        
        '--------------------------------Id funcion----------------------------------------
        Objeto.buscarEstructuras empleado(j), TeIdFuncion, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 0, "FuncionId", arr_Errores, 34, "INT", hubo_error, hubo_warning, empIdFuncion)
        Flog.writeline "Id funcion: " & Objeto.obtenerEstructura
        '----------------------------------------------------------------------------------
        
        '-------------------------Emp categoria subrogacion--------------------------------
        Objeto.buscarEstructuras empleado(j), TeCategoria, Date, "estrdext"
        aux = Objeto.obtenerEstructura
        
        Objeto.buscarEstructuras empleado(j), TeCategoria2, Date, "estrcodext"
        aux1 = Objeto.obtenerEstructura

        empCategoriaSub = aux & "-" & aux1
        Flog.writeline "FUNIDCAT: " & empCategoriaSub
        
        'Objeto.buscarEstructuras Empleado(j), TeCategoriaSub, Date, "estrdabr"
        'Call checkError(Objeto.obtenerEstructura, False, 5, "FunIdCat", arr_Errores, 35, "CHAR", hubo_error, hubo_warning, empCategoriaSub)
        'Flog.writeline "Categoria subrogacion: " & Objeto.obtenerEstructura
        '----------------------------------------------------------------------------------
         
        '-------------------------Emp FunId2Cat--------------------------------
        'Objeto.buscarEstructuras Empleado(j), TeFunId2Cat, Date, "estrdabr"
        'aux = Objeto.obtenerEstructura
        
        'Objeto.buscarEstructuras Empleado(j), TeFunId2Cat, Date, "estrcodext"
        'aux1 = Objeto.obtenerEstructura
        ''Call checkError(Objeto.obtenerEstructura, False, 5, "FunIdCat", arr_Errores, 35, "CHAR", hubo_error, hubo_warning, empCategoriaSub)
        'empFunId2Cat = aux & "-" & aux1
        'Flog.writeline "FUNID2CAT: " & empFunId2Cat
        '----------------------------------------------------------------------------------
         
         
        'fecha antiguedad = empFechaAlta 'es lo mismo
        
        '----------------------------Empleado nacionalidad---------------------------------
        Call checkError(Objeto.obtenerNacionalidad, False, 0, "NacionId", arr_Errores, 36, "INT", hubo_error, hubo_warning, empNacionalidad)
        Flog.writeline "Nacionalidad del empleado: " & Objeto.obtenerNacionalidad
        '----------------------------------------------------------------------------------
        
        '---------------------------Telefono Interno del emp-------------------------------
        Call checkError(Objeto.obtenerTelInterno, False, 0, "FunTelInte", arr_Errores, 37, "INT", hubo_error, hubo_warning, empTelInterno)
        Flog.writeline "Tel Interno del empleado: " & Objeto.obtenerTelInterno
        '----------------------------------------------------------------------------------
        
        '-------------------Profesion - se busca el titulo del empleado--------------------
        Objeto.buscarTitulo empleado(j)
        Call checkError(Objeto.obtenerTitulo, False, 0, "Titrprfid", arr_Errores, 38, "INT", hubo_error, hubo_warning, empTitulo)
        Flog.writeline "Titulo del empleado: " & Objeto.obtenerTitulo
        '----------------------------------------------------------------------------------
        
        
        '-------------------Institucion educativa del empleado-----------------------------
        'Call checkError(Objeto.obtenerInstitucion, False, 30, "FunTitInst", arr_Errores, 39, "CHAR", hubo_error, hubo_warning, empInstitucion)
        'Flog.writeline "Institucion Educativa: " & Objeto.obtenerInstitucion
        empInstitucion = ""
        '----------------------------------------------------------------------------------
        
        '----------------------------Lugar Fisico------------------------------------------
        Objeto.buscarEstructuras empleado(j), TeLugarFisico, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 30, "LugFisId", arr_Errores, 40, "INT", hubo_error, hubo_warning, empLugarFisico)
        Flog.writeline "Lugar fisico: " & Objeto.obtenerEstructura
        '----------------------------------------------------------------------------------
        
        '--------------------------------CUIL del empl-------------------------------------
        'CUIL del empleado - si es argentino sino otro tipo doc
        empPais = Objeto.obtenerPaisDelEmpleado
        Select Case empPais
            Case "1":
                'si es argentino busco el cuil
                empCuil = Left(Objeto.obtenerCUIL, 14)
            Case "25"
                'si es uruguayo nada
                empCuil = Left("", 14)
        End Select
        Call checkError(empCuil, False, 14, "FunNroCuil", arr_Errores, 41, "CHAR", hubo_error, hubo_warning, empCuil)
        Flog.writeline "Cuil del empleado: " & empCuil
        '----------------------------------------------------------------------------------
        
        '-----------------------------Regimen Horario--------------------------------------
        Objeto.buscarEstructuras empleado(j), TeRegimenHorario, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 0, "ReghId", arr_Errores, 42, "INT", hubo_error, hubo_warning, empRegimenHorario)
        Flog.writeline "Regimen Horario: " & Objeto.obtenerEstructura
        '----------------------------------------------------------------------------------
        
        
        '-----------------------------Tipo de documento del empleado-----------------------
        Call checkError(Objeto.obtenerTipoDoc, False, 3, "FunTipDoc", arr_Errores, 43, "CHAR", hubo_error, hubo_warning, empTipoDoc)
        Flog.writeline "Tipo de documento del empleado: " & Objeto.obtenerTipoDoc
        '----------------------------------------------------------------------------------
        
        '-----------------------------Inst Aporte------------------------------------------
        Objeto.buscarEstructuras empleado(j), TeInstitucionAporte, Date, "estrcodext"
        Call checkError(Objeto.obtenerEstructura, False, 0, "FunInstApo", arr_Errores, 44, "CHAR", hubo_error, hubo_warning, empInstAporte)
        Flog.writeline "Institucion Aporte: " & Objeto.obtenerEstructura
        'If Not EsNulo(Objeto.obtenerTipoDoc) Then
        '    empInstAporte = Objeto.obtenerTipoDoc
        'Else
        '    empInstAporte = "N"
        'End If
        '----------------------------------------------------------------------------------
    
        'Si hubo errores salto el empleado e informo
        If hubo_error Then
            Flog.writeline Espacios(Tabulador * 0) & "Se encontraron errores en el empleado: " & empleado(j)
            'Incompleto = True
        End If

            'INICIALIZO LA TRANSACCION
            objconnInsert.BeginTrans
            empnrodocAux = Replace(empnrodoc, "-", "")
            If IsNumeric(empnrodocAux) Then
                If existeEmpleadoEnPayroll(empnrodocAux, empLegajo, HayError) Then
                    'Lo actualizo, SOLO LOS CAMPOS ACTUALIZABLES
                    If Not HayError Then
                        If empCategoria = "99" Or empCategoria = "0" Then  'cambio pedido por esther 02/02/16 si la cat sub=99 poner el valor de la categoria
                            empCategoria = empCategoriaSub
                        End If
                        StrSql = " UPDATE FUNCIO "
                        StrSql = StrSql & " SET "
                        StrSql = StrSql & " FunStsEsta= '" & empEstado & "', "
                        StrSql = StrSql & " FunNomAbre= '" & empNombreAbreviado & "', "
                        StrSql = StrSql & " FunIdMon= " & empMonedaCobro & ", "
                        'StrSql = StrSql & " UniestId2= " & empUnidad & ", "
                        StrSql = StrSql & " FunNom1= '" & empNombre & "', "
                        StrSql = StrSql & " FunNom2= '" & empSegundoNombre & "' ,"
                        StrSql = StrSql & " FunApe1= '" & empApellido & "' ,"
                        StrSql = StrSql & " FunApe2= '" & empSegundoApellido & "' ,"
                        StrSql = StrSql & " FunIdCodca= '" & empCargo & "' ,"
                        StrSql = StrSql & " CategId= '" & empCategoria & "' ,"
                        StrSql = StrSql & " FunDocId= '" & empnrodoc & "' ,"
                        StrSql = StrSql & " FunStsCivi= '" & empEstadoCivil & "' ,"
                        StrSql = StrSql & " FamiFlgSal= '" & empSalarioFamiliar & "' , "
                        StrSql = StrSql & " FunFlgHC= '" & FunFlgHC & "' ," 'SS 05/02/2016
                        If Not EsNulo(empFechaBaja) Then
                            StrSql = StrSql & " FunFchBaja= " & ConvFecha(empFechaBaja) & ", "
                        Else
                            StrSql = StrSql & " FunFchBaja= NULL, "
                        End If
                        
                        If Not EsNulo(empFechaAlta) Then
                            StrSql = StrSql & " FunFchIngr= " & ConvFecha(empFechaAlta) & " ,"
                        Else
                            StrSql = StrSql & " FunFchIngr= NULL, "
                        End If
                        
                        If Not EsNulo(empFechaReingreso) Then
                            StrSql = StrSql & " FunFchRein= " & ConvFecha(empFechaReingreso) & " ,"
                        Else
                            StrSql = StrSql & " FunFchRein= NULL, "
                        End If
                        StrSql = StrSql & " FunFlgCont= '" & empContrato & "' , "
                        StrSql = StrSql & " FunFchNaci= " & ConvFecha(empFNacimiento) & ", "
                        StrSql = StrSql & " FunTxtLugn= '" & empLugarNac & "' ,"
                        If Not EsNulo(empFechaEstadoCivil) Then
                            StrSql = StrSql & " FunFchCasa= " & ConvFecha(empFechaEstadoCivil) & " , "
                        Else
                            StrSql = StrSql & " FunFchCasa= NULL, "
                        End If
                        StrSql = StrSql & " FunIdHorar= " & empCantHoras & ", "
                        StrSql = StrSql & " FunIdAc= " & empAmbienteContable & ", "
                        StrSql = StrSql & " FunIdCc= " & empCentroCosto & ", "
                        StrSql = StrSql & " OfiperIdNr= " & empLugarFisicoTrabajo & ", "
                        StrSql = StrSql & " GrupoIdPag= '" & empGrupo & "', "
                        'StrSql = StrSql & " LugpagIdNr= '" & empLugarPago & "', "
                        'StrSql = StrSql & " FunNroCtab= '" & empNroCuenta & "', "
                        StrSql = StrSql & " CauegrId= '" & empCausaBaja & "', "
                        StrSql = StrSql & " FunFlgMate= '" & empMaternidad & "', "
                        StrSql = StrSql & " FunImail= '" & empEmail & "', "
                        StrSql = StrSql & " FuncionId= " & empIdFuncion & ", "
                        StrSql = StrSql & " FunIdCat= '" & empCategoriaSub & "' , "
                        StrSql = StrSql & " NacionId= " & empNacionalidad & " , "
                        StrSql = StrSql & " TitprfId= " & empTitulo & ", "
                        StrSql = StrSql & " LugfisId= " & empLugarFisico & " ,"
                        StrSql = StrSql & " FunNroCuil= '" & empCuil & "'  ,"
                        StrSql = StrSql & " ReghId= " & empRegimenHorario & " , "
                        StrSql = StrSql & " FunTipDoc= '" & empTipoDoc & "' ,"
                        StrSql = StrSql & " FunInstApo = '" & empInstAporte & "'"
                        StrSql = StrSql & " WHERE FunIdNro=" & empnrodocAux
                        StrSql = StrSql & " AND HldId=1 AND EmpId=1 "
                        Flog.writeline "strsql update: " & StrSql
                        On Error Resume Next
                        objconnInsert.Execute StrSql, , adExecuteNoRecords
                        If Err <> 0 Then
                            Flog.writeline "ERROR PRODUCIDO: " & Err.Description
                            ok = False
                            GoTo siguiente
                        Else
                            Flog.writeline "Se updateo el empleado "
                            'actualizo la table empsincpayroll
                            StrSql = "UPDATE empsincPayroll set sinc=-1"
                            StrSql = StrSql & " WHERE ternro=" & empleado(j)
                            objConn.Execute StrSql, , adExecuteNoRecords
                            listaUpdate = listaUpdate & "," & empleado(j)
                        End If
                    End If
                Else
                    'INSERTO EN LAS TABLAS DE PAYROLL [FUNCIO]
                    If Not HayError Then
                        If empCategoria = "99" Or empCategoria = "0" Then  'cambio pedido por esther 02/02/16 si la cat sub=99 poner el valor de la categoria
                            empCategoria = empCategoriaSub
                        End If
                        StrSql = " INSERT INTO FUNCIO "
                        StrSql = StrSql & "("
                        StrSql = StrSql & " HldId, EmpId, FunIdNro, FunStsEsta, FunNomAbre, FunIdMon,  UniestId2, "
                        StrSql = StrSql & " FunNom1, FunNom2, FunApe1, FunApe2, FunIdCodca, "
                        StrSql = StrSql & " FunIdRemun, CategId,FunDocId, FunStsCivi, FamiFlgSal, "
                        StrSql = StrSql & " FunIdMut,"
                        StrSql = StrSql & " FunNroMutu, FamiFlg1Co, FunFchBaja, FunFch1Ing, FunFchIngr, FunFchRein, FunFlgCont, funfch1ven, funNroCont, "
                        StrSql = StrSql & " FunFchNaci, FunTxtLugn, FunIdSexo, FunFchCasa, FunIdHorar, FunIdAc, "
                        StrSql = StrSql & " FunIdCc, OfiperIdNr, GrupoIdPag, LugpagIdNr, FunNroTarj, "
                        StrSql = StrSql & " BanId, FunNroCtab, CauegrId, TipctaId, FunFlgMate, FunImail, "
                        StrSql = StrSql & " RelfunId, FuncionId, FunIdCat, FunfchAnti, NacionId, "
                        StrSql = StrSql & " TitprfId, FunTitInst, LugfisId, FunNroCuil, ReghId, "
                        StrSql = StrSql & " FunTipDoc, FunInstApo, FunId2Cat,FunPreNroP,FunFlgHC)"
                        StrSql = StrSql & " VALUES "
                        StrSql = StrSql & "("
                        StrSql = StrSql & " 1, 1, " & empnrodocAux & ",'" & empEstado & "','" & empNombreAbreviado & "'," & empMonedaCobro & "," & empUnidad & ","
                        StrSql = StrSql & " '" & empNombre & "','" & empSegundoNombre & "','" & empApellido & "','" & empSegundoApellido & "'," & empCargo & ","
                        StrSql = StrSql & " 'M', '" & empCategoria & "','" & empnrodoc & "','" & empEstadoCivil & "','" & empSalarioFamiliar & "', "
                        StrSql = StrSql & "0,"
                        StrSql = StrSql & "'" & empNroAfiliadoMutual & "','" & EmpFamiFlg1Co & "',"
                        If EsNulo(empFechaBaja) Then
                            StrSql = StrSql & " NULL,"
                        Else
                            StrSql = StrSql & ConvFecha(empFechaBaja) & ","
                        End If
                        
                        If EsNulo(empFechaAlta) Then
                            StrSql = StrSql & " NULL,"
                        Else
                            StrSql = StrSql & ConvFecha(empFechaAlta) & ","
                        End If
                        
                        If EsNulo(empFechaAlta) Then
                            StrSql = StrSql & " NULL,"
                        Else
                            StrSql = StrSql & ConvFecha(empFechaAlta) & ","
                        End If
                        
                        If EsNulo(empFechaReingreso) Then
                            StrSql = StrSql & " NULL,"
                        Else
                            StrSql = StrSql & ConvFecha(empFechaReingreso) & ","
                        End If
                        
                        StrSql = StrSql & "'" & empContrato & "',"
                        
                        If EsNulo(empVencCont) Then ' fecha de vencimiento del contrato
                            StrSql = StrSql & " NULL,"
                        Else
                            StrSql = StrSql & ConvFecha(empVencCont) & ","
                        End If
                        
                        StrSql = StrSql & "'" & empNroContr & "'," 'Nro de contrato
                        StrSql = StrSql & ConvFecha(empFNacimiento) & ",'" & empLugarNac & "','" & empSexo & "',"
                        If EsNulo(empFechaEstadoCivil) Then
                            StrSql = StrSql & " NULL,"
                        Else
                            StrSql = StrSql & ConvFecha(empFechaEstadoCivil) & ","
                        End If
                        StrSql = StrSql & empCantHoras & "," & empAmbienteContable & ", "
                        StrSql = StrSql & empCentroCosto & "," & empLugarFisicoTrabajo & ",'" & empGrupo & "'," & empLugarPago & "," & empLegajo & ","
                        StrSql = StrSql & empBanco & ",'" & empNroCuenta & "','" & empCausaBaja & "'," & empTipoCuenta & ",'" & empMaternidad & "','" & empEmail & "',"
                        StrSql = StrSql & empRelFuncional & "," & empIdFuncion & ",'" & empCategoriaSub & "',"
                        If EsNulo(empFechaAlta) Then
                            StrSql = StrSql & " NULL,"
                        Else
                            StrSql = StrSql & ConvFecha(empFechaAlta) & ","
                        End If
                        StrSql = StrSql & empNacionalidad & ","
                        StrSql = StrSql & empTitulo & ",'" & empInstitucion & "'," & empLugarFisico & ", '" & empCuil & "'," & empRegimenHorario & ","
                        StrSql = StrSql & "'" & empTipoDoc & "','" & empInstAporte & "','" & empFunId2Cat & "'"
                        StrSql = StrSql & "," & empLegajo
                        StrSql = StrSql & ", '" & FunFlgHC & "'"
                        StrSql = StrSql & ")"
                        Flog.writeline "-----------------INSERT---------------------------------"
                        Flog.writeline StrSql
                        Flog.writeline "-----------------FIN INSERT-----------------------------"
                        On Error Resume Next
                        objconnInsert.Execute StrSql, , adExecuteNoRecords
                        If Err <> 0 Then
                            Flog.writeline "Error producido en insert: " & Err.Description
                            ok = False
                            GoTo siguiente
                        Else
                            Flog.writeline "Se inserto el registro"
                            StrSql = "UPDATE empsincPayroll set sinc=-1"
                            StrSql = StrSql & " WHERE ternro=" & empleado(j)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
                'FINALIZO LA TRANSACCION
                objconnInsert.CommitTrans
                
                'hago update de los campos nulos
                Dim QUERY
                QUERY = ""
                StrSql = " SELECT * FROM funcio "
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " HldId=1 AND EmpId=1 AND FunIdNro=" & empnrodocAux
                OpenRecordsetExt StrSql, rs_datos, cn_externa
                If Not rs_datos.EOF Then
                    Do While Not rs_datos.EOF
                        For I = 0 To rs_datos.Fields.Count
                            If IsNull(rs_datos.Fields(I).Value) Then
                                Select Case rs_datos.Fields(I).Type
                                    Case 2: 'smallint
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=0"
                                    Case 3: 'int
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=0"
                                    Case 6: 'smallmoney
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=0"
                                    Case 7: 'datetime
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=" & ConvFecha("1753/01/01")
                                    Case 129: 'char
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=' '"
                                    Case 131: 'decimal
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=0"
                                    Case 135: 'smalldatetime
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=" & ConvFecha("1753/01/01")
                                End Select
                            End If
                        Next I
                        StrSql = "UPDATE funcio SET" & Replace(QUERY, ",", "", 1, 1) & " WHERE  HldId=1 AND EmpId=1 AND FunIdNro=" & empnrodocAux
                        objconnInsert.Execute StrSql, , adExecuteNoRecords
                        Flog.writeline "SE ACTUALIZARON LOS NULOS: " & StrSql
                    rs_datos.MoveNext
                    Loop
                End If
                'hasta aca
    
                
                
                If hubo_warning Then
                    'Incompleto = True
                    Flog.writeline "Se produjeron errores de tipo Warning"
                End If
            Else
                Flog.writeline "El empleado tiene un Nro de documento invalido"
            End If
        'End If

        'If hubo_warning Or hubo_error Then
        '    Call insertarError(NroProcesoBatch, arr_Errores, str_error, empresa, hubo_error, hubo_warning)
        'End If
        

        GoTo datosOK
            
'------------------cuando ocurre un error sigo con el proximo--------------
siguiente:
    If hubo_warning Or hubo_error Then
        Call insertarError(NroProcesoBatch, arr_Errores, str_error, empresa, hubo_error, hubo_warning)
        'Call crearProcesosMensajeria(str_error, empresa)
    End If
    Flog.writeline "Ocurrio un error"
    Flog.writeline "ULTIMA QUERY:" & StrSql
    Flog.writeline "descripcion del error: " & Err.Description
    ok = False
    Err.Clear
'--------------------------hasta aca---------------------------------------
    


datosOK:
    If hubo_warning Or hubo_error Then
        Call insertarError(NroProcesoBatch, arr_Errores, str_error, empresa, hubo_error, hubo_warning)

        'Call crearProcesosMensajeria(str_error, empresa)
    End If
    If ok Then
        Flog.writeline "datos ok"

    End If
    
        TiempoAcumulado = GetTickCount
        Progreso = Progreso + IncPorc
        Flog.writeline "Actualizo el progreso: " & Progreso
        StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "Se actualizo el progreso"
    
    Next
    
    str_error = str_error & "</table>" & vbCrLf
    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "
    Call crearProcesosMensajeria(str_error, empresa)
    

    
Flog.writeline "La interfaz Adp finalizo con exito"
Flog.writeline "Lista de terceros actualizados correctos:" & listaUpdate

Exit Sub
Fin:
    Flog.writeline "Ocurrio un error"
    Flog.writeline "ULTIMA QUERY:" & StrSql
    Flog.writeline "descripcion del error: " & Err.Description
    Flog.writeline "Lista de terceros actualizados correctos:" & listaUpdate
    
    Exit Sub
End Sub

Public Function esFamiliar(ByVal ternro As Integer) As Boolean
'funcion que indica si un tercero es familiar o no
    StrSql = "SELECT ternro FROM familiar "
    StrSql = StrSql & " WHERE familiar.ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        esFamiliar = True
    Else
        esFamiliar = False
    End If
    rs_datos.Close
End Function

'Public Function existeEnPayroll(ByVal FunIdNro As Long, ByVal empnrodoc As String, ByRef nroLinea As Integer) As Boolean
Public Function existeEnPayroll(ByVal FunIdNro As Long, ByVal renglon As Long) As Boolean
Dim cn_payroll As New ADODB.Connection


    StrSql = "SELECT cnstring FROM conexion "
    StrSql = StrSql & " WHERE cnnro=" & nroConexion
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        cn_payroll.ConnectionString = rs_datos!cnstring
    Else
        Exit Function
    End If
    rs_datos.Close
    cn_payroll.Open
    
    StrSql = "SELECT famidocid, FamiNroRen FROM famila "
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " HldId=1 AND EmpId=1"
    StrSql = StrSql & " AND famiNroRen = " & renglon
    StrSql = StrSql & " AND funidnro = " & Replace(FunIdNro, "-", "")
    'StrSql = StrSql & " AND replace(replace(famidocid,'-',''),'.','')='" & Replace(Replace(empnrodoc, "-", ""), ".", "") & "'"
    OpenRecordsetExt StrSql, rs_datos, cn_payroll
    Flog.writeline "consulta verifica existencia: " & StrSql
    If Not rs_datos.EOF Then
        existeEnPayroll = True
        'renglon = rs_datos!FamiNroRen
        Flog.writeline "El familiar existe en payroll: " & StrSql
    Else
        existeEnPayroll = False
        Flog.writeline "El familiar no existe en payroll"
    End If
    
    'busco el numero de la ultima linea
    'If Not existeEnPayroll Then
    '    StrSql = " SELECT max(FamiNroRen) renglon FROM famila " & _
    '             " WHERE HldId=1 AND EmpId=1 AND funidnro=" & FunIdNro
    '    OpenRecordsetExt StrSql, rs_datos, cn_payroll
    '    If Not rs_datos.EOF Then
    '        If EsNulo(rs_datos!renglon) Then
    '            nroLinea = 1
    '        Else
    '            nroLinea = CLng(rs_datos!renglon) + 1
    '        End If
    '    Else
    '        nroLinea = 1
    '    End If
    'End If
    'Flog.writeline "El familiar no existe en payroll"

    rs_datos.Close
End Function
Public Function existeEmpleadoEnPayroll(ByVal empnrodoc As Long, ByVal Legajo As Long, ByRef HayError As Boolean)
Dim cn_payroll As New ADODB.Connection

HayError = False
StrSql = "SELECT cnstring FROM conexion "
StrSql = StrSql & " WHERE cnnro=" & nroConexion
OpenRecordset StrSql, rs_datos
If Not rs_datos.EOF Then
    cn_payroll.ConnectionString = rs_datos!cnstring
Else
    Exit Function
End If
rs_datos.Close
cn_payroll.Open

StrSql = "SELECT FunIdNro FROM funcio "
StrSql = StrSql & " WHERE "
StrSql = StrSql & " HldId=1 AND EmpId=1"
StrSql = StrSql & " AND FunIdNro=" & empnrodoc
OpenRecordsetExt StrSql, rs_datos, cn_payroll
Flog.writeline "consulta existe empleado:" & StrSql
If Not rs_datos.EOF Then
    'existe con el documento
    'ahora verifico si es el mismo legajo
    StrSql = "SELECT FunIdNro FROM funcio "
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " HldId=1 AND EmpId=1"
    StrSql = StrSql & " AND FunIdNro=" & empnrodoc
    StrSql = StrSql & " AND FunNroTarj=" & Legajo
    OpenRecordsetExt StrSql, rs_datos, cn_payroll
    Flog.writeline "consulta existe empleado2:" & StrSql
    If Not rs_datos.EOF Then
        existeEmpleadoEnPayroll = True
        Flog.writeline "El empleado existe en payroll"
    Else
        existeEmpleadoEnPayroll = False
        Flog.writeline "El empleado existe con otro numero de legajo, no se procesa."
        HayError = True
    End If
Else
    existeEmpleadoEnPayroll = False
    Flog.writeline "El empleado no existe en payroll"
End If
rs_datos.Close
End Function
Public Sub interfazFamiliares(ByVal ternro As Integer)
    
'---------VARIABLES PARA LOS FAMILIARES---------
Dim FamHldId As Integer
Dim FamEmpId As Integer
Dim FunIdNro As String
Dim FamiNroRen As Integer
Dim FamiStsTip As String
Dim FamiNomCom As String
Dim FamiStsSex As String
Dim FamiFchNac As String
Dim FamiStsCon As String
Dim FamiFlgCob As String
Dim FamiStsEdu As String
Dim FamiFlgYac As String
Dim FamiFchAni As Integer
Dim FamiStsMut As String
Dim FamiIdMut As Integer
Dim FamiNroMut As String
Dim FamiFlgTra As String
Dim FamiCntEda As Integer
Dim FamiFlgCGI As String
Dim FamiFchFal As String
Dim FamiDocId As String
Dim FamiFlgHC As String
Dim FamiFlgEst As String
Dim FamiApe1 As String
Dim FamiApe2 As String
Dim FamiNom1 As String
Dim FamiNom2 As String
Dim FamiFlgBec As String
Dim FamiAnioAE As Integer
Dim FamiTipDoc As String
Dim FamiConMut As String
Dim FamiCobMut As String
Dim FamiIdEmeM As Integer
Dim FamiCobEme As String
'Dim FamiIdEmeM As String

Dim Objeto As New datosPersonales
Dim familiar As New Familiares

Dim nombreCompleto As String
'Dim ternro As Integer
Dim nroLinea As Integer
Dim rs_datos As New ADODB.Recordset
Dim connExt As New ADODB.Connection

connExt.ConnectionString = strconexionAux
Dim existe As Boolean
Dim j As Integer
Dim cantFamiliares As Long
Dim fam
Dim listaUpdate As String
ReDim arr_Errores(71) As String
Dim hubo_error As Boolean
Dim hubo_warning As Boolean

Dim empTernro As Integer

Dim str_error As String
Dim empresa As String

Dim I As Integer
I = 0
Dim renglon As Integer
renglon = 0
listaUpdate = 0
nroLinea = 0
'----------------HASTA ACA----------------------
    
Dim listaEmpleados
listaEmpleados = 0
'-------------------------OBTENGO LOS TERCEROS A SINCRONIZAR--------------------------
If ternro = 0 Then
    StrSql = "SELECT ternro,tipo,tipnro "
    StrSql = StrSql & " FROM empsincPayroll "
    StrSql = StrSql & " WHERE tipnro=3 and sinc=0 "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        Do While Not rs_datos.EOF
            listaEmpleados = listaEmpleados & "," & rs_datos("ternro")
        rs_datos.MoveNext
        Loop
    End If
    rs_datos.Close
Else
    listaEmpleados = listaEmpleados & "," & ternro
End If
    
    
fam = Split(listaEmpleados, ",")
cantFamiliares = UBound(fam)
    
If cantFamiliares = 0 Then
    Flog.writeline "No hay familiares para procesar"
    Exit Sub
End If
    

If modeloInterfaz = -1 Then
    If cantFamiliares > 0 Then
        IncPorc = 50 / cantFamiliares
    Else
        IncPorc = 50
    End If
Else
    If cantFamiliares > 0 Then
        IncPorc = 99 / cantFamiliares
    Else
        IncPorc = 99
    End If
End If

Dim arr_i
For j = 1 To cantFamiliares
    ternro = fam(j)
    Flog.writeline "Procesamiento del tercero familiar: " & ternro
    For arr_i = 0 To UBound(arr_Errores)
        arr_Errores(arr_i) = ""
    Next
    hubo_error = False
    hubo_warning = False
    arr_Errores(0) = ternro
    
    FamHldId = 1
    
    FamEmpId = 1
    
    'BUSCO EL TERNO DEL EMPLEADO AL CUAL CORRESPONDE EL FAMILIAR
    StrSql = " SELECT * FROM familiar "
    StrSql = StrSql & " WHERE ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        empTernro = rs_datos!empleado
        renglon = rs_datos!famnrocorr
        Flog.writeline " Tercero del empleado: " & empTernro
    Else
        Flog.writeline "No es familiar-----------"
    End If
    rs_datos.Close
    
    
    'Id del empleado----------------------------------------------
    Objeto.buscarDatosPersonales empTernro
    Objeto.buscarNroDoc empTernro, 0
    Call checkError(Objeto.obtenerNroDoc, True, 0, "FunIdNro", arr_Errores, 25, "INT", hubo_error, hubo_warning, FunIdNro)
    Flog.writeline "Documento del familiar: " & FunIdNro
    '-------------------------------------------------------------
    
    'Buscar datos del familiar------------------------------------
    Objeto.buscarDatosPersonales ternro
    'If IsNumeric(FunIdNro) Then
    '    existe = existeEnPayroll(CLng(FunIdNro), nroLinea)
    '    If nroLinea <> -1 Then
    '        FamiNroRen = nroLinea
    '    End If
        
        'busco si es Hijo, Padre, Conyuge, etc--------------------
        familiar.buscarRelacionFamiliar ternro
        Call checkError(familiar.obtenerRelacionFamiliar, False, 1, "FamiStsTip", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiStsTip)
        'FamiStsTip = Left(familiar.obtenerRelacionFamiliar, 1)
        Flog.writeline "Parentesco del familiar: " & FamiStsTip
        '---------------------------------------------------------
        
        
        'Call checkError(Objeto.obtenerNombreApellido("nombre"), False, 20, "FamiNom1", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, nombreCompleto)
        nombreCompleto = Objeto.obtenerNombreApellido("apellido")
        
        
        If Not EsNulo(Objeto.obtenerNombreApellido("nombre")) Then
            nombreCompleto = nombreCompleto & ", " & Objeto.obtenerNombreApellido("nombre")
        End If
        
        'nombreCompleto = nombreCompleto & " " & Objeto.obtenerNombreApellido("apellido")
        
        'If Not EsNulo(Objeto.obtenerNombreApellido("apellido2")) Then
        '    nombreCompleto = nombreCompleto & " " & Objeto.obtenerNombreApellido("apellido2")
        'End If
        
        FamiNomCom = Left(nombreCompleto, 30)
        Flog.writeline "Nombre Completo del empleado:" & Left(nombreCompleto, 30)
        
        'Sexo del familiar-------------------------------------------
        'FamiStsSex = Left(Objeto.obtenerSexo, 1)
        Call checkError(Objeto.obtenerSexo, False, 1, "FamiStsSex", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiStsSex)
        Flog.writeline "Sexo del familiar: " & FamiStsSex
        '------------------------------------------------------------
        
        'Fecha de nacimiento---------------------------------------------
        'FamiFchNac = Objeto.obtenerFNacimiento
        Call checkError(Objeto.obtenerFNacimiento, False, 0, "FamiFchNac", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiFchNac)
        Flog.writeline "Fecha de nacimiento del empleado: " & FamiFchNac
        '----------------------------------------------------------------
        
        
        'buscar condicion del familiar Normal - Minusvalido--------------
        'FamiStsCon = familiar.obtenerCondicion
        Call checkError(familiar.obtenerCondicion, False, 1, "FamiStsCon", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiStsCon)
        Flog.writeline "La condicion del familiar es: " & FamiStsCon
        '----------------------------------------------------------------
        
        'Cobra o no salario familiar-------------------------------------
        familiar.buscarSalarioFamiliar (ternro)
        Call checkError(familiar.obtenerSalarioFamiliar, False, 1, "FamiFlgCob", arr_Errores, 13, "CHAR", hubo_error, hubo_warning, FamiFlgCob)
        Flog.writeline "Cobra salario familiar: " & Objeto.obtenerSalarioFamiliar
        '----------------------------------------------------------------
        
        'Nivel de estudio------------------------------------------------
        familiar.buscarTitulo ternro
        Call checkError(familiar.obtenerNivel, False, 1, "FamiStsEdu", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiStsEdu)
        'FamiStsEdu = Objeto.obtenerNivel
        Flog.writeline "Nivel de estudio: " & FamiStsEdu
        '----------------------------------------------------------------
        
        FamiFlgYac = ""
        
        FamiFchAni = 0
        
        FamiStsMut = ""
        
        'Objeto.buscarEstructuras ternro, TeMutual, Date, "estrcodext"
        'FamiIdMut = Objeto.obtenerEstructura
        'Flog.writeline "Mutual del familiar: " & FamiIdMut
        
        'Objeto.buscarNroDoc ternro, 38
        'FamiNroMut = Objeto.obtenerNroDoc
        'Flog.writeline "Nro de afiliado del familiar: " & FamiNroMut
        
        'TRABAJA FAMILIAR---------------------------------------------
        'FamiFlgTra = familiar.obtenerTrabaja
        Call checkError(familiar.obtenerTrabaja, False, 1, "FamiFlgTra", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiFlgTra)
        Flog.writeline "Familiar trabaja: " & FamiFlgTra
        '-------------------------------------------------------------
        
        Dim edad
        Dim fnac
        fnac = Objeto.obtenerFNacimiento
        edad = DateDiff("YYYY", fnac, Date)
        FamiCntEda = edad
        Flog.writeline "Edad del familiar: " & FamiCntEda
        
        FamiFlgCGI = "N"
        
        'Nro de doc del familiar----------------------------------------
        Objeto.buscarNroDoc ternro, 0
        Call checkError(Objeto.obtenerNroDoc, False, 9, "FamiDocId", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiDocId)
        'FamiDocId = Objeto.obtenerNroDoc
        Flog.writeline "Nro de documento del familiar: " & FamiDocId
        '---------------------------------------------------------------
        
        'If IsNumeric(FamiDocId) Then
            'existe = existeEnPayroll(Replace(FunIdNro, "-", ""), FamiDocId, nroLinea)
            existe = existeEnPayroll(Replace(FunIdNro, "-", ""), renglon)
            FamiNroRen = renglon
            'If nroLinea <> -1 Then
            '    FamiNroRen = nroLinea
            'End If
        'Else
        '    existe = False
        'End If
        
        'Estudia el familiar--------------------------------------------
        'FamiFlgEst = familiar.obtenerEstudia
        'Call checkError(familiar.obtenerEstudia, False, 1, "FamiFlgEst", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiFlgEst)
        If familiar.obtenerNivel = "N" Then
            FamiFlgEst = "N"
        Else
            FamiFlgEst = "S"
        End If
        Flog.writeline "Estudia el familiar: " & FamiFlgEst
        '---------------------------------------------------------------
        
        'Nombre del empleado--------------------------------------------
        Call checkError(Objeto.obtenerNombreApellido("nombre"), False, 20, "FamiNom1", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiNom1)
        Flog.writeline "Nombre del familiar: " & FamiNom1
        '---------------------------------------------------------------
        
        'Segundo Nombre-------------------------------------------------
        Call checkError(Objeto.obtenerNombreApellido("nombre2"), False, 20, "FamiNom2", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiNom2)
        Flog.writeline "Segundo Nombre del familiar: " & FamiNom2
        '---------------------------------------------------------------
        
        'Apellido-------------------------------------------------------
        Call checkError(Objeto.obtenerNombreApellido("apellido"), False, 20, "FamiApe1", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiApe1)
        Flog.writeline "Apellido del familiar: " & FamiApe1
        '---------------------------------------------------------------
        
        'Segundo apellido-----------------------------------------------
        Call checkError(Objeto.obtenerNombreApellido("apellido2"), False, 20, "FamiApe2", arr_Errores, 25, "CHAR", hubo_error, hubo_warning, FamiApe2)
        Flog.writeline "Segundo Apellido del familiar: " & FamiApe2
        '---------------------------------------------------------------
        
        FamiFlgBec = "N"
        
        FamiAnioAE = 0
        
        FamiTipDoc = Objeto.obtenerTipoDoc
        Flog.writeline "tipo de documento del familiar: " & FamiTipDoc
        
        'Familiar Emergencias Medicas
        'Objeto.buscarEstructuras ternro, TeEmeM, Date, "estrcodext"
        
        FamiIdEmeM = familiar.obtenerEmerMed
        If FamiIdEmeM = 0 Then
            'FamiIdEmeM = "NULL"
        End If
        
        If Not existe Then
            'inserto porque no existe
            StrSql = " INSERT INTO famila "
            StrSql = StrSql & "("
            StrSql = StrSql & " HldId, EmpId, FunIdNro, FamiNroRen, FamiStsTip, "
            StrSql = StrSql & " FamiNomCom, FamiStsSex, FamiFchNac, FamiStsCon, FamiFlgCob, "
            StrSql = StrSql & " FamiStsEdu,  FamiFlgYac, FamiFchAni, FamiStsMut, "
            StrSql = StrSql & " FamiIdMut, FamiNroMut, FamiFlgTra, FamiCntEda, FamiFlgCGI, "
            StrSql = StrSql & " FamiDocId, FamiFlgEst, FamiApe1, FamiApe2, FamiNom1, "
            StrSql = StrSql & " FamiNom2, FamiFlgBec, FamiTipDoc "
            StrSql = StrSql & ")"
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "("
            StrSql = StrSql & FamHldId & "," & FamEmpId & "," & Replace(FunIdNro, "-", "") & "," & FamiNroRen & ",'" & Left(FamiStsTip, 1) & "',"
            StrSql = StrSql & "'" & FamiNomCom & "','" & FamiStsSex & "'," & ConvFecha(FamiFchNac) & ",'" & FamiStsCon & "','" & FamiFlgCob & "',"
            StrSql = StrSql & "'" & FamiStsEdu & "', '' , 0,'' ,"
            StrSql = StrSql & " Null, '' ,'" & FamiFlgTra & "'," & FamiCntEda & ",'" & FamiFlgCGI & "',"
            StrSql = StrSql & "'" & FamiDocId & "','" & FamiFlgEst & "','" & FamiApe1 & "','" & FamiApe2 & "','" & FamiNom1 & "',"
            StrSql = StrSql & "'" & FamiNom2 & "','" & FamiFlgBec & "','" & FamiTipDoc & "'"
            StrSql = StrSql & ")"
            On Error Resume Next
            objconnInsert.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Numero de error:" & Err
            If Err <> 0 Then
                Flog.writeline " No se inserto el familiar: " & StrSql
                Err.Clear
            Else
                Flog.writeline " Se inserto el familiar: " & StrSql
                'actualizo la table empsincpayroll
                StrSql = "UPDATE empsincPayroll set sinc=-1"
                StrSql = StrSql & " WHERE ternro=" & ternro
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                listaUpdate = listaUpdate & "," & ternro
                Err.Clear

                'hago update de los campos nulos
                Dim QUERY
                QUERY = ""
                StrSql = "SELECT * FROM famila "
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " HldId=1 AND EmpId=1"
                StrSql = StrSql & " AND FamiNroRen= " & FamiNroRen
                StrSql = StrSql & " AND funidnro=" & FunIdNro
                'StrSql = StrSql & " AND replace(replace(famidocid,'-',''),'.','')='" & Replace(Replace(FamiDocId, "-", ""), ".", "") & "'"
                'StrSql = StrSql & " AND famidocid='" & FamiDocId & "'"
                Flog.writeline "query famila: " & StrSql
                OpenRecordsetExt StrSql, rs_datos, cn_externa
                If Not rs_datos.EOF Then
                    Do While Not rs_datos.EOF
                        For I = 0 To rs_datos.Fields.Count
                            If IsNull(rs_datos.Fields(I).Value) Then
                                Select Case rs_datos.Fields(I).Type
                                    Case 2: 'smallint
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=0"
                                    Case 3: 'int
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=0"
                                    Case 6: 'smallmoney
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=0"
                                    Case 7: 'datetime
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=" & ConvFecha("1753/01/01")
                                    Case 129: 'char
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=' '"
                                    Case 131: 'decimal
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=0"
                                    Case 135: 'smalldatetime
                                        QUERY = QUERY & ", " & rs_datos.Fields(I).Name & "=" & ConvFecha("1753/01/01")
                                End Select
                            End If
                            Flog.writeline "query :" & QUERY
                        Next I
                        StrSql = "UPDATE famila SET" & Replace(QUERY, ",", "", 1, 1) & " WHERE  HldId=1 AND EmpId=1 "
                        StrSql = StrSql & " AND FamiNroRen= " & FamiNroRen
                        StrSql = StrSql & " AND funidnro=" & Replace(FunIdNro, "-", "")
                        'StrSql = StrSql & " AND replace(replace(famidocid,'-',''),'.','')='" & Replace(Replace(FamiDocId, "-", ""), ".", "") & "'"
                        'AND famidocid='" & FamiDocId & "'"
                        Flog.writeline "query Update generico: " & StrSql
                        objconnInsert.Execute StrSql, , adExecuteNoRecords
                    rs_datos.MoveNext
                    Loop
                End If
                'hasta aca
                
                
            End If
        Else
            'actualizo porque existe
            StrSql = " UPDATE famila SET "
            StrSql = StrSql & " FunIdNro=" & Replace(Replace(FunIdNro, "-", ""), ".", "") & ",FamiNroRen=" & FamiNroRen & ",FamiStsTip='" & Left(FamiStsTip, 1) & "',"
            StrSql = StrSql & " FamiNomCom='" & FamiNomCom & "',FamiStsSex='" & FamiStsSex & "',FamiFchNac=" & ConvFecha(FamiFchNac) & ", FamiStsCon='" & FamiStsCon & "', FamiFlgCob='" & FamiFlgCob & "',"
            StrSql = StrSql & " FamiStsEdu='" & FamiStsEdu & "',FamiFlgYac= '' ,FamiFchAni= 0, FamiStsMut= '', "
            StrSql = StrSql & " FamiIdMut= Null, FamiNroMut='', FamiFlgTra='" & FamiFlgTra & "', FamiCntEda=" & FamiCntEda & ", FamiFlgCGI='" & FamiFlgCGI & "', "
            StrSql = StrSql & " FamiDocId='" & FamiDocId & "', FamiFlgEst='" & FamiFlgEst & "', FamiApe1='" & FamiApe1 & "', FamiApe2='" & FamiApe2 & "', FamiNom1='" & FamiNom1 & "', "
            StrSql = StrSql & " FamiNom2='" & FamiNom2 & "', FamiFlgBec='" & FamiFlgBec & "', FamiTipDoc='" & FamiTipDoc & "'"
            StrSql = StrSql & " Where HldId = 1 And EmpId = 1"
            StrSql = StrSql & " AND FamiNroRen= " & FamiNroRen
            StrSql = StrSql & " AND funidnro=" & Replace(FunIdNro, "-", "")
            'StrSql = StrSql & " WHERE EmpId=" & FamEmpId & " AND replace(replace(FUNIDNRO,'-',''),'.','')=" & Replace(Replace(FunIdNro, "-", ""), ".", "")
            'StrSql = StrSql & " AND replace(replace(famidocid,'-',''),'.','')='" & Replace(Replace(FamiDocId, "-", ""), ".", "") & "'"
            '& " AND famidocid='" & FamiDocId & "'"
            On Error Resume Next
            objconnInsert.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Numero de error:" & Err
            If Err <> 0 Then
                Flog.writeline " No se inserto el familiar: " & StrSql
                Err.Clear
            Else
                Flog.writeline " Se inserto el familiar: " & StrSql
                'actualizo la table empsincpayroll
                StrSql = "UPDATE empsincPayroll set sinc=-1"
                StrSql = StrSql & " WHERE ternro=" & ternro
                objconnProgreso.Execute StrSql, , adExecuteNoRecords
                listaUpdate = listaUpdate & "," & ternro
                Err.Clear
            End If
        End If
        

    'End If
        
    'actualizar proceso
    TiempoAcumulado = GetTickCount
    Progreso = Progreso + IncPorc
    Flog.writeline "Actualizo el progreso: " & Progreso
    StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Se actualizo el progreso"
    'hasta aca
    
    If hubo_warning Or hubo_error Then
        Call insertarError(NroProcesoBatch, arr_Errores, str_error, empresa, hubo_error, hubo_warning)
        'Call crearProcesosMensajeria(str_error, empresa)
    End If

    
Next

    str_error = str_error & "</table>" & vbCrLf
    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "
    Call crearProcesosMensajeria(str_error, empresa)
'-------------------HASTA ACA-------------------------------------

End Sub

Public Sub checkError(ByRef dato As Variant, obligatorio As Boolean, Longitud As Integer, columna As String, ByRef arr_Error() As String, ByVal nroCampo As String, ByVal tipoDato As String, ByRef hubo_error As Boolean, ByRef hubo_warning As Boolean, ByRef campo)
'----------------------------------------------------------------------------------------------------------------
'   Checkea longitudes, campos obligatorios, en caso de cortar la corta a la longitud deseada, si el dato no se puede cortar se informa error
'   Dato:           Dato a chequear
'   obligatorio:    Verdadero o falso para checkear si el dato es obligatorio o no
'   longitud:       longitud de la cadena del dato de entrada
'
'----------------------------------------------------------------------------------------------------------------
On Error GoTo errorCheck

    Dim Mensaje As String
    If obligatorio Then
        'Select Case nroCampo
            'Case 1, 4, 15:
                If Len(dato) <= 0 And Longitud > 0 Then
                    Flog.writeline Espacios(Tabulador * 2) & "Error: " & columna & " es nulo o vacio."
                    Mensaje = "ERROR: " & columna & " es nulo o vacio."
                    arr_Error(nroCampo) = Mensaje & "@error"
                    hubo_error = True
                Else
                    If Len(dato) > Longitud And Longitud > 0 Then
                        Flog.writeline Espacios(Tabulador * 2) & "Error: " & columna & " tiene una longitud mayor a: " & Longitud
                        Mensaje = "ERROR: " & columna & " tiene una longitud mayor a: " & Longitud & " dato recibido:" & dato
                        arr_Error(nroCampo) = Mensaje & "@error"
                        hubo_error = True
                    Else
                        If tipoDato = "INT" Then
                            If Not IsNumeric(dato) Then
                                Flog.writeline Espacios(Tabulador * 2) & "Error: " & columna & " Se espara un tipo de dato: " & tipoDato
                                Mensaje = "ERROR en el campo: " & columna & " Se espara un tipo de dato: " & tipoDato & " dato recibido:" & dato
                                arr_Error(nroCampo) = Mensaje & "@error"
                                hubo_error = True
                                campo = dato
                            Else
                                campo = dato
                            End If
                        End If
                    End If
                End If
        'End Select
    Else
        If Longitud > 0 Then
            'If nroCampo <> 1 And nroCampo <> 4 And nroCampo <> 15 Then
                If Len(dato) > Longitud Then
                    Mensaje = "WARNING: " & columna & " tiene una longitud mayor a: " & Longitud & " dato recibido:" & dato
                    dato = Left(dato, Longitud)
                    campo = dato
                    arr_Error(nroCampo) = Mensaje & "@warning"
                    hubo_warning = True
                
                Else
                    campo = dato
                End If
            'End If
        Else
            If tipoDato = "INT" Then
                 If Not IsNumeric(dato) Then
                     Flog.writeline Espacios(Tabulador * 2) & "Error: " & columna & " Se espara un tipo de dato: " & tipoDato
                     Mensaje = "ERROR: " & columna & " Se espara un tipo de dato: " & tipoDato & " dato recibido:" & dato
                     arr_Error(nroCampo) = Mensaje & "@error"
                     hubo_error = True
                 Else
                    campo = dato
                 End If
            Else
                campo = dato
            End If
        End If
    End If
    Exit Sub
errorCheck:
    Flog.writeline "error en validacion"
    Flog.writeline "Error: " & Err.Description
    Exit Sub
End Sub

Sub crearProcesosMensajeria(ByRef str_error As String, ByVal empresa As String)

Dim objRs As New ADODB.Recordset
Dim fs2, MsgFile
Dim titulo As String
Dim bpronroMail As Long
Dim mails As String
Dim notiNro As Long
Dim mailFileName As String
Dim mailFile


    ' Directorio Salidas
    StrSql = "SELECT sis_dirsalidas FROM sistema"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        dirsalidas = objRs!sis_dirsalidas & "\attach"
        Flog.writeline "Directorio de Salidas: " & dirsalidas
    Else
        Flog.writeline "No se encuentra configurado sis_dirsalidas"
        Exit Sub
    End If
    If objRs.State = adStateOpen Then objRs.Close

    'Busco el codigo de la notificacion
    StrSql = "SELECT confval FROM confrep "
    StrSql = StrSql & " WHERE conftipo = 'TN' AND repnro = 474"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        notiNro = objRs!confval
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No esta configurado el tipo de alerta, para el envio de mail."
        notiNro = 0
    End If
    
    If objRs.State = adStateOpen Then objRs.Close
    'FGZ - 04/09/2006 - Saco esto y lo pongo afuera
    StrSql = "insert into batch_proceso "
    StrSql = StrSql & "(btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, bprcparam, "
    StrSql = StrSql & "bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, bprcempleados) "
    StrSql = StrSql & "values (25," & ConvFecha(Date) & ",'" & usuario & "','" & FormatDateTime(Time, 4) & ":00'"
    StrSql = StrSql & ",null,null,'1','Pendiente',null,null,null,null,0,null)"
    
    objConn.Execute StrSql, , adExecuteNoRecords
    
    bpronroMail = getLastIdentity(objConn, "batch_proceso")


    '--------------------------------------------------
    'Busco todos los usuarios a los cuales les tengo que enviar los mails
    StrSql = "SELECT usremail FROM user_per "
    StrSql = StrSql & "inner join noti_usuario on user_per.iduser = noti_usuario.iduser "
    StrSql = StrSql & "where notinro = " & notiNro
    OpenRecordset StrSql, objRs
    mails = ""
    Do Until objRs.EOF
        If Not IsNull(objRs!usremail) Then
            If Len(objRs!usremail) > 0 Then
                mails = mails & objRs!usremail & ";"
            End If
        End If
        objRs.MoveNext
    Loop
    
    mailFileName = dirsalidas & "\msg_" & bpronroMail & "_interface_RHProToPayroll_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now)
    Set fs2 = CreateObject("Scripting.FileSystemObject")
    Set mailFile = fs2.CreateTextFile(mailFileName & ".html", True)
    
    mailFile.writeline "<html><head>"
    mailFile.writeline "<title> Interface RHProToPayroll - RHPro &reg; </title></head><body>"
    'mailFile.writeline "<h4>Errores Detectados</h4>"
    mailFile.writeline str_error
    mailFile.writeline "</body></html>"
    mailFile.Close
    '--------------------------------------------------


    Set fs2 = CreateObject("Scripting.FileSystemObject")
    
    'FGZ - 05/09/2006
    'Los nombres de los archivos para los mails de esta alerta empiezan con el bpronro de este proceso
    Set MsgFile = fs2.CreateTextFile(mailFileName & ".msg", True)
    
    MsgFile.writeline "[MailMessage]"
    MsgFile.writeline "FromName=RHPro - Errores Payroll"
    MsgFile.writeline "Subject=Informe Errores Payroll " & empresa
    MsgFile.writeline "Body1="
    If Len(mailFileName) > 0 Then
       MsgFile.writeline "Attachment=" & mailFileName & ".html"
    Else
       MsgFile.writeline "Attachment="
    End If
    
    MsgFile.writeline "Recipients=" & mails
    'MsgFile.writeline "Recipients=sstremel@rhpro.com"
    
    If objRs.State = adStateOpen Then objRs.Close
    
    StrSql = "select cfgemailfrom,cfgemailhost,cfgemailport,cfgemailuser,cfgemailpassword,cfgssl from conf_email where cfgemailest = -1"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        MsgFile.writeline "FromAddress=" & objRs!cfgemailfrom
        MsgFile.writeline "Host=" & objRs!cfgemailhost
        MsgFile.writeline "Port=" & objRs!cfgemailport
        MsgFile.writeline "User=" & objRs!cfgemailuser
        MsgFile.writeline "Password=" & objRs!cfgemailpassword
    Else
        Flog.writeline "No existen datos configurados para el envio de emails, o no existe configuracion activa"
        Exit Sub
    End If
    MsgFile.writeline "CCO="
    MsgFile.writeline "CC="
    MsgFile.writeline "HTMLBody="
    MsgFile.writeline "HTMLMailHeader="
    MsgFile.writeline "SSL=" & objRs!cfgssl
    
    If objRs.State = adStateOpen Then objRs.Close

End Sub

Public Sub insertarError(ByVal bpronro As Long, ByRef arr_Error() As String, ByRef str_error As String, ByVal empresa As String, ByVal hubo_error As Boolean, ByVal hubo_warning As Boolean)
Dim rs_error As New ADODB.Recordset
Dim indice As Integer
Dim encabezado As String
    
    If str_error = "" Then
        str_error = "<TABLE width='50%' border='1' bordercolor='#333333' cellpadding='0' cellspacing='0'>" & vbCrLf
        str_error = str_error & "<tr><TH style='background-color:#EAEBE3; font-size:11px;' colspan='2' align='center'><b>Listado de Advertencias</b></TH></tr>" & vbCrLf
        str_error = str_error & "<tr><TH style='background-color:#EAEBE3; font-size:11px;' align='center'><b>Empleado</b></TH>" & vbCrLf
        str_error = str_error & "<TH style='background-color:#EAEBE3; font-size:11px;' align='center'><b>Descripcion</b></TH></tr>" & vbCrLf
    End If
        
    'str_error = str_error & "<TABLE width='50%' border='1' bordercolor='#333333' cellpadding='0' cellspacing='0'>" & vbCrLf
    If UBound(arr_Error) > 0 Then
        StrSql = " SELECT empleg FROM empleado WHERE ternro = " & arr_Error(0)
        OpenRecordset StrSql, rs_error
        If Not rs_error.EOF Then
            encabezado = "Legajo Empleado: " & rs_error!empleg
        Else
            StrSql = " SELECT ternro, empleado FROM familiar WHERE ternro = " & arr_Error(0)
            OpenRecordset StrSql, rs_error
            If Not rs_error.EOF Then
                encabezado = "Familiar ternro: " & rs_error!ternro & " del empleado: " & rs_error!empleado
            End If
        End If
        If rs_error.State = adStateOpen Then rs_error.Close
        
        'str_error = str_error & "<tr><TH style='background-color:#8D877F;' align='center'><b>Empleado</b></TH>" & vbCrLf
        'str_error = str_error & "<TH style='background-color:#d13528;color:#FFFFFF;' align='center'><b>" & encabezado & "</b></TH></tr>" & vbCrLf
        'str_error = str_error & "<TH style='background-color:#8D877F;' align='center'><b>Descripcion</b></TH></tr>" & vbCrLf
    End If
    Dim datosError
    If hubo_error Or hubo_warning Then
        For indice = 1 To UBound(arr_Error)
            If arr_Error(indice) <> "" Then
                datosError = Split(arr_Error(indice), "@")
                If datosError(1) = "error" Then
                    str_error = str_error & "<tr><td style='background-color:#FF2A2A;color:#FFFFFF; font-size:11px;'>" & encabezado & "</td>" & vbCrLf
                    str_error = str_error & "<td style='background-color:#FF2A2A;color:#FFFFFF; font-size:11px;'>" & datosError(0) & "</td></tr>" & vbCrLf
                Else
                    str_error = str_error & "<tr><td style='background-color:#FDED08; font-size:11px;'>" & encabezado & "</td>" & vbCrLf
                    str_error = str_error & "<td style='background-color:#FDED08; font-size:11px;'>" & datosError(0) & "</td></tr>" & vbCrLf
                End If
            End If
        Next
    Else
        str_error = str_error & "<tr><td colspan='2' align='center'><b>No hay Errores</b></td></tr>" & vbCrLf
    End If
'    str_error = str_error & "</table>" & vbCrLf
'    str_error = str_error & " <table><tr><td>&nbsp;</td></tr><tr><td>&nbsp;</td></tr></table> "
    
    If rs_error.State = adStateOpen Then rs_error.Close
End Sub

Sub analizarError(ByVal arr_Errores, ByRef campo, ByRef Valor As String, ByVal nroCampo As Integer)
Dim datos
Dim hubo_error
Dim hubo_warning
hubo_error = False
hubo_warning = False
    datos = Split(arr_Errores, "@")
    If UBound(datos) > 0 Then
        If datos(1) <> "" Then
            If datos(1) = "error" Then
                hubo_error = True
            Else
                If datos(1) = "warning" Then
                    hubo_warning = True
                    campo = Valor
                End If
            End If
        End If
    End If
    If Not hubo_error Then
        campo = Valor
    Else
        campo = 0
    End If
End Sub

Public Sub OpenConnExt(strConnectionString As String, ByRef objConn As ADODB.Connection)
' ---------------------------------------------------------------------------------------------
' Descripcion: Abre conexion externa
' Autor      : Martin Ferraro
' Fecha      : 18/04/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    If objConn.State <> adStateClosed Then objConn.Close
    objConn.CursorLocation = adUseClient
    
    'Indica que desde una transacción se pueden ver cambios que no se han producido en otras transacciones.
    objConn.IsolationLevel = adXactReadUncommitted
    
    objConn.CommandTimeout = 3600 'segundos
    objConn.ConnectionTimeout = 60 'segundos
    objConn.Open strConnectionString
End Sub

Sub borrarEmpleados(ByVal ternro As Long)

Dim rs_datos As New ADODB.Recordset
Dim cn_payroll As New ADODB.Connection
Dim j As Integer
Dim fam
Dim nroRenglon As Integer
Dim listaEmpleados
Dim cantFamiliares As Long
Dim Objeto As New datosPersonales
Dim arr_ErroresAux(100) As String
Dim hubo_error As Boolean
Dim hubo_warning As Boolean
Dim FunIdNro
Dim empleado
listaEmpleados = 0
cantFamiliares = 0
'--------me conecto a payroll------------------------
StrSql = "SELECT cnstring FROM conexion "
StrSql = StrSql & " WHERE cnnro=" & nroConexion
OpenRecordset StrSql, rs_datos
If Not rs_datos.EOF Then
    cn_payroll.ConnectionString = rs_datos!cnstring
End If
rs_datos.Close
cn_payroll.Open
'---------------------------------------------------

'-------------------------OBTENGO LOS TERCEROS A SINCRONIZAR--------------------------
If ternro = 0 Then
    StrSql = "SELECT ternro "
    StrSql = StrSql & " FROM empsincPayroll "
    StrSql = StrSql & " WHERE tipo='B' and sinc=0 "
    OpenRecordset StrSql, rs_datos
    If Not rs_datos.EOF Then
        Do While Not rs_datos.EOF
            listaEmpleados = listaEmpleados & "," & rs_datos("ternro")
        rs_datos.MoveNext
        Loop
    End If
    rs_datos.Close
Else
    listaEmpleados = listaEmpleados & "," & ternro
End If
Flog.writeline "Lista de empleados a borrar de Payroll:" & listaEmpleados


fam = Split(listaEmpleados, ",")
cantFamiliares = UBound(fam)
For j = 1 To cantFamiliares
    ternro = fam(j)
    Flog.writeline "Procesamiento del tercero familiar: " & ternro
    
    '------busco los datos del familiar en RHPro--------
    StrSql = " SELECT empleado, renglon FROM empsincPayroll "
    StrSql = StrSql & " WHERE ternro=" & ternro
    OpenRecordset StrSql, rs_datos
    Flog.writeline "Renglon en familiar RHPRO: " & StrSql
    If Not rs_datos.EOF Then
        If EsNulo(rs_datos!empleado) Then
            Flog.writeline "empleado nulo"
            empleado = -1
        Else
            nroRenglon = IIf(EsNulo(rs_datos!renglon), Null, rs_datos!renglon)
            empleado = IIf(EsNulo(rs_datos!empleado), Null, rs_datos!empleado)
        End If
    End If
    rs_datos.Close
    '---------------------------------------------------
    If empleado = -1 Then
        Flog.writeline "No hay borrado"
    Else
        'Objeto.buscarDatosPersonales empTernro
        Objeto.buscarNroDoc empleado, 0
        Call checkError(Objeto.obtenerNroDoc, True, 0, "FunIdNro", arr_ErroresAux, 25, "INT", hubo_error, hubo_warning, FunIdNro)
        Flog.writeline "Documento del parentesco del familiar a borrar: " & FunIdNro
        
        StrSql = "SELECT famidocid, FamiNroRen FROM famila "
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " HldId=1 AND EmpId=1"
        StrSql = StrSql & " AND famiNroRen = " & nroRenglon
        StrSql = StrSql & " AND funidnro= " & Replace(FunIdNro, "-", "")
        Flog.writeline "Familiar a borrar en payroll: " & StrSql
        OpenRecordsetExt StrSql, rs_datos, cn_payroll
        If Not rs_datos.EOF Then
            Do While Not rs_datos.EOF
                StrSql = "DELETE FROM famila "
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " HldId=1 AND EmpId=1"
                StrSql = StrSql & " AND famiNroRen = " & nroRenglon
                StrSql = StrSql & " AND funidnro= " & Replace(FunIdNro, "-", "")
                Flog.writeline "Delete en payroll: " & StrSql
                objconnInsert.Execute StrSql, , adExecuteNoRecords
            rs_datos.MoveNext
            Loop
        End If
    End If
Next

End Sub

