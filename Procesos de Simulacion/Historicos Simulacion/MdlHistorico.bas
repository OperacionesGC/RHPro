Attribute VB_Name = "MdlHistorico"
Option Explicit

'Copia las tablas de simulacion al Historicos de Simulacion.
'Const Version = "1.00"
'Const FechaVersion = "31/08/2011"
'Const Autor = "Lisandro Moro"
'Const Modificacion = "Version inicial"

'Const Version = "1.01"
'Const FechaVersion = "20/10/2011"
'Const Autor = "Lisandro Moro"
'Const Modificacion = "Depuracion" 'Se quitaron tablas de prestamos - embargo - tickets

'Const Version = "1.02"
'Const FechaVersion = "21/12/2011"
'Const Autor = "Lisandro Moro"
'Const Modificacion = "Correccion al insertar en ter_tip "

'Const Version = "1.03"
'Const FechaVersion = "10/12/2012"
'Const Autor = "Lisandro Moro"
'Const Modificacion = "Se cambiaron las listas de elementos por las sql correspondientes en todos los IN."

Const Version = "1.04"
Const FechaVersion = "12/11/2014"
Const Autor = "Ruiz Miriam"
Const Modificacion = "Se corrige problema de superposición de novedades"


'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

Global Descripcion As String
Global Cantidad As Single
Dim I As Long
Global arregloTablas() As String
Global ListaProcesos As String
Global SQLListaProcesos As String
Global ListaCabeceras As String
Global SQLListaCabeceras As String
Global ListaTerceros As String
Global SQLListaTerceros As String
Global pliqnroDesde As Long
Global pliqnroHasta As Long
Global AnioDesde As Long
Global AnioHasta As Long
Global MesDesde As Long
Global MesHasta As Long
Global FechaDesde As String
Global FechaHasta As String
Global CantPasos As Double
Global IncPasos As Double
Global Progreso As Double
Global AfectedRows As Long


'Global pliqdesde As String
'Global pliqhasta As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial
' Autor      : Lisandro Moro
' Fecha      : 31/08/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
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

    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "HistoricoSimulacion" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "Modificacion             : " & Modificacion
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    'Genero el porcesntaje de incremento del progreso del proceso
    CantPasos = 30 ' cantidad de tablas + 1
    IncPasos = 100 / CantPasos
    Progreso = 0
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 307 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Historico(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "No encontró el proceso"
    End If
    
    Flog.writeline "---------------------------------------------------"
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        actualizarProgreso 100
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
Fin:
    Flog.Close
    'If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
        Flog.writeline Espacios(Tabulador * 0) & " Error General: " & Err.Description
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


Public Sub Historico(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que realiza el la copia delas tablas de simulacion a las de Historico de simulaciòn (sim_ <-> sim_his_)
' Autor      : Lisandro Moro
' Fecha      : 31/08/2011
' Modificacion:
' --------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset

Dim cantRegistros
Dim CantidadEmpleados As Long
Dim Sep As String
Dim simhisnro As Long
Dim tipo As String

Dim ArrParam
Dim ArrCabliq

On Error GoTo CE

TiempoAcumulado = GetTickCount

'----------------------------------------------------------------------------
' Levanto cada parametro por separado
'----------------------------------------------------------------------------
'Params tipo
'---------------------
'G  'Generar Histórico
'R  'Recuperar Histórico
'B  'Borrar Histórico

Flog.writeline "Levantando parametros. " & Parametros
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
         'creo un array con todos los numeros de procesos que me van a servir para las novedades retroactivas
         ArrParam = Split(Parametros, "@")
         
         tipo = ArrParam(0)
         If UBound(ArrParam) > 0 Then
            simhisnro = CLng(ArrParam(1))
         End If
    End If
Else
    Flog.writeline "Error - Parametros nulos"
    HuboError = True
    Exit Sub
End If
Flog.writeline "Terminó de levantar los parametros"


'Busco los parametros correspondientes al historico simulacion
buscarDatos simhisnro

'arregloTablas(0) = "sim_his_acu_liq"
'arregloTablas(1) = "sim_his_acu_mes"
'arregloTablas(2) = "sim_his_cabliq"
'arregloTablas(3) = "sim_his_curva"
'arregloTablas(4) = "sim_his_DatosBaja"
'arregloTablas(5) = "sim_his_desliq"
'arregloTablas(6) = "sim_his_desmen"
'arregloTablas(7) = "sim_his_detliq"
'arregloTablas(8) = "sim_his_embargo"
'arregloTablas(9) = "sim_his_embcuota"
'arregloTablas(10) = "sim_his_emp_lic"
'arregloTablas(11) = "sim_his_emp_ticket"
'arregloTablas(12) = "sim_his_emp_tikdist"
'arregloTablas(13) = "sim_his_empleado"
'arregloTablas(14) = "sim_his_fases"
'arregloTablas(15) = "sim_his_ficharet"
'arregloTablas(16) = "sim_his_gti_acunov"
'arregloTablas(17) = "sim_his_his_estructura"
'arregloTablas(18) = "sim_his_impgralarg"
'arregloTablas(19) = "sim_his_impmesarg"
'arregloTablas(20) = "sim_his_impproarg"
'arregloTablas(21) = "sim_his_novaju"
'arregloTablas(22) = "sim_his_novemp"
'arregloTablas(23) = "sim_his_pre_cuota"
'arregloTablas(24) = "sim_his_prestamo"
'arregloTablas(25) = "sim_his_proceso"
'arregloTablas(26) = "sim_his_rep_borradordeta"
'arregloTablas(27) = "sim_his_rep_borrdeta_det"
'arregloTablas(28) = "sim_his_traza"
'arregloTablas(29) = "sim_his_traza_gan"
'arregloTablas(30) = "sim_his_traza_gan_item_top"
'arregloTablas(31) = "sim_his_vacpagdesc"
'arregloTablas(32) = "sim_his_vales"
'arregloTablas(33) = "sim_historico" ' esta no va

Select Case tipo
    Case "G" 'Generar Histórico
        GenerarHistórico (simhisnro)
    Case "R" 'Recuperar Histórico
        RecuperarHistórico (simhisnro)
    Case "B" 'Borrar Histórico
        borrarHistoricoCompleto (simhisnro)
End Select

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error (sub Copiado): " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyRollbackTrans
    MyBeginTrans
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Fix(((CantidadEmpleados - cantRegistros) * 100) / CantidadEmpleados) & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & bpronro
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub

Private Sub GenerarHistórico(simhisnro)

    'Borramos primero
    borrarHistoricoSimHis (simhisnro)
    actualizarProgreso

    'Copiamos los datos desde las sim_ a las sim_his_
    'CopiarTablaHis
    
    'sim_cabliq -> sim_his_cabliq
    Flog.writeline "sim_cabliq -> sim_his_cabliq "
    StrSql = " INSERT INTO sim_his_cabliq SELECT " & simhisnro & ", * FROM sim_cabliq "
    StrSql = StrSql & " WHERE sim_cabliq.cliqnro IN (" & SQLListaCabeceras & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_acu_liq -> sim_his_acu_liq
    Flog.writeline "sim_acu_liq -> sim_his_acu_liq"
    StrSql = " INSERT INTO sim_his_acu_liq SELECT " & simhisnro & ", * FROM sim_acu_liq "
    StrSql = StrSql & " WHERE sim_acu_liq.cliqnro IN (" & SQLListaCabeceras & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_acu_mes -> sim_his_acu_mes
    Flog.writeline "sim_acu_mes -> sim_his_acu_mes"
    StrSql = " INSERT INTO sim_his_acu_mes SELECT " & simhisnro & ", * FROM sim_acu_mes "
    StrSql = StrSql & " WHERE ((amanio >=" & AnioDesde & " AND ammes >=  " & MesDesde & " )"
    StrSql = StrSql & " AND(amanio <= " & AnioHasta & " AND ammes >= " & MesHasta & "))"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_detliq -> sim_his_detliq
    Flog.writeline "sim_detliq -> sim_his_detliq"
    StrSql = " INSERT INTO sim_his_detliq SELECT " & simhisnro & ", * FROM sim_detliq "
    StrSql = StrSql & " WHERE sim_detliq.cliqnro IN (" & SQLListaCabeceras & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso

    'sim_curva -> sim_his_curva
'    Flog.writeline "sim_curva -> sim_his_curva"
'    StrSql = " INSERT INTO sim_his_curva SELECT " & simhisnro & ", * FROM sim_curva "
'    'StrSql = StrSql & " WHERE "
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
'    actualizarProgreso
    
    'sim_DatosBaja -> sim_his_DatosBaja
    Flog.writeline "sim_DatosBaja -> sim_his_DatosBaja"
    StrSql = " INSERT INTO sim_his_DatosBaja SELECT " & simhisnro & ", * FROM sim_DatosBaja "
    StrSql = StrSql & " WHERE sim_DatosBaja.pronro IN (" & SQLListaProcesos & ")"
    'StrSql = StrSql & " AND sim_DatosBaja.ternro IN (" & ListaTerceros & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso

    'sim_desliq -> sim_his_desliq
    Flog.writeline "sim_desliq -> sim_his_desliq"
    StrSql = " INSERT INTO sim_his_desliq SELECT " & simhisnro & ", * FROM sim_desliq "
    StrSql = StrSql & " WHERE sim_desliq.pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_desmen -> sim_his_desmen
    Flog.writeline "sim_desmen -> sim_his_desmen"
    StrSql = " INSERT INTO sim_his_desmen SELECT " & simhisnro & ", * FROM sim_desmen "
    StrSql = StrSql & " WHERE sim_desmen.empleado IN (" & SQLListaTerceros & ")"
    StrSql = StrSql & " AND ((desfecdes >=" & ConvFecha(FechaDesde)
    StrSql = StrSql & " AND desfecdes <=" & ConvFecha(FechaHasta) & ") "
    StrSql = StrSql & " OR (desfechas >=" & ConvFecha(FechaDesde)
    StrSql = StrSql & " AND desfechas <=" & ConvFecha(FechaHasta) & ") "
    StrSql = StrSql & " OR (desfecdes <=" & ConvFecha(FechaDesde)
    StrSql = StrSql & " AND desfechas >=" & ConvFecha(FechaDesde) & ") "
    StrSql = StrSql & " OR (desfecdes <=" & ConvFecha(FechaHasta)
    StrSql = StrSql & " AND desfechas >=" & ConvFecha(FechaHasta) & ")) "
    'StrSql = StrSql & " AND ((desfecdes >=  " & ConvFecha(FechaDesde) & " )"
    'StrSql = StrSql & " AND (desfechas <= " & ConvFecha(FechaHasta) & "))"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
'    'sim_embargo -> sim_his_embargo
'    Flog.writeline "sim_embargo -> sim_his_embargo"
'    StrSql = " INSERT INTO sim_his_embargo SELECT " & simhisnro & ", * FROM sim_embargo "
'    StrSql = StrSql & " WHERE sim_embargo.ternro IN (" & ListaTerceros & ")"
'    'StrSql = StrSql & " WHERE sim_embargo.cliqnro IN (" & ListaCabeceras & ")"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
'    actualizarProgreso
    
'    'sim_embcuota -> sim_his_embcuota
'    Flog.writeline "sim_embcuota -> sim_his_embcuota"
'    StrSql = " INSERT INTO sim_his_embcuota SELECT " & simhisnro & ", * FROM sim_embcuota "
'    StrSql = StrSql & " WHERE embnro IN ( SELECT embnro FROM sim_embargo WHERE sim_embargo.ternro IN (" & ListaTerceros & "))"
'    'StrSql = StrSql & " WHERE sim_embcuota.embnro IN ( SELECT embnro FROM sim_his_embargo WHERE simhisnro = " & simhisnro & ")"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
'    actualizarProgreso
    
    'sim_emp_lic -> 'sim_his_emp_lic 'ver si no va por lista de proesos
    Flog.writeline "sim_emp_lic -> 'sim_his_emp_lic"
    StrSql = " INSERT INTO sim_his_emp_lic SELECT " & simhisnro & ", * FROM sim_emp_lic "
    'StrSql = StrSql & " WHERE ((elfechadesde >=  " & ConvFecha(FechaDesde) & " )"
    'StrSql = StrSql & " AND (elfechahasta <= " & ConvFecha(FechaHasta) & "))"
    StrSql = StrSql & " WHERE empleado IN (" & SQLListaTerceros & ") "
    StrSql = StrSql & " AND ((elfechadesde >=" & ConvFecha(FechaDesde)
    StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(FechaHasta) & ") "
    StrSql = StrSql & " OR (elfechahasta >=" & ConvFecha(FechaDesde)
    StrSql = StrSql & " AND elfechahasta <=" & ConvFecha(FechaHasta) & ") "
    StrSql = StrSql & " OR (elfechadesde <=" & ConvFecha(FechaDesde)
    StrSql = StrSql & " AND elfechahasta >=" & ConvFecha(FechaDesde) & ") "
    StrSql = StrSql & " OR (elfechadesde <=" & ConvFecha(FechaHasta)
    StrSql = StrSql & " AND elfechahasta >=" & ConvFecha(FechaHasta) & ")) "
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso

'    'sim_emp_ticket -> sim_his_emp_ticket
'    Flog.writeline "sim_emp_ticket -> sim_his_emp_ticket"
'    StrSql = " INSERT INTO sim_his_emp_ticket SELECT " & simhisnro & ", * FROM sim_emp_ticket "
'    StrSql = StrSql & " WHERE ((etikfecha >=  " & ConvFecha(FechaDesde) & " )"
'    StrSql = StrSql & " AND (etikfecha <= " & ConvFecha(FechaHasta) & "))"
'    StrSql = StrSql & " AND empleado IN (" & ListaTerceros & ") "
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
'    actualizarProgreso
    
'    'sim_emp_tikdist -> sim_his_emp_tikdist
'    Flog.writeline "sim_emp_tikdist -> sim_his_emp_tikdist"
'    StrSql = " INSERT INTO sim_his_emp_tikdist SELECT " & simhisnro & ", * FROM sim_emp_tikdist "
'    StrSql = StrSql & " WHERE etiknro IN ( SELECT etiknro FROM sim_his_emp_ticket "
'    StrSql = StrSql & "     WHERE simhisnro = " & simhisnro & " )"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
'    actualizarProgreso
    
    'sim_empleado -> sim_his_empleado
    Flog.writeline "sim_empleado -> sim_his_empleado"
    StrSql = " INSERT INTO sim_his_empleado SELECT " & simhisnro & ", * FROM sim_empleado "
    StrSql = StrSql & " WHERE sim_empleado.ternro IN (" & SQLListaTerceros & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso

    'sim_fases -> sim_his_fases
    Flog.writeline "sim_fases -> sim_his_fases"
    StrSql = " INSERT INTO sim_his_fases SELECT " & simhisnro & ", * FROM sim_fases "
    'StrSql = StrSql & " WHERE ternro IN ( SELECT ternro FROM sim_his_fases WHERE simhisnro = " & simhisnro & " )"
    StrSql = StrSql & " WHERE sim_fases.empleado IN (" & SQLListaTerceros & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_ficharet -> sim_his_ficharet
    Flog.writeline "sim_ficharet -> sim_his_ficharet"
    StrSql = " INSERT INTO sim_his_ficharet SELECT " & simhisnro & ", * FROM sim_ficharet "
    StrSql = StrSql & " WHERE sim_ficharet.pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_gti_acunov -> sim_his_gti_acunov
    Flog.writeline "sim_gti_acunov -> sim_his_gti_acunov"
    StrSql = " INSERT INTO sim_his_gti_acunov SELECT " & simhisnro & ", * FROM sim_gti_acunov "
    StrSql = StrSql & " WHERE sim_gti_acunov.ternro IN (" & SQLListaTerceros & ")"
    StrSql = StrSql & " AND (sim_gti_acunov.pronro IS NULL OR sim_gti_acunov.pronro IN (" & SQLListaProcesos & "))"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_his_estructura -> 'sim_his_his_estructura
    Flog.writeline "sim_his_estructura -> 'sim_his_his_estructura"
    StrSql = " INSERT INTO sim_his_his_estructura SELECT " & simhisnro & ", * FROM sim_his_estructura "
    StrSql = StrSql & " WHERE sim_his_estructura.ternro IN ( " & SQLListaTerceros & ")"
    'StrSql = StrSql & " WHERE sim_his_estructura.ternro IN ( SELECT ternro FROM sim_his_empleado WHERE simhisnro = " & simhisnro & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso

    'sim_impgralarg -> sim_his_impgralarg
    Flog.writeline "sim_impgralarg -> sim_his_impgralarg"
    StrSql = " INSERT INTO sim_his_impgralarg SELECT " & simhisnro & ", * FROM sim_impgralarg "
    StrSql = StrSql & " WHERE sim_impgralarg.pronro IN (" & SQLListaProcesos & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_impmesarg -> sim_his_impmesarg
    Flog.writeline "sim_impmesarg -> sim_his_impmesarg"
    StrSql = " INSERT INTO sim_his_impmesarg SELECT " & simhisnro & ", * FROM sim_impmesarg "
    StrSql = StrSql & " WHERE sim_impmesarg.ternro IN (" & SQLListaTerceros & " )"
    StrSql = StrSql & " AND ((imaanio >=" & AnioDesde & " AND imames >=  " & MesDesde & " )"
    StrSql = StrSql & " AND(imaanio <= " & AnioHasta & " AND imames >= " & MesHasta & "))"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso

    'sim_impproarg -> sim_his_impproarg
    Flog.writeline "sim_impproarg -> sim_his_impproarg"
    StrSql = " INSERT INTO sim_his_impproarg SELECT " & simhisnro & ", * FROM sim_impproarg "
    StrSql = StrSql & " WHERE sim_impproarg.cliqnro IN (" & SQLListaCabeceras & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_novaju -> sim_his_novaju
    Flog.writeline "sim_novaju -> sim_his_novaju"
    StrSql = " INSERT INTO sim_his_novaju SELECT " & simhisnro & ", * FROM sim_novaju "
    StrSql = StrSql & " WHERE sim_novaju.empleado IN (" & SQLListaTerceros & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_novemp -> sim_his_novemp
    Flog.writeline "sim_novemp -> sim_his_novemp"
    StrSql = " INSERT INTO sim_his_novemp SELECT " & simhisnro & ", * FROM sim_novemp "
    StrSql = StrSql & " WHERE sim_novemp.empleado IN (" & SQLListaTerceros & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
'    'sim_prestamo -> sim_his_prestamo
'    Flog.writeline "sim_prestamo -> sim_his_prestamo"
'    StrSql = " INSERT INTO sim_his_prestamo SELECT " & simhisnro & ", * FROM sim_prestamo "
'    StrSql = StrSql & " WHERE sim_prestamo.ternro IN (" & ListaTerceros & " )"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
'    actualizarProgreso
        
'    'sim_pre_cuota -> sim_his_pre_cuota
'    Flog.writeline "sim_pre_cuota -> sim_his_pre_cuota"
'    StrSql = " INSERT INTO sim_his_pre_cuota SELECT " & simhisnro & ", * FROM sim_pre_cuota "
'    StrSql = StrSql & " WHERE prenro IN ( SELECT prenro FROM sim_prestamo WHERE ternro IN ( " & ListaTerceros & " ))"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
'    actualizarProgreso
    
    'sim_his_proceso -> sim_his_proceso
    Flog.writeline "sim_his_proceso -> sim_his_proceso"
    StrSql = " INSERT INTO sim_his_proceso SELECT " & simhisnro & ", * FROM sim_proceso "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso

    'sim_rep_borradordeta -> sim_his_rep_borradordeta
    Flog.writeline "sim_rep_borradordeta -> sim_his_rep_borradordeta"
    StrSql = " INSERT INTO sim_his_rep_borradordeta SELECT " & simhisnro & ", * FROM sim_rep_borradordeta "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    StrSql = StrSql & " AND ternro IN (" & SQLListaTerceros & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso

    'sim_rep_borrdeta_det -> sim_his_rep_borrdeta_det
    Flog.writeline "sim_rep_borrdeta_det -> sim_his_rep_borrdeta_det"
    StrSql = " INSERT INTO sim_his_rep_borrdeta_det SELECT " & simhisnro & ", * FROM sim_rep_borrdeta_det "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    StrSql = StrSql & " AND ternro IN (" & SQLListaTerceros & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso

    'sim_traza -> sim_his_traza
    Flog.writeline "sim_traza -> sim_his_traza"
    StrSql = " INSERT INTO sim_his_traza SELECT " & simhisnro & ", * FROM sim_traza "
    StrSql = StrSql & " WHERE cliqnro IN (" & SQLListaCabeceras & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_traza_gan -> sim_his_traza_gan
    Flog.writeline "sim_traza_gan -> sim_his_traza_gan"
    StrSql = " INSERT INTO sim_his_traza_gan SELECT " & simhisnro & ", * FROM sim_traza_gan "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    StrSql = StrSql & " AND ternro IN (" & SQLListaTerceros & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_traza_gan_item_top -> sim_his_traza_gan_item_top
    Flog.writeline "sim_traza_gan_item_top -> sim_his_traza_gan_item_top"
    StrSql = " INSERT INTO sim_his_traza_gan_item_top SELECT " & simhisnro & ", * FROM sim_traza_gan_item_top "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    StrSql = StrSql & " AND ternro IN (" & SQLListaTerceros & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_vacpagdesc -> sim_his_vacpagdesc
    Flog.writeline "sim_vacpagdesc -> sim_his_vacpagdesc"
    StrSql = " INSERT INTO sim_his_vacpagdesc SELECT " & simhisnro & ", * FROM sim_vacpagdesc "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    StrSql = StrSql & " AND ternro IN (" & SQLListaTerceros & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
    
    'sim_vales -> sim_his_vales
    Flog.writeline "sim_vales -> sim_his_vales"
    StrSql = " INSERT INTO sim_his_vales SELECT " & simhisnro & ", * FROM sim_vales "
    StrSql = StrSql & " WHERE empleado IN (" & SQLListaTerceros & ")"
    StrSql = StrSql & " AND pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    actualizarProgreso
      
End Sub


Private Sub buscarDatos(simhisnro As Long)
    
    StrSql = " SELECT pliqnrodesde, pliqnroHasta FROM sim_historico WHERE simhisnro = " & simhisnro
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "Error: No se encontro el Historico: " & simhisnro
        Flog.writeline " No se puede continuar. "
        Exit Sub
    Else
        pliqnroDesde = rs!pliqnroDesde
        pliqnroHasta = rs!pliqnroHasta
        Flog.writeline "Periodo Desde: " & pliqnroDesde
        Flog.writeline "Periodo Hasta: " & pliqnroHasta
    End If
    rs.Close
    
    'Desde
    StrSql = " SELECT pliqDesde, pliqanio, pliqmes FROM periodo WHERE pliqnro = " & pliqnroDesde
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "Error: No se encontro el Periodo desde: " & pliqnroDesde
        Flog.writeline " No se puede continuar. "
        Exit Sub
    Else
        AnioDesde = rs!pliqanio
        MesDesde = rs!pliqmes
        pliqdesde = rs!pliqdesde
    End If
    rs.Close
    
    'Hasta
    StrSql = " SELECT pliqHasta, pliqanio, pliqmes FROM periodo WHERE pliqnro = " & pliqnroHasta
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "Error: No se encontro el Periodo Hasta: " & pliqnroHasta
        Flog.writeline " No se puede continuar. "
        Exit Sub
    Else
        AnioHasta = rs!pliqanio
        MesHasta = rs!pliqmes
        pliqhasta = rs!pliqhasta
    End If
    rs.Close
    
    
    FechaDesde = "01/" & Format(MesDesde, "00") & "/" & AnioDesde
    Flog.writeline "Fecha Desde:" & FechaDesde
    
    FechaHasta = DateAdd("d", -1, DateAdd("m", 1, "01/" & MesHasta & "/" & AnioHasta))
    Flog.writeline "Fecha Hasta:" & FechaHasta

    
    ListaProcesos = "0"
    StrSql = " SELECT pronro FROM sim_proceso "
    StrSql = StrSql & " INNER JOIN periodo ON sim_proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " WHERE pliqdesde >= " & ConvFecha(pliqdesde) & " And pliqhasta <= " & ConvFecha(pliqhasta)
    SQLListaProcesos = StrSql
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "Error: No se encontraron procesos entre los periodos. Periodo desde: " & pliqnroDesde & " - Periodo Hasta: " & pliqnroHasta
        Flog.writeline " No se puede continuar. "
        ''Exit Sub
    Else
        ListaProcesos = "0"
        Do While Not rs.EOF
            ListaProcesos = ListaProcesos & "," & rs!Pronro
            rs.MoveNext
        Loop
        Flog.writeline "Lisata de procesos encontrados:" & ListaProcesos
    End If
    rs.Close
    
    
    ListaCabeceras = "0"
    StrSql = " SELECT cliqnro FROM sim_cabliq "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ") "
    SQLListaCabeceras = "SELECT cliqnro FROM sim_cabliq WHERE pronro IN (" & SQLListaProcesos & ") "
    OpenRecordset StrSql, rs
    If rs.EOF Then
        ListaCabeceras = "0"
        Flog.writeline "Error: No se encontraron Cabeceras de Liquidacion entre los procesos: " & ListaProcesos
        Flog.writeline "Lisata de Cabeceras encontrados:" & ListaCabeceras
        'Flog.writeline "No se puede continuar. "
        'Exit Sub
    Else
        ListaCabeceras = "0"
        Do While Not rs.EOF
            ListaCabeceras = ListaCabeceras & "," & rs!cliqnro
            rs.MoveNext
        Loop
        Flog.writeline "Lisata de Cabeceras encontrados:" & ListaCabeceras
    End If
    rs.Close
    
    StrSql = " SELECT DISTINCT empleado FROM sim_cabliq "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ") "
    SQLListaTerceros = " SELECT DISTINCT empleado FROM sim_cabliq WHERE pronro IN (" & SQLListaProcesos & ") "
    OpenRecordset StrSql, rs
    If rs.EOF Then
        ListaTerceros = "0"
        Flog.writeline "Error: No se encontraron Terceros entre los procesos: " & ListaProcesos
        Flog.writeline "Lisata de Terceros encontrados:" & ListaTerceros
        'Flog.writeline "No se puede continuar. "
        'Exit Sub
    Else
        ListaTerceros = "0"
        Do While Not rs.EOF
            ListaTerceros = ListaTerceros & "," & rs!Empleado
            rs.MoveNext
        Loop
        Flog.writeline "Lisata de Terceros encontrados:" & ListaTerceros
    End If
    rs.Close
    
    
    
End Sub

Private Sub RecuperarHistórico(simhisnro As Long)
    Dim tablaS As String
    Dim tablaSH As String
    Dim cols As String
    Dim cond As String
    Dim cond2 As String
    
    'Borramos primero
    borrarSimulacion
    actualizarProgreso
    
    'Copiamos los datos desde las sim_ a las sim_his_
    
    'sim_cabliq -> sim_his_cabliq
    tablaS = "sim_cabliq"
    tablaSH = "sim_his_cabliq"
    cols = "cliqnro, pronro, empleado, ppagnro, cliqtexto, cliqdesde, cliqhasta, nrorecibo, cliqnrocorr, nroimp, fechaimp, entregado, fechaentrega, Procesado "
    cond = " WHERE simhisnro = " & simhisnro '<- Generica para todos por ahora.
    CopiarTablaHis tablaSH, simhisnro, cols, cond, True
    actualizarProgreso
    
    'sim_his_acu_liq -> sim_acu_liq
    tablaSH = "sim_his_acu_liq"
    cols = "acunro, cliqnro, almonto, alcant, alfecret, almontoreal"
    'cond2 = " AND ternro NOT IN (SELECT ternro FROM sim_empleado )"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
    
    'sim_his_acu_mes -> sim_acu_mes
    tablaSH = "sim_his_acu_mes"
    cols = "ternro, acunro, amanio, ammonto, amcant, ammes, ammontoreal"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
    
    'sim_his_detliq -> sim_detliq
    tablaSH = "sim_his_detliq"
    cols = "concnro, dlimonto, dlifec, cliqnro, dlicant, dlimonto_base, dliporcent, dlitexto, fornro, tconnro, dliretro, ajustado, dliqdesde, dliqhasta"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso

'    'sim_his_curva -> sim_curva
'    tablaSH = "sim_his_curva"
'    cols = "curnro, curdesc, curmes1, curmes2, curmes3, curmes4, curmes5, curmes6, curmes7, curmes8, curmes9, curmes10, curmes11, curmes12"
'    CopiarTablaHis tablaSH, simhisnro, cols, cond, True
'    actualizarProgreso
    
    'sim_his_DatosBaja -> sim_DatosBaja
    tablaSH = "sim_his_DatosBaja"
    cols = "ternro, pronro, caunro, bajfec"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso

    'sim_his_desliq -> sim_desliq
    tablaSH = "sim_his_desliq"
    cols = "itenro, empleado, dlfecha, pronro, dlmonto, dlprorratea"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
    
    'sim_his_desmen -> sim_desmen
    tablaSH = "sim_his_desmen"
    cols = "itenro, empleado, desmondec, desmenprorra, desano, desfecdes, desfechas, descuit, desrazsoc, pronro, sim"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso

'    'sim_his_embargo -> sim_embargo
'    tablaSH = "sim_his_embargo"
'    cols = "tpenro, ternro, embest, embprioridad, embdesext, embimp, embcantcuo, embquincenal, embnro, embanioini, embmesini, embquinini, embaniofin, embmesfin, embquifin, bennom, embfecaut, fpagnro, embimpfij, embimppor, embexp, embcar, embjuz, embcgoem, embiva, embnroof, embsec, bencuit, bencuenta, benbanco, embdeuda , embimpmin, embfecest, monnro, retley"
'    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
'    actualizarProgreso
    
'    'sim_his_embcuota -> sim_embcuota
'    tablaSH = "sim_his_embcuota"
'    cols = "embnro, embcimp, embcnro, pronro, embccancela, embcanio, embcmes, embcquin, embcimpreal, embcretro, embcaran, cliqnro"
'    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
'    actualizarProgreso
    
    'sim_his_emp_lic -> sim_emp_lic 'ver si no va por lista de proesos el otro
    tablaSH = "sim_his_emp_lic"
    cols = "elfechadesde, elfechahasta, empleado, tdnro, emp_licnro, eldiacompleto, elhoradesde, elhorahasta, thnro, elcantdias, elcantdiashab, elcantdiasfer, elcanthrs, eltipo, elorden , elmaxhoras, licnrosig, elfechacert, Pronro, licestnro, elobs"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
                      
'    'sim_his_emp_ticket -> sim_emp_ticket
'    tablaSH = "sim_his_emp_ticket"
'    cols = "etiknro, empleado, tiknro, tikpednro, etikfecha, etikmonto, etikcant, etikhora, etikuser, etikmanual, pronro"
'    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
'    actualizarProgreso
    
'    'sim_his_emp_tikdist -> sim_emp_tikdist
'    tablaSH = "sim_his_emp_tikdist"
'    cols = "etiknro, tikvalnro, tiknro, etikdmonto, etikdmontouni, etikdcant"
'    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
'    actualizarProgreso
    
    'sim_his_empleado -> sim_empleado
    tablaSH = "sim_his_empleado"
    cols = "empleg, empfecbaja, empfbajaprev, empest, empfaltagr, ternro, empremu"
    cond2 = " AND ternro NOT IN (SELECT ternro FROM sim_empleado )"
    CopiarTablaHis tablaSH, simhisnro, cols, cond & cond2, False
    actualizarProgreso
    
    'Genero los tertip(26)
    StrSql = " INSERT INTO ter_tip (tipnro, ternro) "
    StrSql = StrSql & " SELECT 26, ternro FROM sim_his_empleado "
    StrSql = StrSql & cond '" WHERE simhisnro = " & simhisnro '<- Generica para todos por ahora.
    StrSql = StrSql & " AND ternro IN (SELECT ternro FROM sim_empleado WHERE ternro NOT IN (SELECT ternro FROM ter_tip where tipnro = 26))"
    'StrSql = StrSql & " AND ternro NOT IN(SELECT ternro FROM ter_tip WHERE tipnro = 26) "
    'StrSql = StrSql & " AND ternro NOT IN (SELECT ternro FROM sim_empleado )"
    Flog.writeline StrSql
    objConn.Execute StrSql, , adExecuteNoRecords
    actualizarProgreso
    
    'sim_his_fases -> sim_fases
    tablaSH = "sim_his_fases"
    cols = "fasnro, empleado, caunro, altfec, bajfec, estado, empantnro, sueldo, vacaciones, indemnizacion, real, fasrecofec, CATalta, CATbaja"
    cond2 = " AND sim_his_fases.empleado NOT IN (SELECT ternro FROM sim_empleado )"
    CopiarTablaHis tablaSH, simhisnro, cols, cond & cond2, False
    actualizarProgreso
    
    'sim_his_ficharet -> sim_ficharet
    tablaSH = "sim_his_ficharet"
    cols = "fecha, importe, pronro, liqsistema, empleado"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso

    'sim_his_gti_acunov -> sim_gti_acunov
    tablaSH = "sim_his_gti_acunov"
    cols = "acnovnro, concnro, acnovvalor, tpanro, acnovhornro, acnovfecaprob, ternro, gpanro, pronro"
    cond2 = " AND (sim_his_gti_acunov.acnovnro NOT IN (select acnovnro from sim_gti_acunov) ) "
    CopiarTablaHis tablaSH, simhisnro, cols, cond & cond2, False
    actualizarProgreso
    
    'sim_his_his_estructura -> sim_his_estructura
    tablaSH = "sim_his_his_estructura"
    cols = "tenro, ternro, estrnro, htetdesde, htethasta, hismotivo, tipmotnro"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso

    'sim_his_impgralarg -> sim_impgralarg
    tablaSH = "sim_his_impgralarg"
    cols = "Pronro , tconnro, ipgtopemonto, ipgtopecant"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
    
    'sim_his_impmesarg -> sim_impmesarg
    tablaSH = "sim_his_impmesarg"
    cols = "ternro , acuNro, tconnro, imaanio, imames, imacant, imamonto"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso

    'sim_his_impproarg -> sim_impproarg
    tablaSH = "sim_his_impproarg"
    cols = "acunro, cliqnro, tconnro, ipacant, ipamonto"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
    
    'sim_his_novaju -> sim_novaju
    tablaSH = "sim_his_novaju"
    cols = "concnro, empleado, navalor, navigencia, nadesde, nahasta, naretro, naajuste, pronro, nanro, natexto, napliqdesde, napliqhasta, sim"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, True
    actualizarProgreso

    'sim_his_novemp -> sim_novemp
    tablaSH = "sim_his_novemp"
    cols = "concnro, tpanro, empleado, nevalor, nevigencia, nedesde, nehasta, neretro, nepliqdesde, nepliqhasta, pronro, nenro, netexto, sim, tipmotnro, motivo"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, True
    actualizarProgreso

'    'sim_his_prestamo -> sim_prestamo
'    tablaSH = "sim_his_prestamo"
'    cols = "prenro, predesc, preimp, preanio, precantcuo, estnro, ternro, preimpcuo, premes, pretpotor, prequin, pretna, monnro, quincenal, lnprenro, perfecaut, iduser, prefecotor, sucursal , precompr, pliqnro, prediavto, preiva, preotrosgas, nroimp, fechaimp"
'    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
'    actualizarProgreso

'    'sim_his_pre_cuota -> sim_pre_cuota
'    tablaSH = "sim_his_pre_cuota"
'    cols = "prenro, cuoimp, cuonro, pronro, cuocancela, cuoano, cuomes, cuoquin, cuofecvto, cuonrocuo, cuogastos, cuoiva, cuototal, cuocapital, cuointeres, cuosaldo"
'    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
'    actualizarProgreso
    
    'sim_his_his_proceso -> sim_proceso
    tablaSH = "sim_his_proceso"
    cols = " pronro, prodesc, propend, profeccorr, profecplan, pliqnro, tprocnro, prosist, profecpago, profecini, profecfin, empnro, proaprob, proestdesc, protipoSim, profecbaja, caunro, pliqnroreal, tprocnroreal, pronroreal, proconceptos, procajuretro, pronropago"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, True
    actualizarProgreso

    'sim_his_rep_borradordeta -> sim_rep_borradordeta
    tablaSH = "sim_his_rep_borradordeta"
    cols = "bpronro, Legajo, ternro, Pronro, prodesc, pliqdesc, Descripcion, Apellido, apellido2, nombre, nombre2, empresa, emprnro, pliqnro, fecalta, fecbaja, contrato, categoria, centrocosto, documento, acumval1, acumdesc1, acumval2, acumdesc2, acumval3, acumdesc3, acumval4, acumdesc4, tedabr1, tedabr2, tedabr3, estrdabr1, estrdabr2, estrdabr3, orden, tidsigla, acumval5, acumdesc5, acumval6, acumdesc6, acumval7, acumdesc7, acumval8, acumdesc8, Cuil"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso

    'sim_his_rep_borrdeta_det -> sim_rep_borrdeta_det
    tablaSH = "sim_his_rep_borrdeta_det"
    cols = "bpronro, ternro, pronro, concabr, conccod, concnro, concimp, dlicant, dlimonto"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso

    'sim_his_traza -> sim_traza
    tablaSH = "sim_his_traza"
    cols = "cliqnro, concnro, tpanro, tradesc, travalor, trafrecuencia, trazaformula, tranro"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, True
    actualizarProgreso
    
    'sim_his_traza_gan -> sim_traza_gan
    tablaSH = "sim_his_traza_gan"
    cols = "pliqnro , concnro, empresa, fecha_pago, ternro, msr, nomsr, nogan, jubilacion, osocial, cuota_medico, prima_seguro, sepelio, estimados, otras, donacion, dedesp, noimpo, car_flia, conyuge, hijo, otras_cargas, retenciones, promo, saldo, sindicato, ret_mes, mon_conyuge, mon_hijo, mon_otras, viaticos, amortizacion, entidad1, entidad2, entidad3, entidad4, entidad5, entidad6, entidad7, entidad8, entidad9, entidad10, entidad11, entidad12, entidad13, entidad14, cuit_entidad1, cuit_entidad2, cuit_entidad3, cuit_entidad4, cuit_entidad5, cuit_entidad6, cuit_entidad7, cuit_entidad8, cuit_entidad9, cuit_entidad10, cuit_entidad11, cuit_entidad12, cuit_entidad13, cuit_entidad14, monto_entidad1, monto_entidad2, monto_entidad3, monto_entidad4, monto_entidad5, monto_entidad6, monto_entidad7, monto_entidad8 "
    cols = cols & ",monto_entidad9 , monto_entidad10, monto_entidad11, monto_entidad12, monto_entidad13, monto_entidad14, ganimpo, ganneta, total_entidad1, total_entidad2, total_entidad3, total_entidad4, total_entidad5, total_entidad6, total_entidad7, total_entidad8, total_entidad9, total_entidad10, total_entidad11, total_entidad12, total_entidad13, total_entidad14, Pronro, imp_deter, eme_medicas, seguro_optativo, seguro_retiro, tope_os_priv, empleg, deducciones, art23, porcdeduc"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
    
    'sim_his_traza_gan_item_top -> sim_traza_gan_item_top
    tablaSH = "sim_his_traza_gan_item_top"
    cols = "itenro, ternro, pronro, empresa, monto, ddjj, old_liq, liq, prorr"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
    
    'sim_his_vacpagdesc -> sim_vacpagdesc
    tablaSH = "sim_his_vacpagdesc"
    cols = "vacpdnro, pronro, cantdias, pliqnro, tprocnro, manual, pago_dto, ternro, concnro, emp_licnro, vacnro"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
    
    'sim_his_vales -> sim_vales
    tablaSH = "sim_his_vales"
    cols = "valnro, empleado, ppagnro, monnro, valmonto, valfecped, valfecprev, pliqnro, valdesc, pliqdto, pronro, tvalenro, valrevis, valautoriz, val_estnro, nroimp, fechaimp, valusuario, valaprosup"
    CopiarTablaHis tablaSH, simhisnro, cols, cond, False
    actualizarProgreso
    
End Sub

Private Sub borrarHistoricoCompleto(simhisnro As Long)
    
    borrarHistoricoSimHis (simhisnro)
    actualizarProgreso 40
    
    borrarTablaHis "sim_historico", simhisnro
    actualizarProgreso 80
    
End Sub

Private Sub borrarSimulacion()
    'borrarHistoricoSimHis (0)
    Flog.writeline "Eliminando datos."
    
    
    'sim_acu_liq
    Flog.writeline "Eliminando sim_acu_liq "
    StrSql = " DELETE FROM sim_acu_liq "
    StrSql = StrSql & " WHERE sim_acu_liq.cliqnro IN (" & SQLListaCabeceras & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_acu_mes
    Flog.writeline "Eliminando sim_acu_mes"
    StrSql = " DELETE FROM sim_acu_mes "
    StrSql = StrSql & " WHERE ((amanio >=" & AnioDesde & " AND ammes >=  " & MesDesde & " )"
    StrSql = StrSql & " AND(amanio <= " & AnioHasta & " AND ammes >= " & MesHasta & "))"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."

    'sim_detliq
    Flog.writeline "Eliminando sim_detliq "
    StrSql = " DELETE FROM sim_detliq "
    StrSql = StrSql & " WHERE sim_detliq.cliqnro IN (" & SQLListaCabeceras & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."

'    'sim_curva
'    Flog.writeline "Eliminando sim_curva "
'    StrSql = " DELETE FROM sim_curva "
'    'StrSql = StrSql & " WHERE "
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_DatosBaja
    Flog.writeline "Eliminando sim_DatosBaja "
    StrSql = " DELETE FROM sim_DatosBaja "
    StrSql = StrSql & " WHERE sim_DatosBaja.pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."

    'sim_desliq
    Flog.writeline "Eliminando sim_desliq "
    StrSql = " DELETE FROM sim_desliq "
    StrSql = StrSql & " WHERE sim_desliq.pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_desmen
    Flog.writeline "Eliminando sim_desmen "
    StrSql = " DELETE FROM sim_desmen "
    StrSql = StrSql & " WHERE sim_desmen.empleado IN (" & SQLListaTerceros & ")"
    StrSql = StrSql & " AND ((desfecdes >=  " & ConvFecha(FechaDesde) & " )"
    StrSql = StrSql & " AND (desfechas <= " & ConvFecha(FechaHasta) & "))"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
'    'sim_embcuota
'    Flog.writeline "Eliminando sim_embcuota "
'    StrSql = " DELETE FROM sim_embcuota "
'    StrSql = StrSql & " WHERE embnro IN ( SELECT embnro FROM sim_embargo WHERE sim_embargo.ternro IN (" & ListaTerceros & "))"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
'    'sim_embargo
'    Flog.writeline "Eliminando sim_embargo "
'    StrSql = " DELETE FROM sim_embargo "
'    StrSql = StrSql & " WHERE sim_embargo.ternro IN (" & ListaTerceros & ")"
'    'StrSql = StrSql & " WHERE sim_embargo.cliqnro IN (" & ListaCabeceras & ")"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_emp_lic
    Flog.writeline "Eliminando sim_emp_lic "
    StrSql = " DELETE FROM sim_emp_lic "
    StrSql = StrSql & " WHERE ((elfechadesde >=  " & ConvFecha(FechaDesde) & " )"
    StrSql = StrSql & " AND (elfechahasta <= " & ConvFecha(FechaHasta) & "))"
    StrSql = StrSql & " AND empleado IN (" & SQLListaTerceros & ") "
    'StrSql = StrSql & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."

'    'sim_emp_tikdist
'    Flog.writeline "Eliminando sim_emp_tikdist "
'    StrSql = " DELETE FROM sim_emp_tikdist "
'    StrSql = StrSql & " WHERE etiknro IN ( SELECT etiknro FROM sim_emp_ticket "
'    StrSql = StrSql & "     WHERE ((etikfecha >=  " & ConvFecha(FechaDesde) & " )"
'    StrSql = StrSql & "     AND (etikfecha <= " & ConvFecha(FechaHasta) & "))"
'    StrSql = StrSql & "     AND empleado IN (" & ListaTerceros & ") "
'    StrSql = StrSql & " )"
'    'StrSql = StrSql & " WHERE etiknro IN ( SELECT etiknro FROM sim_emp_ticket "
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
'    'sim_emp_ticket
'    Flog.writeline "Eliminando sim_emp_ticket "
'    StrSql = " DELETE FROM sim_emp_ticket "
'    StrSql = StrSql & " WHERE ((etikfecha >=  " & ConvFecha(FechaDesde) & " )"
'    StrSql = StrSql & " AND (etikfecha <= " & ConvFecha(FechaHasta) & "))"
'    StrSql = StrSql & " AND empleado IN (" & ListaTerceros & ") "
'    'StrSql = StrSql & " )"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_empleado
'    Flog.writeline "Eliminando sim_empleado "
'    StrSql = " DELETE FROM sim_empleado "
'    StrSql = StrSql & " WHERE ternro IN ( " & ListaTerceros & ")"
'    'objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros." 'not yet!!!

    'sim_fases
'    Flog.writeline "Eliminando sim_fases "
'    StrSql = " DELETE FROM sim_fases "
'    StrSql = StrSql & " WHERE ternro NOT IN ( " & ListaTerceros & ")"
'    'objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros." 'not yet
    
    'sim_ficharet
    Flog.writeline "Eliminando sim_ficharet "
    StrSql = " DELETE FROM sim_ficharet "
    StrSql = StrSql & " WHERE sim_ficharet.pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_gti_acunov
    Flog.writeline "Eliminando sim_gti_acunov "
    StrSql = " DELETE FROM sim_gti_acunov "
    StrSql = StrSql & " WHERE sim_gti_acunov.ternro IN (" & SQLListaTerceros & ")"
    StrSql = StrSql & " AND (sim_gti_acunov.pronro IS NULL OR sim_gti_acunov.pronro IN (" & SQLListaProcesos & "))"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_his_estructura
'    Flog.writeline "Eliminando sim_his_estructura "
'    StrSql = " DELETE FROM sim_his_estructura "
'    StrSql = StrSql & " WHERE sim_his_estructura.ternro IN (" & ListaTerceros & ")"
'    'objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros." 'not yet

    'sim_impgralarg
    Flog.writeline "Eliminando sim_impgralarg "
    StrSql = " DELETE FROM sim_impgralarg "
    StrSql = StrSql & " WHERE sim_impgralarg.pronro IN (" & SQLListaProcesos & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_impmesarg
    Flog.writeline "Eliminando sim_impmesarg "
    StrSql = " DELETE FROM sim_impmesarg "
    StrSql = StrSql & " WHERE sim_impmesarg.ternro IN (" & SQLListaTerceros & " )"
    StrSql = StrSql & " AND ((imaanio >=" & AnioDesde & " AND imames >=  " & MesDesde & " )"
    StrSql = StrSql & " AND(imaanio <= " & AnioHasta & " AND imames >= " & MesHasta & "))"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."

    'sim_impproarg
    Flog.writeline "Eliminando sim_impproarg "
    StrSql = " DELETE FROM sim_impproarg "
    StrSql = StrSql & " WHERE sim_impproarg.cliqnro IN (" & SQLListaCabeceras & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_novaju
    Flog.writeline "Eliminando sim_novaju "
    StrSql = " DELETE FROM sim_novaju "
    StrSql = StrSql & " WHERE sim_novaju.empleado IN (" & SQLListaTerceros & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_novemp -> sim_his_novemp
    Flog.writeline "Eliminando sim_novemp "
    StrSql = " DELETE FROM sim_novemp "
    StrSql = StrSql & " WHERE sim_novemp.empleado IN (" & SQLListaTerceros & " )"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
'    'sim_pre_cuota
'    Flog.writeline "Eliminando sim_pre_cuota "
'    StrSql = " DELETE FROM sim_pre_cuota "
'    StrSql = StrSql & " WHERE prenro IN ( SELECT prenro FROM sim_prestamo WHERE ternro IN ( " & ListaTerceros & " ) )"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
'    'sim_prestamo
'    Flog.writeline "Eliminando sim_prestamo "
'    StrSql = " DELETE FROM sim_prestamo "
'    StrSql = StrSql & " WHERE sim_prestamo.ternro IN (" & ListaTerceros & " )"
'    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
'    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    

    'sim_rep_borrdeta_det
    Flog.writeline "Eliminando sim_rep_borrdeta_det "
    StrSql = " DELETE FROM sim_rep_borrdeta_det "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_rep_borradordeta
    Flog.writeline "Eliminando sim_rep_borradordeta "
    StrSql = " DELETE FROM sim_rep_borradordeta "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."

    'sim_traza
    Flog.writeline "Eliminando sim_traza "
    StrSql = " DELETE FROM sim_traza "
    StrSql = StrSql & " WHERE cliqnro IN (" & SQLListaCabeceras & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_traza_gan
    Flog.writeline "Eliminando sim_traza_gan "
    StrSql = " DELETE FROM sim_his_traza_gan "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    StrSql = StrSql & " AND ternro IN (" & SQLListaTerceros & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_traza_gan_item_top
    Flog.writeline "Eliminando sim_traza_gan_item_top "
    StrSql = " DELETE FROM sim_traza_gan_item_top "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    StrSql = StrSql & " AND ternro IN (" & SQLListaTerceros & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_vacpagdesc
    Flog.writeline "Eliminando sim_vacpagdesc "
    StrSql = " DELETE FROM sim_vacpagdesc "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    StrSql = StrSql & " AND ternro IN (" & SQLListaTerceros & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    'sim_vales
    Flog.writeline "Eliminando sim_vales "
    StrSql = " DELETE FROM sim_vales "
    StrSql = StrSql & " WHERE empleado IN (" & SQLListaTerceros & ")"
    StrSql = StrSql & " AND pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."
    
    
    'sim_cabliq - se cambio el orden de las eliminaciones
    Flog.writeline "Eliminando sim_cabliq"
    StrSql = " DELETE FROM sim_cabliq "
    StrSql = StrSql & " WHERE sim_cabliq.cliqnro IN (" & SQLListaCabeceras & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."

    'sim_proceso - se cambio el orden de las eliminaciones
    Flog.writeline "Eliminando sim_his_proceso "
    StrSql = " DELETE FROM sim_proceso "
    StrSql = StrSql & " WHERE pronro IN (" & SQLListaProcesos & ")"
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se eliminaron " & AfectedRows & " registros."

    
End Sub


Private Sub borrarHistoricoSimHis(simhisnro As Long)
    
    'borrarTablaHis arregloTablas(UBound(arregloTablas)), simhisnro
    borrarTablaHis "sim_his_acu_liq", simhisnro
    borrarTablaHis "sim_his_acu_mes", simhisnro
    borrarTablaHis "sim_his_cabliq", simhisnro
    'borrarTablaHis "sim_his_curva", simhisnro
    borrarTablaHis "sim_his_DatosBaja", simhisnro
    borrarTablaHis "sim_his_desliq", simhisnro
    borrarTablaHis "sim_his_desmen", simhisnro
    borrarTablaHis "sim_his_detliq", simhisnro
    'borrarTablaHis "sim_his_embargo", simhisnro
    'borrarTablaHis "sim_his_embcuota", simhisnro
    borrarTablaHis "sim_his_emp_lic", simhisnro
    'borrarTablaHis "sim_his_emp_ticket", simhisnro
    'borrarTablaHis "sim_his_emp_tikdist", simhisnro
    borrarTablaHis "sim_his_empleado", simhisnro
    borrarTablaHis "sim_his_fases", simhisnro
    borrarTablaHis "sim_his_ficharet", simhisnro
    borrarTablaHis "sim_his_gti_acunov", simhisnro
    borrarTablaHis "sim_his_his_estructura", simhisnro
    borrarTablaHis "sim_his_impgralarg", simhisnro
    borrarTablaHis "sim_his_impmesarg", simhisnro
    borrarTablaHis "sim_his_impproarg", simhisnro
    borrarTablaHis "sim_his_novaju", simhisnro
    borrarTablaHis "sim_his_novemp", simhisnro
    'borrarTablaHis "sim_his_pre_cuota", simhisnro
    'borrarTablaHis "sim_his_prestamo", simhisnro
    borrarTablaHis "sim_his_proceso", simhisnro
    borrarTablaHis "sim_his_rep_borradordeta", simhisnro
    borrarTablaHis "sim_his_rep_borrdeta_det", simhisnro
    borrarTablaHis "sim_his_traza", simhisnro
    borrarTablaHis "sim_his_traza_gan", simhisnro
    borrarTablaHis "sim_his_traza_gan_item_top", simhisnro
    borrarTablaHis "sim_his_vacpagdesc", simhisnro
    borrarTablaHis "sim_his_vales", simhisnro
    'borrarTablaHis "sim_historico", simhisnro ' esta no
    
End Sub

Private Sub borrarTablaHis(tabla As String, simhisnro As Long)
    
    If simhisnro < 0 Then
        'ACA NO ENTRARIA NUNCA - BORRAR
    Else
        Flog.writeline "Borrando los datos de la tabla: " & tabla & " - historico:" & simhisnro
        
        StrSql = " DELETE FROM " & tabla & " WHERE simhisnro = " & simhisnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Datos Borrados de la tabla: " & tabla & " - historico:" & simhisnro
    End If
    
End Sub

Private Sub CopiarTablaHis(tabla As String, simhisnro As Long, Columnas As String, Condicion As String, Identity As Boolean)
    
    On Error GoTo errIdentityt
    
    'copio de las tablas de historico a simulacion
    Flog.writeline "Copiando los datos de la tabla: " & tabla & "  a la tabla : " & Replace(tabla, "sim_his_", "sim_") & " - historico:" & simhisnro
    
    If Identity Then
        identity_on Replace(tabla, "sim_his_", "sim_")
    End If
    
    StrSql = "INSERT INTO " & Replace(tabla, "sim_his_", "sim_")
    StrSql = StrSql & " (" & Columnas & ")"
    StrSql = StrSql & " SELECT distinct " & Columnas
    StrSql = StrSql & " FROM " & tabla
    StrSql = StrSql & " " & Condicion
    objConn.Execute StrSql, AfectedRows, adExecuteNoRecords
    Flog.writeline "Se actualizaron " & AfectedRows & " registros."
    
    If Identity Then
        identity_off Replace(tabla, "sim_his_", "sim_")
    End If
    'Flog.writeline "Datos Copiados "
        
    Exit Sub
    
errIdentityt:
    Flog.writeline "=================================================================="
    Flog.writeline "Error (sub Copiado): " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    If Identity Then
        identity_off Replace(tabla, "sim_his_", "sim_")
    End If
    
End Sub

Sub actualizarProgreso(Optional Valor As Double = 0)
    
    If Valor = 0 Then
        Progreso = Progreso + IncPasos
        If Progreso >= 99 Then
            Progreso = 99
        End If
    Else
        Progreso = Valor
    End If
    
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
End Sub
