Attribute VB_Name = "mdlBorrado"
'----------------------------------------------------------

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial
' Autor      : Sebastian Stremel
' Fecha      : 10/07/2012
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

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    
    'FGZ - 17/10/2012 -----------------------------------------------------------
    On Error Resume Next
    'Abro la conexion
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Then
        Nombre_Arch = PathFLog & "BorradoSimulaciones" & "-" & NroProcesoBatch & ".log"
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Then
        Nombre_Arch = PathFLog & "BorradoSimulaciones" & "-" & NroProcesoBatch & ".log"
        Set Flog = fs.CreateTextFile(Nombre_Arch, True)
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo ME_Gral
   
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 371 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    
    'Si no existe creo la carpeta donde voy a guardar el log
    If Not fs.FolderExists(PathFLog & rs_batch_proceso!iduser) Then fs.CreateFolder (PathFLog & rs_batch_proceso!iduser)
    Nombre_Arch = PathFLog & rs_batch_proceso!iduser & "\" & "BorradoSimulaciones" & "-" & NroProcesoBatch & ".log"
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    'FGZ - 17/10/2012 -----------------------------------------------------------
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
   
    'FGZ - 11/11/2011 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 371, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Fin
    End If
    'FGZ - 11/11/2011 --------- Control de versiones ------
    
    On Error GoTo ME_Main
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    'StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 371 AND bpronro =" & NroProcesoBatch
    'OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Borrado(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    Flog.writeline "---------------------------------------------------"
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        'FGZ - 30/12/2014 ----------------------------------
        'StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ",bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
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
    GoTo Fin
    
ME_Gral:

End Sub

Public Sub Borrado(NroProcesoBatch, bprcparam)

Dim sql As String
Dim cn As New Connection

'variables para levantar los parametros
Dim Parametros
Dim tipoBaja As Integer
Dim periodo As Integer
Dim modelo As Integer
Dim Proceso As Long
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim cliqnro As Long
Dim Ternro As Long

Dim anioPer As String
Dim mesPer As String
Dim borraNov As Integer

Parametros = Split(bprcparam, "@")

'levanto los parametros del filtro
Flog.writeline "Levanto los parametros del filtro"
tipoBaja = Parametros(0)
periodo = Parametros(1)
modelo = Parametros(2)
Proceso = Parametros(3)
borraNov = Parametros(4)

Flog.writeline "Parametros: tipoBaja:" & tipoBaja & " periodo:" & periodo & " modelo:" & modelo & " proceso:" & Proceso & " BorraNov:" & borraNov

'busco año y mes del periodo si periodo es <> 0
'If periodo <> "0" Then
'    StrSql = "SELECT pliqanio, pliqmes FROM periodo "
'    StrSql = StrSql & "WHERE pliqnro =" & periodo
'    OpenRecordset StrSql, rs
'    If Not rs.EOF Then
'        anioPer = rs!pliqanio
'        mesPer = rs!pliqmes
'        Flog.writeline "el anio del periodo es " & anioPer & " ,El mes del periodo es " & mesPer
'    Else
'        Flog.writeline "No se encontro periodo para el valor" & periodo
'    End If
'Else
'    Flog.writeline "Se elegieron todos los periodos"
    
'End If
'rs.Close


'StrSql = " SELECT * FROM sim_proceso "
'StrSql = StrSql & "WHERE protipoSim = " & tipoBaja
'If proceso <> "0" Then
'StrSql = StrSql & " AND pronro = " & proceso
'End If
'OpenRecordset StrSql, rs


'Do While Not rs.EOF
'    StrSql = " SELECT * FROM sim_cabliq "
    'StrSql = StrSql & "WHERE pronro= " & proceso
'    OpenRecordset StrSql, rs2
    'Do While Not rs2.EOF
    '    cliqnro = rs2!cliqnro
    '    ternro = rs2!Empleado
'objConn.Open
objConn.BeginTrans
Flog.writeline "Comenzo el proceso de borrado"

'borro los acumuladores de la liquidacion
StrSql = "DELETE FROM sim_acu_liq "
'l_sql = l_sql & " WHERE cliqnro=" & cliqnro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_acu_liq"


'borro los acumuladores mensuales
StrSql = "DELETE FROM sim_acu_mes "
'If anioPer <> "" Then
'    l_sql = l_slq & " WHERE ammes=" & mesPer & " AND amanio=" & anioPer & " AND ternro=" & ternro
'Else
'    l_sql = l_sql & " WHERE ternro=" & ternro
'End If
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_acu_mes"


'borro la cabecera de liquidacion
StrSql = "DELETE FROM sim_cabliq "
'WHERE cliqnro=" & cliqnro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_cabliq"
'    l_sql = "DELETE FROM Sim_DatosBaja"
'    cmExecute l_cm, l_sql, 0





'borro el desglose de liquidacion
StrSql = "DELETE FROM sim_desliq "
'l_sql = l_sql & " WHERE pronro=" & proceso & " AND empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_desliq"



'borro el desglose mensual
StrSql = "DELETE FROM sim_desmen "
'l_sql = l_sql & " WHERE desano=" & anioPer & " AND empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_desmen"


'borro el detalle de liquidacion
StrSql = "DELETE FROM sim_detliq "
'l_sql = l_sql & " WHERE cliqnro=" & cliqnro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_detliq"


'ACTUALIZO EL PROGRESO
'---------------------------------------
'Actualizo el progreso
IncPorc = 20
Progreso = Progreso + IncPorc
Flog.writeline "Progreso:" & Progreso
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
"' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords
'---------------------------------------


'borro los embargos
StrSql = "DELETE FROM sim_embargo "
'l_sql = l_sql & " WHERE ternro=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_embargo"


'borra las cuotas de los embargos
StrSql = "DELETE FROM sim_embcuota"
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_embcuota"




'borro las licencias del empleado
StrSql = "DELETE FROM sim_emp_lic"
'l_sql = "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_emp_lic"


'borro sim_emp_ticket
StrSql = "DELETE FROM sim_emp_ticket"
'l_sql = "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_emp_ticket"


'borro los sim_emp_tikdist
StrSql = "DELETE FROM sim_emp_tikdist "
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_emp_tikdist"


'FGZ - 30/09/2014 --------------
'borro todas las fases del empleado
StrSql = "DELETE FROM sim_fases_preaviso"
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_fases_preaviso"


'borro todas las fases del empleado
StrSql = "DELETE FROM sim_fases"
'l_sql = l_sql & "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_fases"



StrSql = "DELETE FROM sim_ficharet"
'l_sql = l_sql & "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_ficharet"


'ACTUALIZO EL PROGRESO
'---------------------------------------
'Actualizo el progreso
IncPorc = 20
Progreso = Progreso + IncPorc
Flog.writeline "Progreso:" & Progreso
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
"' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords
'---------------------------------------



StrSql = "DELETE FROM sim_gti_acunov"
'l_sql = l_sql & "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_gti_acunov"

StrSql = "DELETE FROM sim_his_estructura"
'l_sql = l_sql & "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_his_estructura"

StrSql = "DELETE FROM sim_impgralarg"
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_impgralarg"

StrSql = "DELETE FROM sim_impmesarg"
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_impmesarg"




StrSql = "DELETE FROM sim_impproarg"
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_impproarg"

StrSql = "DELETE FROM sim_novaju"
'l_sql = l_sql & "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_novaju"

StrSql = "DELETE FROM sim_novemp"
'l_sql = l_sql & "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_novemp"


'ACTUALIZO EL PROGRESO
'---------------------------------------
'Actualizo el progreso
IncPorc = 20
Progreso = Progreso + IncPorc
Flog.writeline "Progreso:" & Progreso
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
"' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords
'---------------------------------------

StrSql = "DELETE FROM sim_pre_cuota"
'l_sql = l_sql & "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_pre_cuota"

StrSql = "DELETE FROM sim_prestamo"
'l_sql = l_sql & "WHERE empleado=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_prestamo"

StrSql = "DELETE FROM sim_proceso"
'If proceso <> "0" Then
'    l_sql = l_sql & "WHERE pronro=" & proceso
'    If periodo <> "0" Then
'        l_sql = l_sql & " AND pliqnro=" & periodo
'    End If
'End If
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_proceso"

StrSql = "DELETE FROM sim_traza"
'If cliqnro <> "" Then
'    l_sql = l_sql & "WHERE cliqnro=" & cliqnro
'End If
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_traza"

StrSql = "DELETE FROM sim_traza_gan"
'If pliqnro <> "0" Then
'    l_sql = l_sql & " WHERE pliqnro=" & periodo
'    l_sql = l_sql & " AND ternro=" & ternro
'Else
'    l_sql = l_sql & " WHERE ternro=" & ternro
'End If
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_traza_gan"


StrSql = "DELETE FROM sim_traza_gan_item_top"
'If proceso <> "0" Then
'    l_sql = l_sql & " WHERE pronro=" & proceso
'End If
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_traza_gan_item_top"


StrSql = "DELETE FROM sim_vacpagdesc"
'If proceso <> "0" Then
'    l_sql = l_sql & " WHERE pronro=" & proceso
'    If periodo <> "0" Then
'        l_sql = l_sql & " AND pliqnro=" & periodo
'        l_sql = l_sql & " AND ternro=" & ternro
'    Else
'        l_sql = l_sql & " AND ternro=" & ternro
'    End If
'Else
'    l_sql = l_sql & " WHERE ternro=" & ternro
'End If
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_vacpagdesc"


'ACTUALIZO EL PROGRESO
'---------------------------------------
'Actualizo el progreso
IncPorc = 20
Progreso = Progreso + IncPorc
Flog.writeline "Progreso:" & Progreso
TiempoAcumulado = GetTickCount
StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
"' WHERE bpronro = " & NroProcesoBatch
objconnProgreso.Execute StrSql, , adExecuteNoRecords
'---------------------------------------


StrSql = "DELETE FROM sim_vales"
'l_sql = l_sql & " WHERE ternro=" & ternro
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_vales"

StrSql = "DELETE FROM sim_rep_borrdeta_det "
Flog.writeline "Se borraron los datos de la tabla sim_rep_borrdeta_det"

StrSql = "DELETE FROM sim_rep_borradordeta "
Flog.writeline "Se borraron los datos de la tabla sim_rep_borradordeta"

'si es GP borro el tipo
If tipoBaja = "1" Then
    StrSql = "DELETE FROM ter_tip WHERE tipnro = 26 "
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Se borraron los terceros de tipo 26 de la tabla ter_tip"
Else
    Flog.writeline "No Se borraron los terceros de tipo 26 de la tabla ter_tip"
End If

If borraNov = "-1" Then
    StrSql = "DELETE FROM novretro "
    StrSql = StrSql & " WHERE nretronro IN ( "
    StrSql = StrSql & " SELECT n.nretronro FROM novretro n INNER JOIN proceso p ON p.pronro = n.pronro "
    StrSql = StrSql & " WHERE p.proaprob = 0)"
    Flog.writeline "Se borraron las novedades retroactivas de la tabla novretro"
Else
    Flog.writeline "No Se borraron las novedades retroactivas de la tabla novretro"
End If



'EAM - 04/03/2013 ----------------------------------------
' Venta de Vacaciones (Sykes) a sim_vacvendidos
StrSql = "DELETE FROM sim_vacvendidos"
objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Se borraron los datos de la tabla sim_vacvendidos"
       


objConn.CommitTrans
Flog.writeline "Termino el proceso con exito"
        
'rs2.MoveNext
'Loop
'rs.MoveNext
'Loop







'
'     l_sql = "DELETE FROM sim_empleado"
'    cmExecute l_cm, l_sql, 0
'


'
'    l_sql = "DELETE FROM sim_novemp"
'    cmExecute l_cm, l_sql, 0
'
'    l_sql = "DELETE FROM sim_pre_cuota"
'    cmExecute l_cm, l_sql, 0
'
'    l_sql = "DELETE FROM sim_prestamo"
'    cmExecute l_cm, l_sql, 0
'
'    l_sql = "DELETE FROM sim_proceso"
'    cmExecute l_cm, l_sql, 0
'
'    l_sql = "DELETE FROM sim_traza"
'    cmExecute l_cm, l_sql, 0
'
'    l_sql = "DELETE FROM sim_traza_gan"
'    cmExecute l_cm, l_sql, 0
'
'    l_sql = "DELETE FROM sim_traza_gan_item_top"
'    cmExecute l_cm, l_sql, 0
'
'    l_sql = "DELETE FROM sim_vacpagdesc"
'    cmExecute l_cm, l_sql, 0
'
'    l_sql = "DELETE FROM sim_vales"
'    cmExecute l_cm, l_sql, 0

End Sub
