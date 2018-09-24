Attribute VB_Name = "mdlRepNov"
Option Explicit

'Const Version = "1.00"
'Const FechaVersion = "04/07/2008"
''Modificaciones: CS
''    Version Inicial, con varios cambios del esatandar


'Const Version = "1.01"
'Const FechaVersion = "27/08/2008"
'Modificaciones: FGZ
'    Se agregó la posibilidad de buscar en el Acumulado Diario

'Const Version = "1.02"
'Const FechaVersion = "07/10/2008"
'Modificaciones: Cesar Stankunas
'    Se modificó la grabación del usuario en BDD

Const Version = "1.03"
Const FechaVersion = "31/07/2009"
'Modificaciones: Martin Ferraro - Encriptacion de string connection
'                                 Se cambio el formato del nombre del log

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

'Public Type TConfrep
'    Nrocol As Integer
'    Tipo As String
'    val1 As tring
'    val2 As tring
'    Accion As String
'End Type

Dim fs
Dim Flog
Dim FDesde As Date
Dim FHasta As Date

Dim sep As String
Dim objBTurno As New BuscarTurno

Global tiene_turno As Boolean
Global nro_turno As Long
Global Tipo_Turno As Integer
Global Tiene_Justif As Boolean
Global nro_justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global nro_grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global HuboErrores As Boolean
Global Usuario As String
Global FechaDesde As Date
Global FechaHasta As Date

Sub Main()
Dim Archivo As String
Dim pos As Integer
Dim strcmdLine  As String

'Dim objconnMain As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim Fecha As Date
Dim Hora As String
Dim NroProceso As Long
Dim NroReporte As Long
Dim StrParametros As String

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim ArrParametros

'    strcmdLine = Command()
'    ArrParametros = Split(strcmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProceso = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strcmdLine) Then
'            NroProceso = strcmdLine
'        Else
'            Exit Sub
'        End If
'    End If
    
    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
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
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'Creo el archivo de texto del desglose
    'Archivo = PathFLog & "RepNovedades-" & CStr(NroProceso) & Format(Now, "DD-MM-YYYY") & ".log"
    Archivo = PathFLog & "RepNovedades-" & CStr(NroProceso) & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)

    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If

    On Error GoTo CE

    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Flog.writeline Espacios(Tabulador * 0) & "Levanta Proceso y Setea Parámetros:  " & " " & Now
    
    'levanto los parametros del proceso
    StrParametros = ""
    sep = "@"
    StrSql = "SELECT bprcfecdesde,bprcfechasta,bprcparam,bprcfecha,bprchora,iduser  FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Fecha = rs!bprcfecha
        Hora = rs!bprchora
        Usuario = rs!iduser
        If Not IsNull(rs!bprcparam) Then
            If Len(rs!bprcparam) >= 1 Then
                pos = InStr(1, rs!bprcparam, sep)
                NroReporte = CLng(Left(rs!bprcparam, pos - 1))
                StrParametros = Right(rs!bprcparam, Len(rs!bprcparam) - (pos))
            End If
        End If
    Else
        Exit Sub
    End If
    
    depurar = True
    
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "Inicio de Reporte de Novedades: " & " " & Now
    End If
    Call Reporte_01(NroReporte, NroProceso, StrParametros, Fecha, Hora)
    If depurar Then
        Flog.writeline Espacios(Tabulador * 0) & "Fin de Reporte de Novedades: " & " " & Now
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & "copio al historico" & " " & Now
    End If
    
    
    'Actualizo el Btach_Proceso
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' -----------------------------------------------------------------------------------
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso"
        Flog.writeline
    End If
    
    If Not HuboErrores Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "---> Proceso teminado, paso al historico ... " & Now
        End If

        'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso

        
        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
        ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!iduser & "'"
        
        If Not IsNull(rs_Batch_Proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchora & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfechasta)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcprogreso
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecfin)
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!empnro
        End If
        If Not IsNull(rs_Batch_Proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcPid
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcUrgente
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "---> Historico Actualizado " & Now
        End If
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    End If
    Flog.Close
    
    'Cierro y libero todo
    If TransactionRunning Then MyRollbackTrans
    
    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    If CnTraza.State = adStateOpen Then CnTraza.Close
Exit Sub

CE:
    Flog.writeline "Reporte abortado por Error:" & " " & Now
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
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    
End Sub

Private Sub Reporte_01(NroReporte As Long, NroProceso As Long, Parametros As String, Fecha As Date, Hora As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que calcula las novedades.
' Autor      : CS
' Fecha      :
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim i As Integer

Dim Por_Usuario As Boolean
Dim l_apeynom As String
Dim l_doc As String
Dim l_est1 As String
Dim l_est2 As String
Dim l_est3 As String
Dim l_estn1 As String
Dim l_estn2 As String
Dim l_leg As String
Dim TE1 As Long
Dim TE2 As Long
Dim TE3 As Long
Dim Aux_Fecha As Date

Dim l_nrocol As Long
Dim l_tipo As String
Dim l_val1 As Long
Dim l_val2 As String
Dim l_accion As String

Dim l_col1 As Double
Dim l_col2 As Double
Dim l_col3 As Double
Dim l_col4 As Double
Dim l_col5 As Double
Dim l_col6 As Double
Dim l_col7 As Double
Dim l_col8 As Double
Dim l_col9 As Double
Dim l_col10 As Double
Dim l_col11 As Double
Dim l_col12 As Double
Dim l_col13 As Double
Dim l_col14 As Double
Dim l_col15 As Double
Dim l_col16 As Double
Dim l_col17 As Double
Dim l_col18 As Double
Dim l_col19 As Double
Dim l_col20 As Double

Dim l_val As Double

' se supone que estos son parametros de entrada y vienen en "parametros"
Dim l_id As String
Dim l_nivel1 As String
Dim l_nivel2 As String
Dim l_pergtinro As String
Dim l_prcgtinro As String
Dim l_liqresp As String
Dim l_orden As String
Dim l_repnovnro As Long

Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Integer
Dim IncPorc As Single
Dim Progreso As Single

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Doc As New ADODB.Recordset


On Error GoTo ME_Local

' ------------------------------------
'levanto cada parametro por separado, el separador de parametros es "@"
sep = "@"

If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        
        pos1 = 1
        pos2 = InStr(pos1, Parametros, sep) - 1
        l_id = Mid(Parametros, pos1, pos2)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, sep) - 1
        l_nivel1 = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, sep) - 1
        l_nivel2 = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, sep) - 1
        l_pergtinro = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, sep) - 1
        l_prcgtinro = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, sep) - 1
        l_liqresp = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        l_orden = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
    End If
End If


'Busco la fecha hasta del periodo de gti
StrSql = "SELECT pgtidesde, pgtihasta FROM gti_per WHERE pgtinro = " & l_pergtinro
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Aux_Fecha = rs!pgtihasta
    FechaDesde = rs!pgtidesde
    FechaHasta = rs!pgtihasta
End If

'Busco las descripciones de las estructuras
StrSql = "SELECT estrdabr FROM estructura WHERE estrnro = " & l_nivel1
OpenRecordset StrSql, rs
If Not rs.EOF Then
    l_estn1 = rs!estrdabr
Else
    l_estn1 = ""
End If
StrSql = "SELECT estrdabr FROM estructura WHERE estrnro = " & l_nivel2
OpenRecordset StrSql, rs
If Not rs.EOF Then
    l_estn2 = rs!estrdabr
Else
    l_estn2 = ""
End If


'Para el empleado actual ciclo entre las columnas configuradas en el rep y guardo los valores en el detalle
'Busco si el reporte se configura por usuario
StrSql = "SELECT repagr FROM reporte WHERE repnro = 231"
OpenRecordset StrSql, rs3
If Not rs3.EOF Then
    If CBool(rs3!repagr) Then
        Por_Usuario = True
    Else
        Por_Usuario = False
    End If
Else
    Por_Usuario = False
End If


'Busco las estructuras configuradas en el confrep
StrSql = " SELECT confnrocol, conftipo, confval "
StrSql = StrSql & " FROM confrep "
StrSql = StrSql & " WHERE repnro = " & NroReporte
StrSql = StrSql & " AND confnrocol >= 1 and confnrocol <= 3"
If Por_Usuario Then
    StrSql = StrSql & " AND iduser = '" & Usuario & "'"
Else
    StrSql = StrSql & " AND (iduser = '' OR iduser IS NULL )"
End If
OpenRecordset StrSql, rs3
TE1 = 0
TE2 = 0
TE3 = 0
Do While Not rs3.EOF
    Select Case rs3!confnrocol
    Case 1:
        TE1 = rs3!confval
    Case 2:
        TE2 = rs3!confval
    Case 3:
        TE3 = rs3!confval
    End Select
    rs3.MoveNext
Loop


'Guardo en la tabla rep_novgti (cabecera) los datos principales
StrSql = "INSERT INTO rep_novgti (bpronro,iduser,liqresp,nivel1,nivel2,estn1,estn2,pergtinro,procgti,fecha,hora) VALUES ("
StrSql = StrSql & NroProceso & ","
' Modificación: 07/10/2008 - Cesar Stankunas - Se agregó la verif. por Usuario
If Por_Usuario Then
    StrSql = StrSql & "'" & Usuario & "',"
Else
    StrSql = StrSql & "'',"
End If
StrSql = StrSql & "'" & l_liqresp & "',"
StrSql = StrSql & l_nivel1 & ","
StrSql = StrSql & l_nivel2 & ","
StrSql = StrSql & "'" & l_estn1 & "',"
StrSql = StrSql & "'" & l_estn2 & "',"
StrSql = StrSql & l_pergtinro & ","
StrSql = StrSql & "'" & l_prcgtinro & "',"
StrSql = StrSql & ConvFecha(Fecha) & ","
StrSql = StrSql & "'" & Hora & "')"
objConn.Execute StrSql, , adExecuteNoRecords


'Busco el repnovnro insertado
l_repnovnro = getLastIdentity(objConn, "rep_novgti")


'Levanto todos los empleados a Procesar
StrSql = "SELECT batch_empleado.ternro, empleado.empleg, tercero.terape, tercero.ternom, tercero.terape2, tercero.ternom2 "
StrSql = StrSql & " FROM batch_empleado "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = empleado.ternro "
StrSql = StrSql & " WHERE batch_empleado.bpronro = " & NroProceso
OpenRecordset StrSql, rs2

'Seteo las variables de progreso
Progreso = 0
If Not rs2.EOF Then
    CEmpleadosAProc = rs2.RecordCount
    IncPorc = ((100 / CEmpleadosAProc))
Else
    If depurar Then
        Flog.writeline Espacios(Tabulador * 1) & "No hay empleados asociados al proceso" & " " & Now
    End If
    IncPorc = 100
    Exit Sub
End If

'Levanto la configuracion del confrep
StrSql = " SELECT confnrocol, conftipo, confval, confval2, confaccion "
StrSql = StrSql & " FROM confrep "
StrSql = StrSql & " WHERE repnro = " & NroReporte
StrSql = StrSql & " AND conftipo <> 'FIL' "
StrSql = StrSql & " AND conftipo <> 'TE' "
If Por_Usuario Then
    StrSql = StrSql & " AND iduser = '" & Usuario & "'"
Else
    StrSql = StrSql & " AND (iduser = '' OR iduser IS NULL )"
End If
OpenRecordset StrSql, rs3
If rs3.EOF Then
    If depurar Then
        Flog.writeline Espacios(Tabulador * 1) & "No hay ninguna columna configurada para el Reporte 231. " & " " & Now
    End If
    Exit Sub
End If

'Por cada empleado
Do While Not rs2.EOF
    'Inicializo las columnas
    l_col4 = 0
    l_col5 = 0
    l_col6 = 0
    l_col7 = 0
    l_col8 = 0
    l_col9 = 0
    l_col10 = 0
    l_col11 = 0
    l_col12 = 0
    l_col13 = 0
    l_col14 = 0
    l_col15 = 0
    l_col16 = 0
    l_col17 = 0
    l_col18 = 0
    l_col19 = 0
    l_col20 = 0
    
    rs3.MoveFirst
    Do While Not rs3.EOF
        l_nrocol = rs3("confnrocol")
        l_tipo = rs3("conftipo")
        l_val1 = rs3("confval")
        l_val2 = ""
        If Not EsNulo(rs3("confval2")) Then
            l_val2 = rs3("confval2")
        End If
        l_accion = rs3("confaccion")
        'llamo a la subrutina 'CalcularColumna' para que calcule el valor de la columna actual
        Call CalcularColumna(rs2("ternro"), l_tipo, l_val1, l_val2, l_accion, l_prcgtinro, l_val)
        Select Case l_nrocol
        Case 4
            l_col4 = l_col4 + l_val
        Case 5
            l_col5 = l_col5 + l_val
        Case 6
            l_col6 = l_col6 + l_val
        Case 7
            l_col7 = l_col7 + l_val
        Case 8
            l_col8 = l_col8 + l_val
        Case 9
            l_col9 = l_col9 + l_val
        Case 10
            l_col10 = l_col10 + l_val
        Case 11
            l_col11 = l_col11 + l_val
        Case 12
            l_col12 = l_col12 + l_val
        Case 13
            l_col13 = l_col13 + l_val
        Case 14
            l_col14 = l_col14 + l_val
        Case 15
            l_col15 = l_col15 + l_val
        Case 16
            l_col16 = l_col16 + l_val
        Case 17
            l_col17 = l_col17 + l_val
        Case 18
            l_col18 = l_col18 + l_val
        Case 19
            l_col19 = l_col19 + l_val
        Case 20
            l_col20 = l_col20 + l_val
        End Select
        rs3.MoveNext
    Loop

    'Busco las estructuras correspondientes a cada tipo de Estructura
    l_est1 = ""
    l_est2 = ""
    l_est3 = ""
    
    StrSql = " SELECT estrdabr FROM his_estructura "
    StrSql = StrSql & "INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " WHERE his_estructura.ternro = " & rs2!Ternro
    StrSql = StrSql & " AND his_estructura.tenro = " & TE1
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        l_est1 = rs_Estructura!estrdabr
    End If
    
    StrSql = " SELECT estrdabr FROM his_estructura "
    StrSql = StrSql & "INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " WHERE his_estructura.ternro = " & rs2!Ternro
    StrSql = StrSql & " AND his_estructura.tenro = " & TE2
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        l_est2 = rs_Estructura!estrdabr
    End If
        
    StrSql = " SELECT estrdabr FROM his_estructura "
    StrSql = StrSql & "INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro"
    StrSql = StrSql & " WHERE his_estructura.ternro = " & rs2!Ternro
    StrSql = StrSql & " AND his_estructura.tenro = " & TE3
    StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ")"
    StrSql = StrSql & " AND ((" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rs_Estructura
    If Not rs_Estructura.EOF Then
        l_est3 = rs_Estructura!estrdabr
    End If

    l_apeynom = rs2("ternom") & " " & rs2("ternom2") & " " & rs2("terape") & " " & rs2("terape2")
    l_apeynom = Trim(l_apeynom)
    
    'Documento
    StrSql = "SELECT tipodocu.tidsigla, ter_doc.nrodoc "
    StrSql = StrSql & " FROM ter_doc"
    StrSql = StrSql & " LEFT JOIN tipodocu ON ter_doc.tidnro = tipodocu.tidnro "
    StrSql = StrSql & " WHERE ter_doc.tidnro <= 4 AND ternro = " & rs2!Ternro
    OpenRecordset StrSql, rs_Doc
    If Not rs_Doc.EOF Then
        If Not EsNulo(rs_Doc("tidsigla")) Then
            l_doc = rs_Doc("tidsigla")
        Else
            l_doc = "???"
        End If
    
        If Not EsNulo(rs_Doc("nrodoc")) Then
            l_doc = l_doc & " " & Trim(rs_Doc("nrodoc"))
        End If
    Else
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "No se encontro documento basico para el empleado" & " " & Now
        End If
        l_doc = "???"
    End If
       
    l_leg = rs2!empleg
    
    'inserto los campos en el detalle del reporte (rep_novgti_det)
    StrSql = "INSERT INTO rep_novgti_det (repnovnro,bpronro,Ternro,legajo,apeynom,est1,est2,est3,dni,col1,col2,col3,col4,col5,col6,col7,col8,col9,col10,col11,col12,col13,col14,col15,col16,col17,col18,col19,col20) VALUES ("
    StrSql = StrSql & l_repnovnro & ","
    StrSql = StrSql & NroProceso & ","
    StrSql = StrSql & rs2("ternro") & ","
    StrSql = StrSql & rs2("empleg") & ","
    StrSql = StrSql & "'" & l_apeynom & "',"
    StrSql = StrSql & "'" & l_est1 & "',"
    StrSql = StrSql & "'" & l_est2 & "',"
    StrSql = StrSql & "'" & l_est3 & "',"
    StrSql = StrSql & "'" & l_doc & "',"
    StrSql = StrSql & "'" & l_col1 & "',"
    StrSql = StrSql & "'" & l_col2 & "',"
    StrSql = StrSql & "'" & l_col3 & "',"
    StrSql = StrSql & "'" & l_col4 & "',"
    StrSql = StrSql & "'" & l_col5 & "',"
    StrSql = StrSql & "'" & l_col6 & "',"
    StrSql = StrSql & "'" & l_col7 & "',"
    StrSql = StrSql & "'" & l_col8 & "',"
    StrSql = StrSql & "'" & l_col9 & "',"
    StrSql = StrSql & "'" & l_col10 & "',"
    StrSql = StrSql & "'" & l_col11 & "',"
    StrSql = StrSql & "'" & l_col12 & "',"
    StrSql = StrSql & "'" & l_col13 & "',"
    StrSql = StrSql & "'" & l_col14 & "',"
    StrSql = StrSql & "'" & l_col15 & "',"
    StrSql = StrSql & "'" & l_col16 & "',"
    StrSql = StrSql & "'" & l_col17 & "',"
    StrSql = StrSql & "'" & l_col18 & "',"
    StrSql = StrSql & "'" & l_col19 & "',"
    StrSql = StrSql & "'" & l_col20 & "')"
    objConn.Execute StrSql, adExecuteNoRecords
    
    If depurar Then
        Flog.writeline Espacios(Tabulador * 1) & "inserta el empleado: " & l_apeynom
    End If
    
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    rs2.MoveNext
Loop
    
Fin:
'Cierro y libero
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing
If rs2.State = adStateOpen Then rs2.Close
Set rs2 = Nothing
Exit Sub

ME_Local:
    HuboErrores = True
    Flog.writeline Espacios(Tabulador * 1) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 1) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


Private Sub BuscarHorasAusencia(Fecha As Date, Ternro As Long, ByRef Conjunto As String)
Dim rs As New ADODB.Recordset
Dim i As Integer


    ' seteo el turno y el dia
    'Set objBTurno.Conexion = objConn
    'objBTurno.Buscar_Turno Fecha, Ternro, False
    'initVariablesTurno objBTurno
    
    If Not tiene_turno Then
        
        ' -----------------------------------------------------------------
        ' si no tiene turno no puedo buscar en GTI_Config_tur_hor el tipo de hora ausencia
        ' la pregunta es: ¿Qué deberia retornar en el conjunto de horas de Ausencia ????
        ' -----------------------------------------------------------------
        Conjunto = "0"
        Exit Sub
    End If
    
    If tiene_turno Then
        StrSql = "SELECT gti_config_tur_hor.thnro as TipoHora FROM gti_config_hora " & _
                " INNER JOIN gti_config_tur_hor ON gti_config_hora.conhornro = gti_config_tur_hor.conhornro " & _
                " WHERE gti_config_hora.conhornro = 2 AND gti_config_tur_hor.turnro =" & nro_turno
        OpenRecordset StrSql, rs
    End If
    
    Conjunto = ""
    i = 1
    Do While Not rs.EOF
        If i = 1 Then
            Conjunto = rs("TipoHora")
            i = i + 1
        Else
            Conjunto = Conjunto & "," & rs("TipoHora")
        End If
        rs.MoveNext
    Loop
    
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub



Private Sub BuscarTodasLasHorasAusencia(ByRef Conjunto As String)
Dim rs As New ADODB.Recordset
Dim i As Integer
Dim Fecha As Date

' StrSql = "SELECT DISTINCT gti_config_tur_hor.thnro as TipoHora FROM gti_config_hora " & _
'          " INNER JOIN gti_config_tur_hor ON gti_config_hora.conhornro = gti_config_tur_hor.conhornro " & _
'          " WHERE gti_config_hora.conhornro = 2"
  StrSql = "SELECT confrep.confval as TipoHora " & _
           "FROM   confrep " & _
           "WHERE  confrep.repnro = 54"
  OpenRecordset StrSql, rs
' Cambiado para mostrar los tipos de horas convertidas a unidades 'jornadas'
' O.D.A. 29/03/2004
    
  If rs.EOF Then
    Conjunto = "0"
  End If
    
  i = 1
  Do While Not rs.EOF
    If i = 1 Then
      Conjunto = rs("TipoHora")
      i = i + 1
    Else
      Conjunto = Conjunto & "," & rs("TipoHora")
    End If
    rs.MoveNext
  Loop
    
  If rs.State = adStateOpen Then rs.Close
  Set rs = Nothing
End Sub



Private Sub initVariablesTurno(ByRef T As BuscarTurno)
   p_turcomp = T.Compensa_Turno
   nro_grupo = T.Empleado_Grupo
   nro_justif = T.Justif_Numero
   justif_turno = T.justif_turno
   Tiene_Justif = T.Tiene_Justif
   Fecha_Inicio = T.FechaInicio
   Nro_fpgo = T.Numero_FPago
   nro_turno = T.Turno_Numero
   tiene_turno = T.tiene_turno
   Tipo_Turno = T.Turno_Tipo
   P_Asignacion = T.Tiene_PAsignacion
End Sub



Private Sub CalcularColumna(Tercero As Long, Tipo As String, val1 As Long, val2 As String, Accion As String, procesos As String, columna As Double)
    Dim StrSql2 As String
    Dim l_rs4 As New ADODB.Recordset
    Dim l_rs3 As New ADODB.Recordset
    Select Case Tipo
    Case "TH"   'Tipo de Hora
        StrSql = "SELECT SUM(adcanthoras) as Sumahoras FROM  gti_acumdiario "
        StrSql = StrSql & " WHERE gti_acumdiario.ternro = " & Tercero
        StrSql = StrSql & " AND ( " & ConvFecha(FechaDesde) & " <= gti_acumdiario.adfecha  AND gti_acumdiario.adfecha <= " & ConvFecha(FechaHasta) & ")"
        StrSql = StrSql & " AND gti_acumdiario.thnro = " & val1
        OpenRecordset StrSql, l_rs3
        If Not l_rs3.EOF Then
            If Not EsNulo(l_rs3("Sumahoras")) Then
                columna = CStr(l_rs3("Sumahoras"))
            Else
                columna = "0"
            End If
        Else
            columna = "0"
        End If
    Case "GCO"
        StrSql = "SELECT acnovvalor FROM concepto "
        StrSql = StrSql & " INNER JOIN gti_acunov ON gti_acunov.concnro = concepto.concnro "
        StrSql = StrSql & " WHERE concepto.conccod = " & val1
        StrSql = StrSql & " AND gti_acunov.tpanro = " & CLng(val2)
        StrSql = StrSql & " AND gti_acunov.ternro = " & Tercero
        StrSql = StrSql & " AND gti_acunov.gpanro in (" & procesos & ")"
        StrSql = StrSql & " AND gti_acunov.acnovfecaprob IS NOT NULL"
        OpenRecordset StrSql, l_rs3
        If Not l_rs3.EOF Then
            columna = CStr(l_rs3("acnovvalor"))
        Else
            columna = "0"
        End If
    Case "LCO"
        StrSql = "SELECT nevalor FROM concepto "
        StrSql = StrSql & " INNER JOIN novemp ON novemp.concnro = concepto.concnro "
        StrSql = StrSql & " WHERE concepto.conccod = " & val1
        StrSql = StrSql & " AND novemp.tpanro = " & CLng(val2)
        StrSql = StrSql & " AND novemp.empleado = " & Tercero
        
        StrSql2 = "SELECT gpadesde, gpahasta FROM gti_procacum"
        StrSql2 = StrSql2 & " WHERE gpanro in (" & procesos & ")"
        OpenRecordset StrSql2, l_rs4
        If Not l_rs4.EOF Then
            StrSql = StrSql & " AND ("
            StrSql = StrSql & "(novemp.nedesde >= '" & l_rs4("gpadesde") & "' AND novemp.nehasta <= '" & l_rs4("gpahasta") & "')"
            l_rs4.MoveNext
            Do While Not l_rs4.EOF
                StrSql = StrSql & "OR (novemp.nedesde >= '" & l_rs4("gpadesde") & "' AND novemp.nehasta <= '" & l_rs4("gpahasta") & "')"
                l_rs4.MoveNext
            Loop
            StrSql = StrSql & " )"
        End If
        OpenRecordset StrSql, l_rs3
        If Not l_rs3.EOF Then
            columna = CStr(l_rs3("nevalor"))
        Else
            columna = "0"
        End If
    End Select
    
'cerrar y liberar
If l_rs3.State = adStateOpen Then l_rs3.Close
If l_rs4.State = adStateOpen Then l_rs4.Close

Set l_rs3 = Nothing
Set l_rs4 = Nothing
End Sub
