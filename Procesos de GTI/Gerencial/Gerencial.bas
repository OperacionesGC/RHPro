Attribute VB_Name = "mdlGerencial"
Option Explicit


'============================================================================================
'Const Version = "3.00"
'Const FechaVersion = "11/06/2007"
'Modificaciones: FGZ
'      Versión Inicial


Const Version = "3.01"
Const FechaVersion = "31/07/2009"
'Modificaciones: MB - Encriptacion de string connection

'============================================================================================
'Para Sql server
'Global Const strConexionNexus = "DSN=Nexushr-RHPro;database=nexus;uid=sa;pwd="

Global Flog
Global fs
Global HuboError As Boolean
Global HuboErrorTipo As Boolean

'Global objRs As New ADODB.Recordset
Global rsEmp As New ADODB.Recordset
Global rsAnrCab As New ADODB.Recordset
Global rsHistliq As New ADODB.Recordset
Global rsFactor As New ADODB.Recordset
Global rsFactorTotalizador As New ADODB.Recordset
Global rsHistCon As New ADODB.Recordset
Global rsEstructura As New ADODB.Recordset
Global rsRango As New ADODB.Recordset
Global rsConc As New ADODB.Recordset
Global rsAcumDiario As New ADODB.Recordset
Global rsFiltro As New ADODB.Recordset
Global rsTot As New ADODB.Recordset

Global CantFactor As Integer
Global CantFiltro As Integer
Global CantRango As Integer
Global PorcTiempo As Double
Global SumPorcTiempo As Double


Global IncPorc As Single
Global Progreso As Single
Global NroProceso As Long

Global FactorTotalizador As Long
Global Totaliza As Boolean
Global Etiqueta
Global Cantidad_de_OpenRecordset As Long
Global Cantidad_Call_Politicas As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long



Public Sub Main()
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Fecha As Date
Dim objRs As New ADODB.Recordset
Dim objrsEmpleado As New ADODB.Recordset
Dim strCmdLine  As String

Dim Nombre_Arch As String
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim bprcparam As String

Dim PID As String
Dim ArrParametros

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProceso = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProceso = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If

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

    Cantidad_de_OpenRecordset = 0
    Cantidad_Call_Politicas = 0

    'Abro la conexion
'    On Error Resume Next
'    OpenConnection strconexion, objConn
'    If Err.Number <> 0 Then
'        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
'        Exit Sub
'    End If
'    On Error Resume Next
'    OpenConnection strconexion, objConnProgreso
'    If Err.Number <> 0 Then
'        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
'        Exit Sub
'    End If

On Error GoTo CE

    Nombre_Arch = PathFLog & "MIG" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo CE

    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 18 AND bpronro =" & NroProceso
    OpenRecordset StrSql, objRs
    
    HuboError = False
    If Not objRs.EOF Then
        bprcparam = objRs!bprcparam
        objRs.Close
        Set objRs = Nothing
        Call LevantarParamteros(NroProceso, bprcparam)
    End If
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
        
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso = 100, bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    objConnProgreso.Execute StrSql, , adExecuteNoRecords
    
    ' -----------------------------------------------------------------------------------
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
    If Not HuboError Then
        MyBeginTrans
    
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso
        If Not rs_Batch_Proceso.EOF Then
        
            StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
            StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
                     ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!IdUser & "'"
            
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
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
            
            
            'Reviso que haya copiado
            StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
            OpenRecordset StrSql, rs_His_Batch_Proceso
            
            If Not rs_His_Batch_Proceso.EOF Then
                ' Borro de Batch_proceso
                StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
                objConnProgreso.Execute StrSql, , adExecuteNoRecords
            End If
        End If
        ' -----------------------------------------------------------------------------------
        MyCommitTrans
    End If

Final:
    If TransactionRunning Then MyRollbackTrans
    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    If CnTraza.State = adStateOpen Then CnTraza.Close

    If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
    If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    Set rs_Batch_Proceso = Nothing
    Set rs_His_Batch_Proceso = Nothing

    Flog.writeline Espacios(Tabulador * 0) & "Fin :" & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.writeline "Cantidad de Lecturas en BD          : " & Cantidad_de_OpenRecordset
    Flog.writeline "Cantidad de llamadas a politicas    : " & Cantidad_Call_Politicas
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
    
    Exit Sub

CE:
    MyRollbackTrans
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    
    MyBeginTrans
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    GoTo Final
End Sub


Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Nro_Analisis As Long
Dim Tipo_Factor As Integer
Dim Filtrar As Boolean
Dim ArrParametros


'Orden de los parametros
'el primero es número de análisis
'el segundo es el tipo de factor a analizar
'el tercero Si filtra o no (viene siempre en TRUE)

If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        ArrParametros = Split(Parametros, "@")
        
        Nro_Analisis = ArrParametros(0)
        
        Filtrar = True
        
'        pos1 = 1
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Nro_Analisis = Mid(parametros, pos1, pos2)
'
'        'ya no me intersea porque proceso todos los tipos de factores para los factores configurados en la cabecera
'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, parametros, Separador) - 1
'        Tipo_Factor = CInt(Mid(parametros, pos1, pos2 - pos1 + 1))
'
'        pos1 = pos2 + 2
'        pos2 = Len(parametros)
'        Filtrar = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
    End If
End If

'----------------------------
' Fernando Favre - 16-02-05 - Se asigna una porcion de tiempo para procesar cada factor
Dim rsAux As New ADODB.Recordset
StrSql = " SELECT COUNT(anrcab_fact.facnro) AS Cant" & _
    " FROM   anrcab_fact" & _
    " WHERE  anrcab_fact.anrcabnro = " & Nro_Analisis
OpenRecordset StrSql, rsAux
If Not rsAux.EOF Then
    PorcTiempo = CDbl(99) / CDbl(rsAux!cant)
End If
If rsAux.State = adStateOpen Then rsAux.Close
'----------------------------

Call Generar_Analisis(Nro_Analisis, Tipo_Factor, Filtrar)
End Sub

Private Sub Generar_Analisis(ByVal Nro_Analisis As Long, ByVal Tipo_Factor As Long, ByVal Filtrar As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion:  En funcion del tipo de factor ejecuto un procedimiento u otro.
'               Los factores Totalizadores se controlan en cada estos procedimientos.
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim rs_TipoFact As New ADODB.Recordset

StrSql = "SELECT distinct (anrfactor.tipfacnro) FROM anrcab_fact "
StrSql = StrSql & " INNER JOIN anrfactor ON anrcab_fact.facnro = anrfactor.facnro "
StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
StrSql = StrSql & " ORDER BY anrfactor.tipfacnro"
OpenRecordset StrSql, rs_TipoFact

MyBeginTrans
Flog.writeline Espacios(Tabulador * 0) & "Transaccion Iniciada"
SumPorcTiempo = 0

Do While Not rs_TipoFact.EOF
    HuboErrorTipo = False

    Select Case rs_TipoFact!tipfacnro
    Case 1: 'Tipo de Concepto RHPro
        SumPorcTiempo = CDbl(SumPorcTiempo) + CDbl(PorcTiempo)
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Tipo de Concepto RHPro"
    Case 2: 'Concepto de Liquidación  RHPro
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Concepto de Liquidación  RHPro"
        Call ConceptosRHPro(Nro_Analisis, Filtrar)
    Case 3: 'Acumulador Mensual RHPro
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Acumulador Mensual RHPro"
        Call AcumuladoresRHPro(Nro_Analisis, Filtrar)
    Case 4: 'Tipo de Hs de Acu.Diario RHPro
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Tipo de Hs de Acu.Diario RHPro"
        Call AcumuladoDiario(Nro_Analisis, Filtrar)
    Case 5: 'Tipo de Hs de Acu.Parc RHPro
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Tipo de Hs de Acu.Parc RHPro"
        Call AcumuladoParcial(Nro_Analisis, Filtrar)
    Case 6: 'Licencias RH Pro
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Licencias RH Pro"
        Call Licencias(Nro_Analisis, Filtrar)
    Case 7: 'Conceptos Nexus
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Conceptos Nexus"
        Call ConceptosNexus(Nro_Analisis, Filtrar)
    Case 8: 'Suma de otros factores
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "------------------ TOTALIZADORES -------------------------"
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Suma de otros factores"
        Call SumaFactores(Nro_Analisis, Filtrar)
    Case 9: 'Resta de otros factores
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Resta de otros factores"
        Call RestaFactores(Nro_Analisis, Filtrar)
    Case 10: 'Multiplicación de otros factores
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " Multiplicación de otros factores"
        Call ProductoFactores(Nro_Analisis, Filtrar)
    Case 11: 'División de otros factores
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Tipo de Factor " & rs_TipoFact!tipfacnro & " División de otros factores"
        Call DivideFactores(Nro_Analisis, Filtrar)
    End Select
    If Not HuboErrorTipo Then
        'MyCommitTrans
        'Flog.writeline Espacios(Tabulador * 0) & "Transaccion Cometida"
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Ejecutado"
    Else
        'MyRollbackTrans
        'Flog.writeline Espacios(Tabulador * 0) & "Transaccion Abortada"
        Flog.writeline Espacios(Tabulador * 1) & "Analisis Abortado por error"
    End If
    
    ' Actualizo el progreso
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(SumPorcTiempo) & " WHERE bpronro = " & NroProceso
    objConnProgreso.Execute StrSql, , adExecuteNoRecords

    rs_TipoFact.MoveNext
Loop
MyCommitTrans
If Not HuboError Then
    MyCommitTrans
    Flog.writeline Espacios(Tabulador * 0) & "Transaccion Cometida"
Else
    MyRollbackTrans
    Flog.writeline Espacios(Tabulador * 0) & "Transaccion Abortada"
End If
End Sub




Public Sub ObtenerCabecerayFiltro(ByVal Nro_Analisis As Long, ByRef rsAnrCab As ADODB.Recordset, ByRef Filtrar As Boolean, ByRef rs As ADODB.Recordset, ByRef Cantidad As Long, ByRef Ok As Boolean)

Ok = True

StrSql = " SELECT anrcabnro,anrcabfecdesde,anrcabfechasta FROM anrcab " & _
    " WHERE anrcabnro = " & Nro_Analisis
OpenRecordset StrSql, rsAnrCab
    
If rsAnrCab.EOF Then
    Ok = False
End If

Cantidad = 0
If Filtrar Then
    StrSql = " SELECT COUNT( DISTINCT anrcab_filtro.tenro) AS Cant" & _
             " FROM   anrcab_filtro" & _
             " WHERE  anrcab_filtro.anrcabnro = " & rsAnrCab!anrcabnro
    OpenRecordset StrSql, rs

    If rs.EOF Then
        Cantidad = 0
    Else
        Cantidad = rs!cant
    End If

    If (Cantidad <= 0) Then
        Filtrar = False
    End If
End If

End Sub


Private Sub PurgarCubo_OLD(ByVal Nro_Analisis As Long, ByVal TipoFactor As Integer)

StrSql = "DELETE FROM anrcubo " & _
    " WHERE facnro IN " & _
    " (SELECT facnro FROM anrfactor" & _
    " WHERE tipfacnro = " & TipoFactor & ")" & _
    " AND anrcabnro = " & Nro_Analisis & _
    " AND anrcubmanual = 0"
    objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Public Sub PurgarCubo(ByVal Nro_Analisis As Long, ByVal TipoFactor As Integer)
Dim rs_rangos As New ADODB.Recordset

    StrSql = "SELECT * FROM anrrangofec "
    StrSql = StrSql & " WHERE anrcabnro = " & Nro_Analisis
    StrSql = StrSql & " AND anrrangorepro = -1"
    OpenRecordset StrSql, rs_rangos
    
    Do While Not rs_rangos.EOF
        StrSql = "DELETE FROM anrcubo "
        StrSql = StrSql & " WHERE facnro IN "
        StrSql = StrSql & " (SELECT facnro FROM anrfactor"
        StrSql = StrSql & " WHERE tipfacnro = " & TipoFactor & ")"
        StrSql = StrSql & " AND anrcabnro = " & Nro_Analisis
        StrSql = StrSql & " AND anrcubmanual = 0"
        StrSql = StrSql & " AND anrrangnro = " & rs_rangos!anrrangnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        rs_rangos.MoveNext
    Loop
    If rs_rangos.State = adStateOpen Then rs_rangos.Close
    Set rs_rangos = Nothing
End Sub


Public Sub ObtenerLegajos(ByVal TipoGerencial As Integer, ByVal Filtrar As Boolean, ByVal CabNro As Long, ByRef rsFiltro As ADODB.Recordset, ByVal Dia_Inicio_Per_Analizado As Date, ByVal Dia_Fin_Per_Analizado As Date)

Select Case TipoGerencial
Case 1: 'Conceptos Nexus
' obtengo el conjunto de legajos a procesar
    If Filtrar Then
      StrSql = " SELECT     estructura.estrcodext, empleado.empleg, empleado.ternro, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura ea07 " & _
               " ON         ea07.ternro     = empleado.ternro " & _
               " AND ea07.tenro = 7" & _
               " INNER JOIN estructura ON estructura.tenro = ea07.tenro AND estructura.estrnro = ea07.estrnro " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " INNER JOIN anrcab_filtro " & _
               " ON         anrcab_filtro.estrnro     = his_estructura.estrnro " & _
               " AND        anrcab_filtro.anrcabnro   = " & CabNro & _
               " GROUP BY   empleado.empleg, empleado.ternro, estructura.estrcodext" & _
               " ORDER BY   empleado.empleg"

' Se exige que los empleados cumplan con todas las condiciones especificadas. O.D.A. 27/06/2003

    Else
      StrSql = " SELECT     DISTINCT empleado.empleg, empleado.ternro, 0 as cant_te, estructura.estrcodext" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura ea07 " & _
               " ON         ea07.ternro     = empleado.ternro " & _
               " AND ea07.tenro = 7" & _
               " INNER JOIN estructura ON estructura.tenro = ea07.tenro AND estructura.estrnro = ea07.estrnro " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) "
    End If
    
Case 2: 'Acumulados Diarios
    If Filtrar Then
      StrSql = " SELECT     empleado.empleg, empleado.ternro, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " INNER JOIN anrcab_filtro " & _
               " ON         anrcab_filtro.estrnro     = his_estructura.estrnro " & _
               " AND        anrcab_filtro.anrcabnro   = " & CabNro & _
               " GROUP BY   empleado.empleg, empleado.ternro" & _
               " ORDER BY   empleado.empleg"

' Se exige que los empleados cumplan con todas las condiciones especificadas. O.D.A. 27/06/2003

    Else
      StrSql = " SELECT     DISTINCT empleado.empleg, empleado.ternro, 0 as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) "
    End If
    
    
Case 3: 'Acumulados Parciales
    If Filtrar Then
      StrSql = " SELECT     empleado.empleg, empleado.ternro, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " INNER JOIN anrcab_filtro " & _
               " ON         anrcab_filtro.estrnro     = his_estructura.estrnro " & _
               " AND        anrcab_filtro.anrcabnro   = " & CabNro & _
               " GROUP BY   empleado.empleg, empleado.ternro" & _
               " ORDER BY   empleado.empleg"

' Se exige que los empleados cumplan con todas las condiciones especificadas. O.D.A. 27/06/2003

    Else
      StrSql = " SELECT     DISTINCT empleado.empleg, empleado.ternro, 0 as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) "
    End If

Case 4: 'Licencias
    If Filtrar Then
      StrSql = " SELECT     empleado.empleg, empleado.ternro, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " INNER JOIN anrcab_filtro " & _
               " ON         anrcab_filtro.estrnro     = his_estructura.estrnro " & _
               " AND        anrcab_filtro.anrcabnro   = " & CabNro & _
               " GROUP BY   empleado.empleg, empleado.ternro" & _
               " ORDER BY   empleado.empleg"

' Se exige que los empleados cumplan con todas las condiciones especificadas. O.D.A. 27/06/2003

    Else
      StrSql = " SELECT     DISTINCT empleado.empleg, empleado.ternro, 0 as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) "
    End If

Case 5: 'Totalizadores
' obtengo el conjunto de legajos a procesar
    If Filtrar Then
      StrSql = " SELECT     empleado.empleg, empleado.ternro, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " INNER JOIN anrcab_filtro " & _
               " ON         anrcab_filtro.estrnro     = his_estructura.estrnro " & _
               " AND        anrcab_filtro.anrcabnro   = " & CabNro & _
               " GROUP BY   empleado.empleg, empleado.ternro" & _
               " ORDER BY   empleado.empleg"

' Se exige que los empleados cumplan con todas las condiciones especificadas. O.D.A. 27/06/2003

    Else
      StrSql = " SELECT     DISTINCT empleado.empleg, empleado.ternro, 0 as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) "
    End If

Case Else
End Select
    
OpenRecordset StrSql, rsFiltro

End Sub


Private Sub ObtenerLegajos_OLD(ByVal TipoGerencial As Integer, ByVal Filtrar As Boolean, ByVal CabNro As Long, ByRef rsFiltro As ADODB.Recordset, ByVal Dia_Inicio_Per_Analizado As Date, ByVal Dia_Fin_Per_Analizado As Date)

Select Case TipoGerencial
Case 1: 'Conceptos Nexus
' obtengo el conjunto de legajos a procesar
    If Filtrar Then
      StrSql = " SELECT     estructura.estrcodext, empleado.empleg, empleado.ternro, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura ea07 " & _
               " ON         ea07.ternro     = empleado.ternro " & _
               " AND ea07.tenro = 7" & _
               " INNER JOIN estructura ON estructura.tenro = ea07.tenro AND estructura.estrnro = ea07.estrnro " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " INNER JOIN anrcab_filtro " & _
               " ON         anrcab_filtro.estrnro     = his_estructura.estrnro " & _
               " AND        anrcab_filtro.anrcabnro   = " & CabNro & _
               " GROUP BY   empleado.empleg, empleado.ternro, estructura.estrcodext" & _
               " ORDER BY   empleado.empleg"

' Se exige que los empleados cumplan con todas las condiciones especificadas. O.D.A. 27/06/2003

    Else
      StrSql = " SELECT     DISTINCT empleado.empleg, empleado.ternro, 0 as cant_te, estructura.estrcodext" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura ea07 " & _
               " ON         ea07.ternro     = empleado.ternro " & _
               " AND ea07.tenro = 7" & _
               " INNER JOIN estructura ON estructura.tenro = ea07.tenro AND estructura.estrnro = ea07.estrnro " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) "
    End If
    
Case 2: 'Acumulados Diarios
    If Filtrar Then
        StrSql = " SELECT   his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " GROUP BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro, his_estructura.htetdesde, his_estructura.htethasta" & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    Else
      StrSql = " SELECT     DISTINCT his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, 0 as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    End If
    
    
Case 3: 'Acumulados Parciales
    If Filtrar Then
        StrSql = " SELECT   his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " GROUP BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro, his_estructura.htetdesde, his_estructura.htethasta" & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    Else
      StrSql = " SELECT     DISTINCT his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, 0 as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    End If

Case 4: 'Licencias
    If Filtrar Then
        StrSql = " SELECT   his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " GROUP BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro, his_estructura.htetdesde, his_estructura.htethasta" & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    Else
      StrSql = " SELECT     DISTINCT his_estructura.tenro, his_estructura.estrnro, his_estructura.ternro, his_estructura.htethasta, his_estructura.htetdesde, 0 as cant_te" & _
               " FROM       his_estructura " & _
               " WHERE      his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " ORDER BY   his_estructura.ternro, his_estructura.tenro, his_estructura.estrnro"
    End If

Case 5: 'Totalizadores
' obtengo el conjunto de legajos a procesar
    If Filtrar Then
      StrSql = " SELECT     empleado.empleg, empleado.ternro, COUNT( DISTINCT his_estructura.tenro) as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) " & _
               " INNER JOIN anrcab_filtro " & _
               " ON         anrcab_filtro.estrnro     = his_estructura.estrnro " & _
               " AND        anrcab_filtro.anrcabnro   = " & CabNro & _
               " GROUP BY   empleado.empleg, empleado.ternro" & _
               " ORDER BY   empleado.empleg"

' Se exige que los empleados cumplan con todas las condiciones especificadas. O.D.A. 27/06/2003

    Else
      StrSql = " SELECT     DISTINCT empleado.empleg, empleado.ternro, 0 as cant_te" & _
               " FROM       empleado " & _
               " INNER JOIN his_estructura " & _
               " ON         his_estructura.ternro     = empleado.ternro " & _
               " AND        his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
               " AND       (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & " OR " & _
               "            his_estructura.htethasta IS NULL) "
    End If

Case Else
End Select
    
OpenRecordset StrSql, rsFiltro

End Sub

Public Sub ObtenerEstructuras(ByVal Filtrar As Boolean, ByVal tercero As Long, ByVal FechaInicio As Date, ByVal FechaFin As Date, ByRef rs As ADODB.Recordset)
If Filtrar Then
    StrSql = "SELECT * FROM his_estructura" & _
        " WHERE his_estructura.ternro = " & tercero & _
        " AND his_estructura.htetdesde <= " & ConvFecha(FechaFin) & _
        " AND (his_estructura.htethasta >= " & ConvFecha(FechaInicio) & _
        " OR his_estructura.htethasta IS NULL)" & _
        " ORDER BY ternro,tenro,estrnro"
Else ' no se usa el filtro ==> todas las estructuras
    StrSql = "SELECT * FROM his_estructura" & _
        " WHERE his_estructura.ternro = " & tercero & _
        " AND his_estructura.htetdesde <= " & ConvFecha(FechaFin) & _
        " AND (his_estructura.htethasta >= " & ConvFecha(FechaInicio) & _
        " OR his_estructura.htethasta IS NULL)" & _
        " ORDER BY ternro,tenro,estrnro"
End If
OpenRecordset StrSql, rs

End Sub


Public Function Ultimo(ByRef rs As ADODB.Recordset) As Boolean
Dim resultado As Boolean

    'Trato de obtener el próximo
    rs.MoveNext
    'Si es vacío entonces tengo al último del grupo
    If rs.EOF Then
        resultado = True
    Else
        resultado = False
    End If
    
    rs.MovePrevious
    Ultimo = resultado
End Function

Function Last_OF_estrnro() As Boolean
Dim resultado As Boolean
Dim Actual As Long

    Actual = rsEstructura!estrnro
    'Trato de obtener el próximo
    rsEstructura.MoveNext
    'Si es vacío entonces tengo al último del grupo
    If rsEstructura.EOF Then
        resultado = True
    Else
        'Si el proximo es distinto del actual entonces el actual es el ultimo
        If rsEstructura!estrnro <> Actual Then
            resultado = True
        Else
            resultado = False
        End If
    End If
    
    rsEstructura.MovePrevious
    Last_OF_estrnro = resultado
    
End Function

Function Last_OF_tenro() As Boolean
Dim resultado As Boolean
Dim Actual As Long

    Actual = rsEstructura!tenro
    'Trato de obtener el próximo
    rsEstructura.MoveNext
    'Si es vacío entonces tengo al último del grupo
    If rsEstructura.EOF Then
        resultado = True
    Else
        'Si el proximo es distinto del actual entonces el actual es el ultimo
        If rsEstructura!tenro <> Actual Then
            resultado = True
        Else
            resultado = False
        End If
    End If
    
    rsEstructura.MovePrevious
    Last_OF_tenro = resultado
    
End Function



Public Function Espacios(ByVal Cantidad As Integer) As String
    Espacios = Space(Cantidad)
End Function

Public Function CalcularPresupuestado(ByVal valor As Double, ByVal facpresup As Boolean, ByVal facopfijo As Boolean, ByVal facopsuma As Boolean, ByVal facpresupmonto As Double)
' --------------------------------------------------------------------------------------------
' Descripcion: Calcula el valor Presupuestado.
' Autor      : Fernando Favre
' Fecha      : 14-02-2005
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim resultado As Double

    If facpresup Then
        If facopfijo Then
            If facopsuma Then
                resultado = valor + CDbl(facpresupmonto)
            Else
                resultado = valor - CDbl(facpresupmonto)
            End If
        Else
            If facopsuma Then
                resultado = valor + ((valor * CDbl(facpresupmonto)) / 100)
            Else
                resultado = valor - ((valor * CDbl(facpresupmonto)) / 100)
            End If
        End If
    End If
    
    CalcularPresupuestado = resultado
End Function

Public Function EsNulo(ByVal Objeto) As Boolean
    If IsNull(Objeto) Then
        EsNulo = True
    Else
        If UCase(Objeto) = "NULL" Or UCase(Objeto) = "" Then
            EsNulo = True
        Else
            EsNulo = False
        End If
    End If
End Function



