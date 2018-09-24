Attribute VB_Name = "MdlRepObraSocial"
Option Explicit

'Const Version = 1
'Const FechaVersion = "25/10/2006"
'Modificaciones: Mariano Capriz
'                       Version inicial

'Const Version = "1.01"
'Const FechaVersion = "31/07/2009"
'Modificaciones: Martin Ferraro - Encriptacion de string connection

Const Version = "1.03"
Const FechaVersion = "23/11/2012"
'Modificaciones: Sebastian Stremel - CAS-17473 - SEDAMIL - Error reporte OSocial
'queda la version 03 como ultima ya que se habian perdido versiones y en esta entrega se nivela.
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Global IdUser As String
Global Fecha As Date
Global hora As String

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo As Date



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reportes.
' Autor      : FGZ
' Fecha      : 19/02/2004
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

'    strCmdLine = Command()
'    ArrParametros = Split(strCmdLine, " ", -1)
'    If UBound(ArrParametros) > 0 Then
'        If IsNumeric(ArrParametros(0)) Then
'            NroProcesoBatch = ArrParametros(0)
'            Etiqueta = ArrParametros(1)
'        Else
'            Exit Sub
'        End If
'    Else
'        If IsNumeric(strCmdLine) Then
'            NroProcesoBatch = strCmdLine
'        Else
'            Exit Sub
'        End If
'    End If
    
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
    
    Nombre_Arch = PathFLog & "Reporte_ObraSocial" & "-" & NroProcesoBatch & ".log"
    
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

    'Abro la conexion
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
    Flog.writeline
    Flog.writeline "Cambio el estado del proceso"
    Flog.writeline
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    'objConn.Execute StrSql, , adExecuteNoRecords
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Obtengo los datos del proceso"
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 33 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Sueoso06(NroProcesoBatch, bprcparam)
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

Exit Sub
CE:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error General " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
End Sub


Public Sub Sueoso06(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del Reporte de Obra Social
' Autor      : FGZ
' Fecha      : 19/02/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1, pos2 As Integer

Dim fechadesde As Date
Dim fechahasta As Date

Dim Nroliq As Long
Dim Todos_Pro As Boolean
Dim Proc_Aprob As Integer
Dim Empresa As Long
Dim Todas_OSocial As Boolean
Dim Nro_OSocial As Long
Dim Valorizado As Boolean
Dim Agrupado As Integer

Dim Arreglo(20) As Single
Dim I As Integer
Dim R As Integer

Dim Tope_Ampo_Max As Single
Dim Tope_Ampo_Min As Single
Dim par_Osocial As Long
Dim Imp_OS As Single
Dim msr As Single
Dim msr_liq As Single
Dim suma_osoc As Boolean
Dim vacio As Boolean

Dim Ultimo_Empleado As Long
Dim Aux_plprecio As Single
Dim Aux_plnro As Long
Dim Aux_Cuil As String

Dim rs_Reporte As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_Rep03 As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
'Dim rs_Empleados As New ADODB.Recordset
Dim rs_Planos As New ADODB.Recordset
Dim rs_Ter_Doc As New ADODB.Recordset

' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        Nroliq = CLng(Mid(Parametros, pos1, pos2))
    
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Todos_Pro = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        If Not Todos_Pro Then
            pos1 = pos2 + 2
            pos2 = InStr(pos1, Parametros, ".") - 1
            NroProc = Mid(Parametros, pos1, pos2 - pos1 + 1)
            ListaNroProc = Replace(NroProc, "-", ",")
        Else
            NroProc = "0"
            ListaNroProc = "0"
        End If
            
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Proc_Aprob = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Empresa = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Todas_OSocial = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        If Not Todas_OSocial Then
            pos1 = pos2 + 2
            pos2 = InStr(pos1, Parametros, ".") - 1
            Nro_OSocial = Mid(Parametros, pos1, pos2 - pos1 + 1)
        End If
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Valorizado = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        Agrupado = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
    End If
End If

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If
Fecha_Inicio_periodo = rs_Periodo!pliqdesde
Fecha_Fin_Periodo = rs_Periodo!pliqhasta



'Inicializacion
For I = 1 To 20
    Arreglo(I) = 0
Next I
suma_osoc = False
vacio = False

StrSql = "Select * FROM reporte where reporte.repnro = 3"
OpenRecordset StrSql, rs_Reporte
If rs_Reporte.EOF Then
    Flog.writeln "El Reporte Numero 3 no ha sido Configurado"
    Exit Sub
End If
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close


' Comienzo la transaccion
MyBeginTrans

'Depuracion del Temporario
If Not Todos_Pro Then
    StrSql = "DELETE FROM rep03 " & _
             " WHERE pliqnro = " & Nroliq & _
             " AND pronro = '" & NroProc & "'" & _
             " AND empresa = " & Empresa
Else
    StrSql = "DELETE FROM rep03 " & _
             " WHERE pliqnro = " & Nroliq & _
             " AND pronro = '0' " & _
             " AND proaprob = " & CInt(Proc_Aprob) & _
             " AND empresa = " & Empresa
End If
objConn.Execute StrSql, , adExecuteNoRecords

'Busco los procesos a evaluar
StrSql = "SELECT distinct cabliq.*, proceso.*, periodo.*, empleado.*, oSocial.estrnro osnro FROM  empleado "
'StrSql = "SELECT cabliq.*, proceso.*, periodo.*, empleado.*, oSocial.estrnro osnro FROM  empleado "
StrSql = StrSql & " INNER JOIN his_estructura oSocial ON oSocial.ternro = empleado.ternro "
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.empleado = empleado.ternro "
StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro "
StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = proceso.pliqnro "
StrSql = StrSql & " INNER JOIN his_estructura  empresa ON empresa.ternro = empleado.ternro and empresa.tenro = 10 "
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
'FGZ - 14/07/2005
StrSql = StrSql & " INNER JOIN empresa emp ON emp.estrnro = empresa.estrnro AND emp.empnro =" & Empresa
'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro "
StrSql = StrSql & " WHERE oSocial.tenro = 17 AND "
StrSql = StrSql & " (oSocial.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= oSocial.htethasta) or (oSocial.htethasta is null))"
StrSql = StrSql & " AND (empresa.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= empresa.htethasta) or (empresa.htethasta is null))"
StrSql = StrSql & " AND periodo.pliqnro =" & Nroliq
If Not Todas_OSocial Then
    StrSql = StrSql & " AND oSocial.estrnro = " & Nro_OSocial
End If
If Not Todos_Pro Then
    'StrSql = StrSql & " AND proceso.pronro =" & NroProc
    StrSql = StrSql & " AND proceso.pronro IN (" & ListaNroProc & ")"
Else
    StrSql = StrSql & " AND proceso.empnro = " & Empresa & " AND proaprob = " & CInt(Proc_Aprob)
End If
StrSql = StrSql & " ORDER BY empleado.ternro"
OpenRecordset StrSql, rs_Procesos

If rs_Procesos.EOF Then
    Flog.writeline "No se encontró ningun empleado para procesar "
    Flog.writeline "SQL = " & StrSql
End If


'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 3 "
OpenRecordset StrSql, rs_Confrep

If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If

'seteo de las variables de progreso
Progreso = 0
CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
   CConceptosAProc = 1
End If
CEmpleadosAProc = rs_Confrep.RecordCount
If CEmpleadosAProc = 0 Then
   CEmpleadosAProc = 1
End If
IncPorc = ((100 / CEmpleadosAProc) * (100 / CConceptosAProc)) / 100

Ultimo_Empleado = -1
Do While Not rs_Procesos.EOF
    Flog.writeline Espacios(Tabulador * 1) & "Legajo " & rs_Procesos!empleg & " Proceso : " & rs_Procesos!pronro

    If Ultimo_Empleado <> rs_Procesos!Ternro Then
        For I = 1 To 20
            Arreglo(I) = 0
        Next I
    End If
    Ultimo_Empleado = rs_Procesos!Ternro
    
   
    Flog.writeline Espacios(Tabulador * 1) & "CONFREP. Conceptos y Acumuladores"
        rs_Confrep.MoveFirst
        Do While Not rs_Confrep.EOF
            Flog.writeline Espacios(Tabulador * 2) & "Columna: " & rs_Confrep!confnrocol
            Select Case UCase(rs_Confrep!conftipo)
            Case "AC":
                Flog.writeline Espacios(Tabulador * 2) & "Acumulador " & rs_Confrep!confval
                StrSql = "SELECT * FROM acu_liq " & _
                         " INNER JOIN cabliq ON acu_liq.cliqnro = cabliq.cliqnro " & _
                         " WHERE acu_liq.acunro = " & rs_Confrep!confval & _
                         " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
                OpenRecordset StrSql, rs_Acu_Liq
                If rs_Acu_Liq.EOF Then
                    Flog.writeline Espacios(Tabulador * 3) & "No se encontró"
                End If
                Do While Not rs_Acu_Liq.EOF
                    Flog.writeline Espacios(Tabulador * 3) & "Suma " & rs_Acu_Liq!almonto
                    Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Acu_Liq!almonto
                    vacio = False
                    
                    rs_Acu_Liq.MoveNext
                Loop
                
            Case "CO", "EDR":
                Flog.writeline Espacios(Tabulador * 2) & "Concepto " & rs_Confrep!confval & "(" & rs_Confrep!confval2 & ")"
                StrSql = "SELECT * FROM detliq " & _
                         " INNER JOIN cabliq ON detliq.cliqnro = cabliq.cliqnro " & _
                         " INNER JOIN concepto ON detliq.concnro = concepto.concnro " & _
                         " WHERE (concepto.conccod = " & rs_Confrep!confval & _
                         " OR concepto.conccod = '" & rs_Confrep!confval2 & "')" & _
                         " AND cabliq.cliqnro =" & rs_Procesos!cliqnro
                OpenRecordset StrSql, rs_Detliq
                If rs_Detliq.EOF Then
                    Flog.writeline Espacios(Tabulador * 3) & "No se encontró"
                End If
                Do While Not rs_Detliq.EOF
                    Flog.writeline Espacios(Tabulador * 3) & "Suma " & rs_Detliq!dlimonto
                    Arreglo(rs_Confrep!confnrocol) = Arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlimonto
                    vacio = False
                    
                    rs_Detliq.MoveNext
                Loop
                
            End Select
            
                
            If EsElUltimoEmpleado(rs_Procesos, Ultimo_Empleado) And Not vacio Then
                'Si no existe el rep03
                StrSql = "SELECT * FROM rep03 "
                StrSql = StrSql & " WHERE ternro = " & rs_Procesos!Ternro
                StrSql = StrSql & " AND bpronro = " & bpronro
                StrSql = StrSql & " AND pliqnro = " & Nroliq
                StrSql = StrSql & " AND empresa = " & Empresa
                If Not Todos_Pro Then
                    StrSql = StrSql & " AND pronro = '" & NroProc & "'"
                Else
                    StrSql = StrSql & " AND pronro = '0'"
                    StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
                End If
                StrSql = StrSql & " AND ordenrep =" & Agrupado
                OpenRecordset StrSql, rs_Rep03
            
                If rs_Rep03.EOF Then
                    'Busco el doc del tercero
                    StrSql = "SELECT * FROM ter_doc WHERE ternro =" & rs_Procesos!Ternro & _
                             " AND tidnro = 10"
                    OpenRecordset StrSql, rs_Ter_Doc
                    If Not rs_Ter_Doc.EOF Then
                        Aux_Cuil = rs_Ter_Doc!nrodoc
                    End If
                    
                    
                    'Busco el plan de la OS
                    StrSql = "SELECT * FROM His_estructura "
                    StrSql = StrSql & " INNER JOIN replica_estr ON replica_estr.estrnro = his_estructura.estrnro "
                    StrSql = StrSql & " INNER JOIN planos ON planos.plnro = replica_estr.origen "
                    StrSql = StrSql & " WHERE his_estructura.tenro = 23 AND "
                    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
                    StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                    StrSql = StrSql & " AND his_estructura.ternro =" & rs_Procesos!Ternro
                    OpenRecordset StrSql, rs_Planos
                    
                    'StrSql = "SELECT * FROM planos WHERE osocial =" & rs_Procesos!ternro
                    'OpenRecordset StrSql, rs_Planos
                    If Not rs_Planos.EOF Then
                        Aux_plprecio = IIf(IsNull(rs_Planos!plprecio), 0, rs_Planos!plprecio)
                        Aux_plnro = rs_Planos!plnro
                    Else
                        Flog.writeline Espacios(Tabulador * 2) & "No se encontró plano de la Obra Social. plprecio = 0"
                        Aux_plprecio = 0
                        Aux_plnro = 0
                    End If
                
                    Flog.writeline Espacios(Tabulador * 2) & "Inserta"
                    'Inserto
                    StrSql = "INSERT INTO rep03 (bpronro,pliqnro,pronro,proaprob,empresa,iduser,fecha,hora,"
                    StrSql = StrSql & "cuil,plnro,plan_valorizado,ternro,ordenrep,osocial) VALUES ("
                    StrSql = StrSql & bpronro & ","
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
                    StrSql = StrSql & "'" & hora & "',"
                    If Not IsNull(Aux_Cuil) Then
                        StrSql = StrSql & "'" & Aux_Cuil & "',"
                    End If
                    If Valorizado Then
                        StrSql = StrSql & Aux_plnro & ","
                        StrSql = StrSql & Aux_plprecio & ","
                    Else
                        StrSql = StrSql & "0,0" & ","
                    End If
                    StrSql = StrSql & rs_Procesos!Ternro & ","
                    StrSql = StrSql & Agrupado & ","
                    StrSql = StrSql & rs_Procesos!osnro & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    vacio = True
                End If
                
                Flog.writeline Espacios(Tabulador * 2) & "Actualiza valores"
                'Actualizo
                StrSql = "UPDATE rep03 SET bruto =" & Arreglo(1)
                StrSql = StrSql & ", msr = " & Arreglo(2)
                StrSql = StrSql & ", empleado = " & Arreglo(3)
                StrSql = StrSql & ", empleador = " & Arreglo(4)
                StrSql = StrSql & ", adicional = " & Arreglo(5)
                StrSql = StrSql & ", anssal_empleador = " & Arreglo(6)
                StrSql = StrSql & ", anssal_empleado = " & Arreglo(7)
                StrSql = StrSql & ", dif_empleado = " & Arreglo(8)
                StrSql = StrSql & ", empdo_subtotal = " & Arreglo(9)
                StrSql = StrSql & ", empdor_subtotal = " & Arreglo(10)
                StrSql = StrSql & ", os_totdep = " & Arreglo(11)
                StrSql = StrSql & ", anssal_totdep = " & Arreglo(12)
                StrSql = StrSql & ", totdep = " & Arreglo(13)
                StrSql = StrSql & ", adi_mas_plaval = " & Arreglo(14)
                StrSql = StrSql & ", empdo_menos_anss = " & Arreglo(15)
                StrSql = StrSql & ", empdor_menos_anss = " & Arreglo(16)
                StrSql = StrSql & ", neto = " & Arreglo(17)
                StrSql = StrSql & ", dif_os = " & Arreglo(18)
                StrSql = StrSql & ", dif_empdor = " & Arreglo(19)
                StrSql = StrSql & " WHERE ternro = " & rs_Procesos!Ternro
                StrSql = StrSql & " AND bpronro = " & bpronro
                StrSql = StrSql & " AND pliqnro = " & Nroliq
                StrSql = StrSql & " AND empresa = " & Empresa
                If Not Todos_Pro Then
                    StrSql = StrSql & " AND pronro = '" & NroProc & "'"
                Else
                    StrSql = StrSql & " AND pronro = '0'"
                    StrSql = StrSql & " AND proaprob= " & CInt(Proc_Aprob)
                End If
                StrSql = StrSql & " AND ordenrep =" & Agrupado
                objConn.Execute StrSql, , adExecuteNoRecords
                
            End If
            
            'Actualizo el progreso del Proceso
            Progreso = Progreso + IncPorc
            TiempoAcumulado = GetTickCount
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                     "' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
            'Siguiente confrep
            rs_Confrep.MoveNext
        Loop
        Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
        Flog.writeline
        
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop

'Fin de la transaccion
MyCommitTrans


If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close
If rs_Rep03.State = adStateOpen Then rs_Rep03.Close
If rs_Reporte.State = adStateOpen Then rs_Reporte.Close
If rs_Planos.State = adStateOpen Then rs_Planos.Close
If rs_Ter_Doc.State = adStateOpen Then rs_Ter_Doc.Close

Set rs_Ter_Doc = Nothing
Set rs_Procesos = Nothing
Set rs_Confrep = Nothing
Set rs_Detliq = Nothing
Set rs_Acu_Liq = Nothing
Set rs_Rep03 = Nothing
Set rs_Periodo = Nothing
Set rs_Reporte = Nothing
Set rs_Planos = Nothing


Exit Sub
CE:
    HuboError = True
    MyRollbackTrans

End Sub


Public Function EsElUltimoEmpleado(ByVal rs As ADODB.Recordset, ByVal Anterior As Long) As Boolean
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

