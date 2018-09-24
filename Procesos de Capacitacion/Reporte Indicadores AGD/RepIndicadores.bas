Attribute VB_Name = "RepIndicadores"
 Option Explicit

Global Const Version = "1.01"
Global Const FechaModificacion = "22/08/2007"
Global Const UltimaModificacion = " " 'Version Inicial

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Global fs, f
Global Flog
Global NroProceso As Long

Global Path As String
Global HuboErrores As Boolean

'NUEVAS
Global IdUser As String
Global Fecha As Date
Global Hora As String

Private Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de generacion de datos
' Autor      : FAF
' Fecha      : 22/08/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine As String
Dim Nombre_Arch As String

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim PID As String
Dim Parametros As String
Dim ArrParametros

'Dim Empresa As Long
Dim Anio As Integer
Dim fecdesdeAnual As Date
Dim fechastaAnual As Date
Dim fecdesdeCuatr As Date
Dim fechastaCuatr As Date
Dim tenro1 As Integer
Dim estrnro1 As Long
Dim tenro2 As Integer
Dim estrnro2 As Long
Dim tenro3 As Integer
Dim estrnro3 As Long
Dim indicador As String
Dim tedabr1 As String
Dim estrdabr1 As String
Dim tedabr2 As String
Dim estrdabr2 As String
Dim tedabr3 As String
Dim estrdabr3 As String
       
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    TiempoInicialProceso = GetTickCount
    HuboErrores = False
    
    Nombre_Arch = PathFLog & "ReporteIndicadores" & "-" & NroProceso & ".log"
    
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
   
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ", bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Espacios(Tabulador * 0) & "Obtengo los datos del proceso"
    
    TiempoAcumulado = GetTickCount
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND btprcnro = 195"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
       IdUser = rs!IdUser
       Fecha = rs!bprcfecha
       Hora = rs!bprchora
       
       'Obtengo los parametros del proceso
       Parametros = rs!bprcparam
       ArrParametros = Split(Parametros, "@")
       
'       Empresa = CLng(ArrParametros(0))
       Anio = CInt(ArrParametros(0))
       fecdesdeAnual = CDate(ArrParametros(1))
       fechastaAnual = CDate(ArrParametros(2))
       fecdesdeCuatr = CDate(ArrParametros(3))
       fechastaCuatr = CDate(ArrParametros(4))
       tenro1 = CInt(ArrParametros(5))
       estrnro1 = CLng(ArrParametros(6))
       tenro2 = CInt(ArrParametros(7))
       estrnro2 = CLng(ArrParametros(8))
       tenro3 = CInt(ArrParametros(9))
       estrnro3 = CLng(ArrParametros(10))
       indicador = CStr(ArrParametros(11))
       
       tedabr1 = ""
       estrdabr1 = ""
       tedabr2 = ""
       estrdabr2 = ""
       tedabr3 = ""
       estrdabr3 = ""
       If tenro1 <> 0 Then
            StrSql = "SELECT tedabr FROM tipoestructura WHERE tenro = " & tenro1
            OpenRecordset StrSql, rs1
            If Not rs1.EOF Then
                tedabr1 = rs1!tedabr
            End If
            rs1.Close
            
            If estrnro1 <> 0 Then
                StrSql = "SELECT estrdabr FROM estructura WHERE estrnro = " & estrnro1
                OpenRecordset StrSql, rs1
                If Not rs1.EOF Then
                    estrdabr1 = rs1!estrdabr
                End If
                rs1.Close
            End If
       End If
       If tenro2 <> 0 Then
            StrSql = "SELECT tedabr FROM tipoestructura WHERE tenro = " & tenro2
            OpenRecordset StrSql, rs1
            If Not rs1.EOF Then
                tedabr2 = rs1!tedabr
            End If
            rs1.Close
            
            If estrnro2 <> 0 Then
                StrSql = "SELECT estrdabr FROM estructura WHERE estrnro = " & estrnro2
                OpenRecordset StrSql, rs1
                If Not rs1.EOF Then
                    estrdabr2 = rs1!estrdabr
                End If
                rs1.Close
            End If
       End If
       If tenro3 <> 0 Then
            StrSql = "SELECT tedabr FROM tipoestructura WHERE tenro = " & tenro3
            OpenRecordset StrSql, rs1
            If Not rs1.EOF Then
                tedabr3 = rs1!tedabr
            End If
            rs1.Close
            
            If estrnro3 <> 0 Then
                StrSql = "SELECT estrdabr FROM estructura WHERE estrnro = " & estrnro3
                OpenRecordset StrSql, rs1
                If Not rs1.EOF Then
                    estrdabr3 = rs1!estrdabr
                End If
                rs1.Close
            End If
       End If
       
       StrSql = "INSERT INTO rep_ind (bpronro,anio,fdesdeanual,fhastaanual,fdesdecuat,fhastacuat,tenro1,tedabr1,"
       StrSql = StrSql & "estrnro1,estrdabr1,tenro2,tedabr2,estrnro2,estrdabr2,tenro3,tedabr3,estrnro3,estrdabr3,"
       StrSql = StrSql & "indicador,fecha,hora,iduser)"
       StrSql = StrSql & " VALUES (" & NroProceso
       StrSql = StrSql & ", " & Anio
       StrSql = StrSql & ", " & ConvFecha(fecdesdeAnual)
       StrSql = StrSql & ", " & ConvFecha(fechastaAnual)
       StrSql = StrSql & ", " & ConvFecha(fecdesdeCuatr)
       StrSql = StrSql & ", " & ConvFecha(fechastaCuatr)
       StrSql = StrSql & ", " & tenro1
       StrSql = StrSql & ", '" & tedabr1 & "'"
       StrSql = StrSql & ", " & estrnro1
       StrSql = StrSql & ", '" & estrdabr1 & "'"
       StrSql = StrSql & ", " & tenro2
       StrSql = StrSql & ", '" & tedabr2 & "'"
       StrSql = StrSql & ", " & estrnro2
       StrSql = StrSql & ", '" & estrdabr2 & "'"
       StrSql = StrSql & ", " & tenro3
       StrSql = StrSql & ", '" & tedabr3 & "'"
       StrSql = StrSql & ", " & estrnro3
       StrSql = StrSql & ", '" & estrdabr3 & "'"
       StrSql = StrSql & ", '" & indicador & "'"
       StrSql = StrSql & ", " & ConvFecha(Fecha)
       StrSql = StrSql & ", '" & Hora & "'"
       StrSql = StrSql & ", '" & IdUser & "')"
       objConn.Execute StrSql, , adExecuteNoRecords
       
       Select Case indicador
        Case "A":
            'A - Hs/ h promedio de capacitación por empleado por sector
           Call Generar_Datos_A(Anio, fecdesdeAnual, fechastaAnual, fecdesdeCuatr, fechastaCuatr, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3)
        Case "B":
            'B - Porcentaje del tiempo dedicado a capacitación por sector
           Call Generar_Datos_B(Anio, fecdesdeAnual, fechastaAnual, fecdesdeCuatr, fechastaCuatr, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3)
        Case "C":
            'C - Cumplimiento Plan de Capacitación
           Call Generar_Datos_C(Anio, fecdesdeAnual, fechastaAnual, fecdesdeCuatr, fechastaCuatr, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3)
        Case "D":
            'D - Hs/ h promedio de capacitación por empleado por sector
           Call Generar_Datos_D(Anio, fecdesdeAnual, fechastaAnual, fecdesdeCuatr, fechastaCuatr, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3)
        Case "E":
            'E - Evaluación del Alumno
           Call Generar_Datos_E(Anio, fecdesdeAnual, fechastaAnual, fecdesdeCuatr, fechastaCuatr, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3)
        Case "F":
            'F - Asistencia
           Call Generar_Datos_F(Anio, fecdesdeAnual, fechastaAnual, fecdesdeCuatr, fechastaCuatr, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3)
        Case "G":
            'G - Inversi&oacute;n en Capacitación por Empleado
           Call Generar_Datos_G(Anio, fecdesdeAnual, fechastaAnual, fecdesdeCuatr, fechastaCuatr, tenro1, estrnro1, tenro2, estrnro2, tenro3, estrnro3)
       End Select
    Else
        Flog.writeline Espacios(Tabulador * 0) & "No se encontraron los datos del proceso nro " & NroProceso
    End If
    
    
    'Actualizo el estado del proceso
    If Not HuboErrores Then
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Correctamente"
    Else
       StrSql = "UPDATE batch_proceso SET  bprcprogreso =100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
       Flog.writeline Espacios(Tabulador * 0) & "Proceso Finalizado Incompleto"
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords

    TiempoFinalProceso = GetTickCount
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.Close
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set rs1 = Nothing
    objconnProgreso.Close
    objConn.Close
Exit Sub
    
ME_Main:
    HuboErrores = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
End Sub


Private Sub Generar_Datos_A(ByVal AnioIni As Integer, ByVal fecdesdeA As Date, ByVal fechastaA As Date, ByVal fecdesdeC As Date, ByVal fechastaC As Date, ByVal tipo1 As Integer, ByVal estruc1 As Long, ByVal tipo2 As Integer, ByVal estruc2 As Long, ByVal tipo3 As Integer, ByVal estruc3 As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera los datos de Indicador A
' Autor      : FAF
' Fecha      : 21/08/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim cantRegistros As Long

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
    
Dim tenro_conf As String
Dim lista_estrnro_conf As String
Dim fecdesdeA_anioAnt As Date
Dim fechastaA_anioAnt As Date
Dim fecCalculoIni As Date
Dim fecCalculoFin As Date
Dim AnioIni_anioAnt As Integer
Dim NroReporte As Integer
Dim estrnro_aux
Dim estrdabr_aux
Dim ternro_aux
Dim empleg_aux
Dim terape_aux
Dim ternom_aux
Dim terape2_aux
Dim ternom2_aux

    On Error GoTo ME_Local
      
    'Configuracion del Reporte
    NroReporte = 209
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    lista_estrnro_conf = "0"
    tenro_conf = "0"
    Do Until rs.EOF
        If UCase(rs!conftipo) = "TE" And UCase(Mid(rs!confetiq, 1, 1)) = "A" Then
            tenro_conf = CStr(rs!confval)
        End If
        If UCase(rs!conftipo) = "EST" And UCase(Mid(rs!confetiq, 1, 1)) = "A" Then
            lista_estrnro_conf = lista_estrnro_conf & "," & CStr(rs!confval)
        End If

        rs.MoveNext
    Loop
    rs.Close
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "INDICADOR "
    Flog.writeline Espacios(Tabulador * 1) & "  A - Hs/ h promedio de capacitación por empleado por sector"
    Flog.writeline
    
    If tenro_conf = "0" Then
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
        Flog.writeline Espacios(Tabulador * 1) & " No se configuro un Tipo de estructura para EXCLUIR participantes"
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    End If
    
    If tenro_conf = "0" Then
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
        Flog.writeline Espacios(Tabulador * 1) & " No se configuraron las estructuras del tipo anterior para EXCLUIR participantes"
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    End If
   
   
    fecdesdeA_anioAnt = CDate(CStr(Day(fecdesdeA)) & "/" & CStr(Month(fecdesdeA)) & "/" & CStr(Year(fecdesdeA) - 1))
    fechastaA_anioAnt = CDate(CStr(Day(fechastaA)) & "/" & CStr(Month(fechastaA)) & "/" & CStr(Year(fechastaA) - 1))
    AnioIni_anioAnt = AnioIni - 1
    '------------------------------------------------------------------
    'Busco los datos del año indicado en el filtro
    '------------------------------------------------------------------
    If tipo3 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo3
        If estruc3 <> 0 And estruc3 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc3
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    ElseIf tipo2 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo2
        If estruc2 <> 0 And estruc2 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc2
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    Else
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo1
        If estruc1 <> 0 And estruc1 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc1
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    End If
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
    End If
    IncPorc = (99 / cantRegistros)
    
    Do Until rs.EOF
        estrnro_aux = IIf(IsNull(rs!estrnro), 0, rs!estrnro)
        estrdabr_aux = IIf(IsNull(rs!estrdabr), "", rs!estrdabr)
        
        '*****************************************************************
        ' Calculo para el año indicado en el filtro
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.estevenro, cap_evento.evecodext, cap_evento.evedesabr, evefecini, evefecfin, evereqasi "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA)
        StrSql = StrSql & " AND (estevenro = 6 OR estevenro = 7)"
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            End If
            OpenRecordset StrSql, rs3
            Do Until rs3.EOF
                ternro_aux = IIf(IsNull(rs3!ternro), 0, rs3!ternro)
                empleg_aux = IIf(IsNull(rs3!empleg), 0, rs3!empleg)
                terape_aux = IIf(IsNull(rs3!terape), "", rs3!terape)
                ternom_aux = IIf(IsNull(rs3!ternom), "", rs3!ternom)
                terape2_aux = IIf(IsNull(rs3!terape2), "", rs3!terape2)
                ternom2_aux = IIf(IsNull(rs3!ternom2), "", rs3!ternom2)
                
                StrSql = " SELECT ternro FROM his_estructura WHERE his_estructura.tenro = " & tenro_conf
                StrSql = StrSql & " AND his_estructura.estrnro IN (" & lista_estrnro_conf & ")"
                StrSql = StrSql & " AND his_estructura.ternro = " & rs3!ternro
                StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(rs2!evefecini)
                StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(rs2!evefecini) & " OR his_estructura.htethasta IS NULL)"
                OpenRecordset StrSql, rs4
                If rs4.EOF Then
                    If CDate(rs2!evefecini) > fecdesdeA Then
                        fecCalculoIni = rs2!evefecini
                    Else
                        fecCalculoIni = fecdesdeA
                    End If
                    If CDate(rs2!evefecfin) < fechastaA Then
                        fecCalculoFin = rs2!evefecfin
                    Else
                        fecCalculoFin = fechastaA
                    End If
                    
                    
                    
                    'El empleado debe ser considerado
                    If CInt(rs2!evereqasi) = -1 Then
                        'Con asistencia
                        StrSql = " SELECT cap_calendario.calfecha, cap_asistencia.asievehorini horini, cap_asistencia.asievehorfin horfin FROM cap_eventomodulo"
                        StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro"
                        StrSql = StrSql & " INNER JOIN cap_asistencia ON cap_asistencia.calnro = cap_calendario.calnro"
                        StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & rs2!evenro
                        StrSql = StrSql & " AND cap_asistencia.ternro = " & rs3!ternro
                        StrSql = StrSql & " AND cap_calendario.calfecha >= " & ConvFecha(fecCalculoIni) & " AND cap_calendario.calfecha <= " & ConvFecha(fecCalculoFin)
                        StrSql = StrSql & " AND cap_asistencia.asipre = -1"
                        OpenRecordset StrSql, rs5
                        Do Until rs5.EOF
                            Call guardarInd_A(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni, rs5!calfecha, rs5!horini, rs5!horfin)
                            rs5.MoveNext
                        Loop
                        rs5.Close
                        
                    Else
                        'Sin asistencia
                        StrSql = " SELECT cap_calendario.calfecha, cap_calendario.calhordes horini, cap_calendario.calhorhas horfin FROM cap_eventomodulo"
                        StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro"
                        StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & rs2!evenro
                        StrSql = StrSql & " AND cap_calendario.calfecha >= " & ConvFecha(fecCalculoIni) & " AND cap_calendario.calfecha <= " & ConvFecha(fecCalculoFin)
                        OpenRecordset StrSql, rs5
                        Do Until rs5.EOF
                            Call guardarInd_A(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni, rs5!calfecha, rs5!horini, rs5!horfin)
                            rs5.MoveNext
                        Loop
                        rs5.Close
                    End If
                    
                End If
                rs4.Close
                
                rs3.MoveNext
            Loop
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        
        '*****************************************************************
        ' Calculo para el año ANTERIOR indicado en el filtro
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.estevenro, cap_estadoevento.estevedesabr, cap_evento.evecodext, cap_evento.evedesabr, evefecini, evefecfin, evereqasi "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " INNER JOIN cap_estadoevento ON cap_evento.estevenro = cap_estadoevento.estevenro "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA_anioAnt)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA_anioAnt)
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            End If
            OpenRecordset StrSql, rs3
            Do Until rs3.EOF
                ternro_aux = IIf(IsNull(rs3!ternro), 0, rs3!ternro)
                empleg_aux = IIf(IsNull(rs3!empleg), 0, rs3!empleg)
                terape_aux = IIf(IsNull(rs3!terape), "", rs3!terape)
                ternom_aux = IIf(IsNull(rs3!ternom), "", rs3!ternom)
                terape2_aux = IIf(IsNull(rs3!terape2), "", rs3!terape2)
                ternom2_aux = IIf(IsNull(rs3!ternom2), "", rs3!ternom2)
                
                StrSql = " SELECT ternro FROM his_estructura WHERE his_estructura.tenro = " & tenro_conf
                StrSql = StrSql & " AND his_estructura.estrnro IN (" & lista_estrnro_conf & ")"
                StrSql = StrSql & " AND his_estructura.ternro = " & rs3!ternro
                StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(rs2!evefecini)
                StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(rs2!evefecini) & " OR his_estructura.htethasta IS NULL)"
                OpenRecordset StrSql, rs4
                If rs4.EOF Then
                    If CDate(rs2!evefecini) > fecdesdeA_anioAnt Then
                        fecCalculoIni = rs2!evefecini
                    Else
                        fecCalculoIni = fecdesdeA_anioAnt
                    End If
                    If CDate(rs2!evefecfin) < fechastaA_anioAnt Then
                        fecCalculoFin = rs2!evefecfin
                    Else
                        fecCalculoFin = fechastaA_anioAnt
                    End If
                    
                    'El empleado debe ser considerado
                    If CInt(rs2!evereqasi) = -1 Then
                        'Con asistencia
                        StrSql = " SELECT cap_calendario.calfecha, cap_asistencia.asievehorini horini, cap_asistencia.asievehorfin horfin FROM cap_eventomodulo"
                        StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro"
                        StrSql = StrSql & " INNER JOIN cap_asistencia ON cap_asistencia.calnro = cap_calendario.calnro"
                        StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & rs2!evenro
                        StrSql = StrSql & " AND cap_asistencia.ternro = " & rs3!ternro
                        StrSql = StrSql & " AND cap_calendario.calfecha >= " & ConvFecha(fecCalculoIni) & " AND cap_calendario.calfecha <= " & ConvFecha(fecCalculoFin)
                        StrSql = StrSql & " AND cap_asistencia.asipre = -1"
                        OpenRecordset StrSql, rs5
                        Do Until rs5.EOF
                            Call guardarInd_A(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni_anioAnt, rs5!calfecha, rs5!horini, rs5!horfin)
                            rs5.MoveNext
                        Loop
                        rs5.Close
                        
                    Else
                        'Sin asistencia
                        StrSql = " SELECT cap_calendario.calfecha, cap_calendario.calhordes horini, cap_calendario.calhorhas horfin FROM cap_eventomodulo"
                        StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro"
                        StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & rs2!evenro
                        StrSql = StrSql & " AND cap_calendario.calfecha >= " & ConvFecha(fecCalculoIni) & " AND cap_calendario.calfecha <= " & ConvFecha(fecCalculoFin)
                        OpenRecordset StrSql, rs5
                        Do Until rs5.EOF
                            Call guardarInd_A(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni_anioAnt, rs5!calfecha, rs5!horini, rs5!horfin)
                            rs5.MoveNext
                        Loop
                        rs5.Close
                    End If
                    
                End If
                rs4.Close
                
                rs3.MoveNext
            Loop
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        '*****************************************************************
        ' Calculo la cantidad de empleados para el año y la estructura en cuestion. Esto es para cada cuatrimestre
        '*****************************************************************
        Call guardarInd_A_cant_empl(estrnro_aux, AnioIni, tipo1, estruc1, tipo2, estruc2, tipo3, estruc3)
        
        '*****************************************************************
        ' Calculo la cantidad de empleados para el año Anterior y la estructura en cuestion. Esto es para cada cuatrimestre
        '*****************************************************************
        Call guardarInd_A_cant_empl(estrnro_aux, AnioIni_anioAnt, tipo1, estruc1, tipo2, estruc2, tipo3, estruc3)
        
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rs.MoveNext
    Loop
    rs.Close
   
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If rs3.State = adStateOpen Then rs3.Close
    Set rs3 = Nothing
    If rs4.State = adStateOpen Then rs4.Close
    Set rs4 = Nothing
    If rs5.State = adStateOpen Then rs5.Close
    Set rs5 = Nothing
  
Exit Sub

ME_Local:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub

Private Sub Generar_Datos_B(ByVal AnioIni As Integer, ByVal fecdesdeA As Date, ByVal fechastaA As Date, ByVal fecdesdeC As Date, ByVal fechastaC As Date, ByVal tipo1 As Integer, ByVal estruc1 As Long, ByVal tipo2 As Integer, ByVal estruc2 As Long, ByVal tipo3 As Integer, ByVal estruc3 As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera los datos de Indicador B
' Autor      : FAF
' Fecha      : 21/08/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim cantRegistros As Long

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
    
Dim tenro_conf As String
Dim lista_estrnro_conf As String
Dim fecdesdeA_anioAnt As Date
Dim fechastaA_anioAnt As Date
Dim fecCalculoIni As Date
Dim fecCalculoFin As Date
Dim AnioIni_anioAnt As Integer
Dim NroReporte As Integer
Dim estrnro_aux
Dim estrdabr_aux
Dim ternro_aux
Dim empleg_aux
Dim terape_aux
Dim ternom_aux
Dim terape2_aux
Dim ternom2_aux
Dim evenro_aux
Dim evecodext_aux
Dim evedesabr_aux
Dim evereqasi_aux


    On Error GoTo ME_Local
      
    'Configuracion del Reporte
    NroReporte = 209
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    lista_estrnro_conf = "0"
    tenro_conf = "0"
    Do Until rs.EOF
        If UCase(rs!conftipo) = "TE" And UCase(Mid(rs!confetiq, 1, 1)) = "B" Then
            tenro_conf = CStr(rs!confval)
        End If
        If UCase(rs!conftipo) = "EST" And UCase(Mid(rs!confetiq, 1, 1)) = "B" Then
            lista_estrnro_conf = lista_estrnro_conf & "," & CStr(rs!confval)
        End If

        rs.MoveNext
    Loop
    rs.Close
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "INDICADOR "
    Flog.writeline Espacios(Tabulador * 1) & "  B - Porcentaje del tiempo dedicado a capacitación por sector"
    Flog.writeline
    
    If tenro_conf = "0" Then
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
        Flog.writeline Espacios(Tabulador * 1) & " No se configuro un Tipo de estructura para la forma de Liquidación"
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    End If

    If tenro_conf = "0" Then
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
        Flog.writeline Espacios(Tabulador * 1) & " No se configuraron las estructuras indicando la cantidad de horas trabajadas."
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    End If
   
   
    fecdesdeA_anioAnt = CDate(CStr(Day(fecdesdeA)) & "/" & CStr(Month(fecdesdeA)) & "/" & CStr(Year(fecdesdeA) - 1))
    fechastaA_anioAnt = CDate(CStr(Day(fechastaA)) & "/" & CStr(Month(fechastaA)) & "/" & CStr(Year(fechastaA) - 1))
    AnioIni_anioAnt = AnioIni - 1
    '------------------------------------------------------------------
    'Busco los datos del año indicado en el filtro
    '------------------------------------------------------------------
    If tipo3 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo3
        If estruc3 <> 0 And estruc3 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc3
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    ElseIf tipo2 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo2
        If estruc2 <> 0 And estruc2 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc2
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    Else
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo1
        If estruc1 <> 0 And estruc1 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc1
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    End If
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
    End If
    IncPorc = (99 / cantRegistros)
    
    Do Until rs.EOF
        estrnro_aux = IIf(IsNull(rs!estrnro), 0, rs!estrnro)
        estrdabr_aux = IIf(IsNull(rs!estrdabr), "", rs!estrdabr)
        
        '*****************************************************************
        ' Calculo para el año indicado en el filtro
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.estevenro, cap_evento.evecodext, cap_evento.evedesabr, evefecini, evefecfin, evereqasi "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA)
        StrSql = StrSql & " AND (estevenro = 6 OR estevenro = 7)"
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            evenro_aux = rs2!evenro
            evecodext_aux = rs2!evecodext
            evedesabr_aux = rs2!evedesabr
            evereqasi_aux = rs2!evereqasi
            
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            End If
            OpenRecordset StrSql, rs3
            Do Until rs3.EOF
                ternro_aux = IIf(IsNull(rs3!ternro), 0, rs3!ternro)
                empleg_aux = IIf(IsNull(rs3!empleg), 0, rs3!empleg)
                terape_aux = IIf(IsNull(rs3!terape), "", rs3!terape)
                ternom_aux = IIf(IsNull(rs3!ternom), "", rs3!ternom)
                terape2_aux = IIf(IsNull(rs3!terape2), "", rs3!terape2)
                ternom2_aux = IIf(IsNull(rs3!ternom2), "", rs3!ternom2)
                
                If CDate(rs2!evefecini) > fecdesdeA Then
                    fecCalculoIni = rs2!evefecini
                Else
                    fecCalculoIni = fecdesdeA
                End If
                If CDate(rs2!evefecfin) < fechastaA Then
                    fecCalculoFin = rs2!evefecfin
                Else
                    fecCalculoFin = fechastaA
                End If
                
                If CInt(rs2!evereqasi) = -1 Then
                    'Con asistencia
                    StrSql = " SELECT cap_calendario.calfecha, cap_asistencia.asievehorini horini, cap_asistencia.asievehorfin horfin FROM cap_eventomodulo"
                    StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro"
                    StrSql = StrSql & " INNER JOIN cap_asistencia ON cap_asistencia.calnro = cap_calendario.calnro"
                    StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & rs2!evenro
                    StrSql = StrSql & " AND cap_asistencia.ternro = " & rs3!ternro
                    StrSql = StrSql & " AND cap_calendario.calfecha >= " & ConvFecha(fecCalculoIni) & " AND cap_calendario.calfecha <= " & ConvFecha(fecCalculoFin)
                    StrSql = StrSql & " AND cap_asistencia.asipre = -1"
                    OpenRecordset StrSql, rs5
                    Do Until rs5.EOF
                        Call guardarInd_B(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni, rs5!calfecha, rs5!horini, rs5!horfin, evenro_aux, evecodext_aux, evedesabr_aux, evereqasi_aux)
                        rs5.MoveNext
                    Loop
                    rs5.Close
                    
                Else
                    'Sin asistencia
                    StrSql = " SELECT cap_calendario.calfecha, cap_calendario.calhordes horini, cap_calendario.calhorhas horfin FROM cap_eventomodulo"
                    StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro"
                    StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & rs2!evenro
                    StrSql = StrSql & " AND cap_calendario.calfecha >= " & ConvFecha(fecCalculoIni) & " AND cap_calendario.calfecha <= " & ConvFecha(fecCalculoFin)
                    OpenRecordset StrSql, rs5
                    Do Until rs5.EOF
                        Call guardarInd_B(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni, rs5!calfecha, rs5!horini, rs5!horfin, evenro_aux, evecodext_aux, evedesabr_aux, evereqasi_aux)
                        rs5.MoveNext
                    Loop
                    rs5.Close
                End If
                    
                rs3.MoveNext
            Loop
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        
        '*****************************************************************
        ' Calculo para el año ANTERIOR indicado en el filtro
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.estevenro, cap_estadoevento.estevedesabr, cap_evento.evecodext, cap_evento.evedesabr, evefecini, evefecfin, evereqasi "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " INNER JOIN cap_estadoevento ON cap_evento.estevenro = cap_estadoevento.estevenro "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA_anioAnt)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA_anioAnt)
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            evenro_aux = rs2!evenro
            evecodext_aux = rs2!evecodext
            evedesabr_aux = rs2!evedesabr
            evereqasi_aux = rs2!evereqasi
            
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            End If
            OpenRecordset StrSql, rs3
            Do Until rs3.EOF
                ternro_aux = IIf(IsNull(rs3!ternro), 0, rs3!ternro)
                empleg_aux = IIf(IsNull(rs3!empleg), 0, rs3!empleg)
                terape_aux = IIf(IsNull(rs3!terape), "", rs3!terape)
                ternom_aux = IIf(IsNull(rs3!ternom), "", rs3!ternom)
                terape2_aux = IIf(IsNull(rs3!terape2), "", rs3!terape2)
                ternom2_aux = IIf(IsNull(rs3!ternom2), "", rs3!ternom2)
                
                If CDate(rs2!evefecini) > fecdesdeA_anioAnt Then
                    fecCalculoIni = rs2!evefecini
                Else
                    fecCalculoIni = fecdesdeA_anioAnt
                End If
                If CDate(rs2!evefecfin) < fechastaA_anioAnt Then
                    fecCalculoFin = rs2!evefecfin
                Else
                    fecCalculoFin = fechastaA_anioAnt
                End If
                    
                If CInt(rs2!evereqasi) = -1 Then
                    'Con asistencia
                    StrSql = " SELECT cap_calendario.calfecha, cap_asistencia.asievehorini horini, cap_asistencia.asievehorfin horfin FROM cap_eventomodulo"
                    StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro"
                    StrSql = StrSql & " INNER JOIN cap_asistencia ON cap_asistencia.calnro = cap_calendario.calnro"
                    StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & rs2!evenro
                    StrSql = StrSql & " AND cap_asistencia.ternro = " & rs3!ternro
                    StrSql = StrSql & " AND cap_calendario.calfecha >= " & ConvFecha(fecCalculoIni) & " AND cap_calendario.calfecha <= " & ConvFecha(fecCalculoFin)
                    StrSql = StrSql & " AND cap_asistencia.asipre = -1"
                    OpenRecordset StrSql, rs5
                    Do Until rs5.EOF
                        Call guardarInd_B(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni_anioAnt, rs5!calfecha, rs5!horini, rs5!horfin, evenro_aux, evecodext_aux, evedesabr_aux, evereqasi_aux)
                        rs5.MoveNext
                    Loop
                    rs5.Close
                    
                Else
                    'Sin asistencia
                    StrSql = " SELECT cap_calendario.calfecha, cap_calendario.calhordes horini, cap_calendario.calhorhas horfin FROM cap_eventomodulo"
                    StrSql = StrSql & " INNER JOIN cap_calendario ON cap_eventomodulo.evmonro = cap_calendario.evmonro"
                    StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & rs2!evenro
                    StrSql = StrSql & " AND cap_calendario.calfecha >= " & ConvFecha(fecCalculoIni) & " AND cap_calendario.calfecha <= " & ConvFecha(fecCalculoFin)
                    OpenRecordset StrSql, rs5
                    Do Until rs5.EOF
                        Call guardarInd_B(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni_anioAnt, rs5!calfecha, rs5!horini, rs5!horfin, evenro_aux, evecodext_aux, evedesabr_aux, evereqasi_aux)
                        rs5.MoveNext
                    Loop
                    rs5.Close
                End If
                    
                rs3.MoveNext
            Loop
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        '*****************************************************************
        ' Calculo la cantidad de horas trabajadas para cada empleado del año, cuatrimestre y la estructura en cuestion.
        ' Debe calcularse al final porque un empleado puede estar en 2 eventos en un mismo cuatrimestre.
        '*****************************************************************
        Call guardarInd_B_cant_hs_trab(estrnro_aux, AnioIni, tenro_conf, lista_estrnro_conf)
        
        '*****************************************************************
        ' Calculo la cantidad de horas trabajadas para cada empleado del año Anterior, cuatrimestre y la estructura en cuestion.
        ' Debe calcularse al final porque un empleado puede estar en 2 eventos en un mismo cuatrimestre.
        '*****************************************************************
        Call guardarInd_B_cant_hs_trab(estrnro_aux, AnioIni_anioAnt, tenro_conf, lista_estrnro_conf)
        
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rs.MoveNext
    Loop
    rs.Close
   
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If rs3.State = adStateOpen Then rs3.Close
    Set rs3 = Nothing
    If rs4.State = adStateOpen Then rs4.Close
    Set rs4 = Nothing
    If rs5.State = adStateOpen Then rs5.Close
    Set rs5 = Nothing
  
Exit Sub

ME_Local:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub

Private Sub Generar_Datos_C(ByVal AnioIni As Integer, ByVal fecdesdeA As Date, ByVal fechastaA As Date, ByVal fecdesdeC As Date, ByVal fechastaC As Date, ByVal tipo1 As Integer, ByVal estruc1 As Long, ByVal tipo2 As Integer, ByVal estruc2 As Long, ByVal tipo3 As Integer, ByVal estruc3 As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera los datos de Indicador C
' Autor      : FAF
' Fecha      : 21/08/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim cantRegistros As Long

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
    
Dim fecdesdeA_anioAnt As Date
Dim fechastaA_anioAnt As Date
Dim AnioIni_anioAnt As Integer
Dim estrnro_aux
Dim estrdabr_aux
Dim evenro_aux
Dim evecodext_aux
Dim evedesabr_aux
Dim estevedesabr_aux

    On Error GoTo ME_Local
      
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "INDICADOR "
    Flog.writeline Espacios(Tabulador * 1) & "  C - Cumplimiento Plan de Capacitación"
    Flog.writeline
    
    fecdesdeA_anioAnt = CDate(CStr(Day(fecdesdeA)) & "/" & CStr(Month(fecdesdeA)) & "/" & CStr(Year(fecdesdeA) - 1))
    fechastaA_anioAnt = CDate(CStr(Day(fechastaA)) & "/" & CStr(Month(fechastaA)) & "/" & CStr(Year(fechastaA) - 1))
    AnioIni_anioAnt = AnioIni - 1
    '------------------------------------------------------------------
    'Busco los datos del año indicado en el filtro
    '------------------------------------------------------------------
    If tipo3 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo3
        If estruc3 <> 0 And estruc3 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc3
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    ElseIf tipo2 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo2
        If estruc2 <> 0 And estruc2 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc2
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    Else
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo1
        If estruc1 <> 0 And estruc1 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc1
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    End If
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
    End If
    IncPorc = (99 / cantRegistros)
    
    Do Until rs.EOF
        estrnro_aux = IIf(IsNull(rs!estrnro), 0, rs!estrnro)
        estrdabr_aux = IIf(IsNull(rs!estrdabr), "", rs!estrdabr)
        
        '*****************************************************************
        ' Calculo para el año indicado en el filtro
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.estevenro, cap_estadoevento.estevedesabr, cap_evento.evecodext, cap_evento.evedesabr, evefecini, evefecfin, evereqasi "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " INNER JOIN cap_estadoevento ON cap_evento.estevenro = cap_estadoevento.estevenro "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA)
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            evenro_aux = IIf(IsNull(rs2!evenro), 0, rs2!evenro)
            evecodext_aux = IIf(IsNull(rs2!evecodext), "", rs2!evecodext)
            evedesabr_aux = IIf(IsNull(rs2!evedesabr), "", rs2!evedesabr)
            estevedesabr_aux = IIf(IsNull(rs2!estevedesabr), "", rs2!estevedesabr)
            
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            End If
            OpenRecordset StrSql, rs3
            If Not rs3.EOF Then
                If rs2!estevenro = 6 Or rs2!estevenro = 7 Then
                    Call guardarInd_C(estrnro_aux, estrdabr_aux, evenro_aux, evecodext_aux, evedesabr_aux, estevedesabr_aux, AnioIni, True)
                Else
                    Call guardarInd_C(estrnro_aux, estrdabr_aux, evenro_aux, evecodext_aux, evedesabr_aux, estevedesabr_aux, AnioIni, False)
                End If
            End If
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        
        '*****************************************************************
        ' Calculo para el año anterior
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.estevenro, cap_estadoevento.estevedesabr, cap_evento.evecodext, cap_evento.evedesabr, evefecini, evefecfin, evereqasi "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " INNER JOIN cap_estadoevento ON cap_evento.estevenro = cap_estadoevento.estevenro "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA_anioAnt)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA_anioAnt)
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            evenro_aux = IIf(IsNull(rs2!evenro), 0, rs2!evenro)
            evecodext_aux = IIf(IsNull(rs2!evecodext), "", rs2!evecodext)
            evedesabr_aux = IIf(IsNull(rs2!evedesabr), "", rs2!evedesabr)
            estevedesabr_aux = IIf(IsNull(rs2!estevedesabr), "", rs2!estevedesabr)
            
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            End If
            OpenRecordset StrSql, rs3
            If Not rs3.EOF Then
                If rs2!estevenro = 6 Or rs2!estevenro = 7 Then
                    Call guardarInd_C(estrnro_aux, estrdabr_aux, evenro_aux, evecodext_aux, evedesabr_aux, estevedesabr_aux, AnioIni_anioAnt, True)
                Else
                    Call guardarInd_C(estrnro_aux, estrdabr_aux, evenro_aux, evecodext_aux, evedesabr_aux, estevedesabr_aux, AnioIni_anioAnt, False)
                End If
            End If
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        
        'Actualizo el estado del proceso
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rs.MoveNext
    Loop
    rs.Close
    
    
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If rs3.State = adStateOpen Then rs3.Close
    Set rs3 = Nothing
    If rs4.State = adStateOpen Then rs4.Close
    Set rs4 = Nothing
  
Exit Sub

ME_Local:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub

Private Sub Generar_Datos_D(ByVal AnioIni As Integer, ByVal fecdesdeA As Date, ByVal fechastaA As Date, ByVal fecdesdeC As Date, ByVal fechastaC As Date, ByVal tipo1 As Integer, ByVal estruc1 As Long, ByVal tipo2 As Integer, ByVal estruc2 As Long, ByVal tipo3 As Integer, ByVal estruc3 As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera los datos de Indicador D
' Autor      : FAF
' Fecha      : 21/08/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim cantRegistros As Long

    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
        
    Dim tenro_conf As String
    Dim lista_estrnro_conf As String
    Dim fecdesdeA_anioAnt As Date
    Dim fechastaA_anioAnt As Date

Dim fecCalculoIni As Date
Dim fecCalculoFin As Date
Dim AnioIni_anioAnt As Integer
Dim tot_eve
Dim tot_asis
Dim porc_tot
Dim porc
Dim estrnro_aux
Dim estrdabr_aux
Dim ternro_aux
Dim evenro_aux
Dim evecodext_aux
Dim evedesabr_aux
Dim evefecini_aux As Date

    On Error GoTo ME_Local
      
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "INDICADOR "
    Flog.writeline Espacios(Tabulador * 1) & "  D - Evaluaciones del curso"
    Flog.writeline
    
    '------------------------------------------------------------------
    'Busco los datos del año indicado en el filtro
    '------------------------------------------------------------------
    If tipo1 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo1
        If estruc1 <> 0 And estruc1 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc1
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    ElseIf tipo2 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo2
        If estruc2 <> 0 And estruc2 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc2
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    Else
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo3
        If estruc3 <> 0 And estruc3 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc3
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    End If
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
    End If
    IncPorc = (99 / cantRegistros)
    
    Do Until rs.EOF
        estrnro_aux = IIf(IsNull(rs!estrnro), 0, rs!estrnro)
        estrdabr_aux = IIf(IsNull(rs!estrdabr), "", rs!estrdabr)
        
        '*****************************************************************
        ' Calculo para el año indicado en el filtro
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.evenro, cap_evento.evecodext, cap_evento.evedesabr, evefecini, evefecfin, evereqasi, eveporasi "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA)
        StrSql = StrSql & " AND (estevenro = 6 OR estevenro = 7)"
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            evenro_aux = IIf(IsNull(rs2!evenro), 0, rs2!evenro)
            evecodext_aux = IIf(IsNull(rs2!evecodext), "", rs2!evecodext)
            evedesabr_aux = IIf(IsNull(rs2!evedesabr), "", rs2!evedesabr)
            evefecini_aux = IIf(IsNull(rs2!evefecini), "", rs2!evefecini)
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            End If
            OpenRecordset StrSql, rs3
            Do Until rs3.EOF
                ternro_aux = IIf(IsNull(rs3!ternro), 0, rs3!ternro)
                
                Call guardarInd_D(estrnro_aux, estrdabr_aux, ternro_aux, AnioIni, evefecini_aux, evenro_aux, evecodext_aux, evedesabr_aux)
                    
                rs3.MoveNext
            Loop
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rs.MoveNext
    Loop
    rs.Close

Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If rs3.State = adStateOpen Then rs3.Close
    Set rs3 = Nothing
    'If rs4.State = adStateOpen Then rs4.Close
    'Set rs4 = Nothing
    'If rs5.State = adStateOpen Then rs5.Close
    'Set rs5 = Nothing
  
Exit Sub

ME_Local:
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline

End Sub

Private Sub Generar_Datos_E(ByVal AnioIni As Integer, ByVal fecdesdeA As Date, ByVal fechastaA As Date, ByVal fecdesdeC As Date, ByVal fechastaC As Date, ByVal tipo1 As Integer, ByVal estruc1 As Long, ByVal tipo2 As Integer, ByVal estruc2 As Long, ByVal tipo3 As Integer, ByVal estruc3 As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera los datos de Indicador E
' Autor      : FAF
' Fecha      : 21/08/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Dim cantRegistros As Long
    
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
        
    Dim tenro_conf As String
    Dim lista_estrnro_conf As String
    Dim fecdesdeA_anioAnt As Date
    Dim fechastaA_anioAnt As Date
    

Dim fecCalculoIni As Date
Dim fecCalculoFin As Date
Dim AnioIni_anioAnt As Integer
Dim tot_eve
Dim tot_asis
Dim porc_tot
Dim porc
Dim estrnro_aux
Dim estrdabr_aux
Dim ternro_aux
Dim evenro_aux
Dim evecodext_aux
Dim evedesabr_aux
Dim evefecini_aux As Date

    On Error GoTo ME_Local
      
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "INDICADOR "
    Flog.writeline Espacios(Tabulador * 1) & "  E - Evaluacion del Alumno"
    Flog.writeline
    
'    fecdesdeA_anioAnt = CDate(CStr(Day(fecdesdeA)) & "/" & CStr(Month(fecdesdeA)) & "/" & CStr(Year(fecdesdeA) - 1))
'    fechastaA_anioAnt = CDate(CStr(Day(fechastaA)) & "/" & CStr(Month(fechastaA)) & "/" & CStr(Year(fechastaA) - 1))
'    AnioIni_anioAnt = AnioIni - 1
    '------------------------------------------------------------------
    'Busco los datos del año indicado en el filtro
    '------------------------------------------------------------------
    If tipo1 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo1
        If estruc1 <> 0 And estruc1 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc1
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    ElseIf tipo2 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo2
        If estruc2 <> 0 And estruc2 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc2
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    Else
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo3
        If estruc3 <> 0 And estruc3 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc3
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    End If
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
    End If
    IncPorc = (99 / cantRegistros)
    
    Do Until rs.EOF
        estrnro_aux = IIf(IsNull(rs!estrnro), 0, rs!estrnro)
        estrdabr_aux = IIf(IsNull(rs!estrdabr), "", rs!estrdabr)
        
        '*****************************************************************
        ' Calculo para el año indicado en el filtro
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.evenro, cap_evento.evecodext, cap_evento.evedesabr, evefecini, evefecfin, evereqasi, eveporasi "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA)
        StrSql = StrSql & " AND (estevenro = 6 OR estevenro = 7)"
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            evenro_aux = IIf(IsNull(rs2!evenro), 0, rs2!evenro)
            evecodext_aux = IIf(IsNull(rs2!evecodext), "", rs2!evecodext)
            evedesabr_aux = IIf(IsNull(rs2!evedesabr), "", rs2!evedesabr)
            evefecini_aux = IIf(IsNull(rs2!evefecini), "", rs2!evefecini)
            
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            End If
            OpenRecordset StrSql, rs3
            Do Until rs3.EOF
                ternro_aux = IIf(IsNull(rs3!ternro), 0, rs3!ternro)
                
                Call guardarInd_E(estrnro_aux, estrdabr_aux, ternro_aux, AnioIni, evefecini_aux, evenro_aux, evecodext_aux, evedesabr_aux)
                    
                rs3.MoveNext
            Loop
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rs.MoveNext
    Loop
    rs.Close

Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If rs3.State = adStateOpen Then rs3.Close
    Set rs3 = Nothing
    'If rs4.State = adStateOpen Then rs4.Close
    'Set rs4 = Nothing
    'If rs5.State = adStateOpen Then rs5.Close
    'Set rs5 = Nothing
  
Exit Sub

ME_Local:
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline

End Sub

Private Sub Generar_Datos_F(ByVal AnioIni As Integer, ByVal fecdesdeA As Date, ByVal fechastaA As Date, ByVal fecdesdeC As Date, ByVal fechastaC As Date, ByVal tipo1 As Integer, ByVal estruc1 As Long, ByVal tipo2 As Integer, ByVal estruc2 As Long, ByVal tipo3 As Integer, ByVal estruc3 As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera los datos de Indicador A
' Autor      : FAF
' Fecha      : 21/08/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim cantRegistros As Long

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
    
Dim fecdesdeA_anioAnt As Date
Dim fechastaA_anioAnt As Date
Dim fecCalculoIni As Date
Dim fecCalculoFin As Date
Dim AnioIni_anioAnt As Integer
Dim tot_eve
Dim tot_asis
Dim porc_tot
Dim porc
Dim estrnro_aux
Dim estrdabr_aux
Dim ternro_aux
Dim empleg_aux
Dim terape_aux
Dim ternom_aux
Dim terape2_aux
Dim ternom2_aux
Dim evenro_aux
Dim evecodext_aux
Dim evedesabr_aux

    On Error GoTo ME_Local
      
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "INDICADOR "
    Flog.writeline Espacios(Tabulador * 1) & "  F - Asistencia"
    Flog.writeline
    
'    fecdesdeA_anioAnt = CDate(CStr(Day(fecdesdeA)) & "/" & CStr(Month(fecdesdeA)) & "/" & CStr(Year(fecdesdeA) - 1))
'    fechastaA_anioAnt = CDate(CStr(Day(fechastaA)) & "/" & CStr(Month(fechastaA)) & "/" & CStr(Year(fechastaA) - 1))
'    AnioIni_anioAnt = AnioIni - 1
    '------------------------------------------------------------------
    'Busco los datos del año indicado en el filtro
    '------------------------------------------------------------------
    If tipo3 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo3
        If estruc3 <> 0 And estruc3 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc3
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    ElseIf tipo2 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo2
        If estruc2 <> 0 And estruc2 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc2
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    Else
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo1
        If estruc1 <> 0 And estruc1 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc1
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    End If
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
    End If
    IncPorc = (99 / cantRegistros)
    
    Do Until rs.EOF
        estrnro_aux = IIf(IsNull(rs!estrnro), 0, rs!estrnro)
        estrdabr_aux = IIf(IsNull(rs!estrdabr), "", rs!estrdabr)
        
        '*****************************************************************
        ' Calculo para el año ANTERIOR indicado en el filtro
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.evenro, cap_evento.evecodext, cap_evento.evedesabr, evefecini, evefecfin, evereqasi, eveporasi "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA)
        StrSql = StrSql & " AND (estevenro = 6 OR estevenro = 7)"
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF
            evenro_aux = IIf(IsNull(rs2!evenro), 0, rs2!evenro)
            evecodext_aux = IIf(IsNull(rs2!evecodext), "", rs2!evecodext)
            evedesabr_aux = IIf(IsNull(rs2!evedesabr), "", rs2!evedesabr)
            
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            End If
            OpenRecordset StrSql, rs3
            Do Until rs3.EOF
                ternro_aux = IIf(IsNull(rs3!ternro), 0, rs3!ternro)
                empleg_aux = IIf(IsNull(rs3!empleg), 0, rs3!empleg)
                terape_aux = IIf(IsNull(rs3!terape), "", rs3!terape)
                ternom_aux = IIf(IsNull(rs3!ternom), "", rs3!ternom)
                terape2_aux = IIf(IsNull(rs3!terape2), "", rs3!terape2)
                ternom2_aux = IIf(IsNull(rs3!ternom2), "", rs3!ternom2)
                
                If CInt(rs2!evereqasi) = -1 Then
                    'Con asistencia
                    StrSql = " SELECT cap_calendario.calnro, calhordes, calhorhas "
                    StrSql = StrSql & " FROM cap_eventomodulo "
                    StrSql = StrSql & " INNER JOIN cap_calendario ON cap_calendario.evmonro = cap_eventomodulo.evmonro "
                    StrSql = StrSql & " INNER JOIN cap_partcal ON cap_partcal.calnro = cap_calendario.calnro AND cap_partcal.ternro = " & rs3!ternro
                    StrSql = StrSql & " WHERE cap_eventomodulo.evenro = " & rs2!evenro
                    OpenRecordset StrSql, rs4
    
                    tot_eve = 0
                    tot_asis = 0
                    porc_tot = rs2!eveporasi

                    Do Until rs4.EOF
                        tot_eve = tot_eve + DateDiff("n", CDate(Mid(rs4!calhordes, 1, 2) & ":" & Mid(rs4!calhordes, 3, 2)), CDate(Mid(rs4!calhorhas, 1, 2) & ":" & Mid(rs4!calhorhas, 3, 2)))
        
                        ' Calculo el numero total de minutos que Asistio el Empleado
                        StrSql = " SELECT asipre "
                        StrSql = StrSql & " FROM cap_asistencia "
                        StrSql = StrSql & " WHERE cap_asistencia.ternro = " & rs3!ternro & " AND cap_asistencia.calnro = " & rs4!calnro
                        OpenRecordset StrSql, rs5
                        If Not rs5.EOF Then
                            If CInt(rs5!asipre) = -1 Then
                                tot_asis = tot_asis + DateDiff("n", CDate(Mid(rs4!calhordes, 1, 2) & ":" & Mid(rs4!calhordes, 3, 2)), CDate(Mid(rs4!calhorhas, 1, 2) & ":" & Mid(rs4!calhorhas, 3, 2)))
                            End If
                        End If
                        rs5.Close
                        rs4.MoveNext
        
                    Loop
                    rs4.Close
                    
                    If tot_eve = 0 Then
                        porc = 0
                    Else
                        porc = tot_asis * 100 / tot_eve
                    End If

                    If CDbl(porc) >= CDbl(porc_tot) Then
                        Call guardarInd_F(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni, Empty, -1, -1, evenro_aux, evecodext_aux, evedesabr_aux)
                    Else
                        Call guardarInd_F(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni, Empty, -1, 0, evenro_aux, evecodext_aux, evedesabr_aux)
                    End If
                                        
                Else
                    'Sin asistencia
                    Call guardarInd_F(estrnro_aux, estrdabr_aux, ternro_aux, empleg_aux, terape_aux, terape2_aux, ternom_aux, ternom2_aux, AnioIni, Empty, 0, -1, evenro_aux, evecodext_aux, evedesabr_aux)
                End If
                    
                rs3.MoveNext
            Loop
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rs.MoveNext
    Loop
    rs.Close

Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If rs3.State = adStateOpen Then rs3.Close
    Set rs3 = Nothing
    If rs4.State = adStateOpen Then rs4.Close
    Set rs4 = Nothing
    If rs5.State = adStateOpen Then rs5.Close
    Set rs5 = Nothing
  
Exit Sub

ME_Local:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
End Sub

Private Sub Generar_Datos_G(ByVal AnioIni As Integer, ByVal fecdesdeA As Date, ByVal fechastaA As Date, ByVal fecdesdeC As Date, ByVal fechastaC As Date, ByVal tipo1 As Integer, ByVal estruc1 As Long, ByVal tipo2 As Integer, ByVal estruc2 As Long, ByVal tipo3 As Integer, ByVal estruc3 As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que genera los datos de Indicador G
' Autor      : Lisandro Moro
' Fecha      : 06/09/2007
' Ultima Mod :
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Dim cantRegistros As Long
    
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
        
    Dim tenro_conf As String
    Dim lista_estrnro_conf As String
    
    Dim NroReporte As Integer
    Dim lista_itenro_conf As String
    Dim CostoTotalEvento As Double

    Dim CostoInd As Double
    Dim CantRealAlu As Integer
    Dim DescCostos As String
    
    Dim estrnro_aux As Long
    Dim estrdabr_aux As String
    Dim evedesabr_aux As String
    Dim evenro_aux As Long
    Dim ternro_aux As Long
    Dim empleg_aux As Long
    Dim terape_aux As String
    Dim ternom_aux As String
    Dim terape2_aux As String
    Dim ternom2_aux As String
    Dim evecodext_aux As String
    
    
    On Error GoTo ME_Local
      
    'Configuracion del Reporte
    NroReporte = 209
    StrSql = "SELECT * FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & NroReporte
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    lista_itenro_conf = "0"
    
    Do Until rs.EOF
        
        If UCase(rs!conftipo) = "IT" And UCase(Mid(rs!confetiq, 1, 1)) = "G" Then  'Items a no tener en cuenta (16 - 17)
            lista_itenro_conf = lista_itenro_conf & "," & CStr(rs!confval)
        End If
        
        rs.MoveNext
    Loop
    rs.Close

    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "INDICADOR "
    Flog.writeline Espacios(Tabulador * 1) & "  G - Inversión en Capacitación por Empleado"
    Flog.writeline
    
    If lista_itenro_conf = "0" Then
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
        Flog.writeline Espacios(Tabulador * 1) & " No se configuro un Item para excluir de los costos         "
        Flog.writeline Espacios(Tabulador * 1) & "************************************************************"
    End If
    
    'fecdesdeA_anioAnt = CDate(CStr(Day(fecdesdeA)) & "/" & CStr(Month(fecdesdeA)) & "/" & CStr(Year(fecdesdeA) - 1))
    'fechastaA_anioAnt = CDate(CStr(Day(fechastaA)) & "/" & CStr(Month(fechastaA)) & "/" & CStr(Year(fechastaA) - 1))
    'AnioIni_anioAnt = AnioIni - 1
    '------------------------------------------------------------------
    'Busco los datos del año indicado en el filtro
    '------------------------------------------------------------------
    If tipo3 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo3
        If estruc3 <> 0 And estruc3 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc3
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    ElseIf tipo2 <> 0 Then
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo2
        If estruc2 <> 0 And estruc2 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc2
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    Else
        StrSql = "SELECT estructura.estrnro, estructura.estrdabr "
        StrSql = StrSql & " FROM estructura WHERE tenro = " & tipo1
        If estruc1 <> 0 And estruc1 <> -1 Then
            StrSql = StrSql & " AND estrnro = " & estruc1
        End If
        StrSql = StrSql & " ORDER BY estrdabr "
    End If
    OpenRecordset StrSql, rs
    
    'seteo de las variables de progreso
    Progreso = 0
    If Not rs.EOF Then
        cantRegistros = rs.RecordCount
        If cantRegistros = 0 Then
           cantRegistros = 1
           Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
        End If
    Else
        cantRegistros = 1
        Flog.writeline Espacios(Tabulador * 1) & "No se encontraron estructuras"
    End If
    IncPorc = (99 / cantRegistros)
    
    Do Until rs.EOF ' x estructuras
        '*****************************************************************
        ' Calculo para el año indicado en el filtro
        '*****************************************************************
        StrSql = "SELECT cap_evento.evenro, cap_evento.estevenro, cap_estadoevento.estevedesabr, "
        StrSql = StrSql & " cap_evento.evecodext, cap_evento.evedesabr, "
        StrSql = StrSql & " evefecini, evefecfin, evereqasi, "
        StrSql = StrSql & " evecostogral, evecostoind, evecanrealalu "
        StrSql = StrSql & " FROM cap_evento "
        StrSql = StrSql & " INNER JOIN cap_estadoevento ON cap_evento.estevenro = cap_estadoevento.estevenro "
        StrSql = StrSql & " WHERE evefecfin >= " & ConvFecha(fecdesdeA)
        StrSql = StrSql & " AND evefecini <= " & ConvFecha(fechastaA)
        StrSql = StrSql & " AND (( cap_evento.estevenro = 6 ) OR (cap_evento.estevenro = 7))"
        'Ver si no van los estado de los eventos en 6 y 7 - licho
        StrSql = StrSql & " ORDER BY evefecini "
        OpenRecordset StrSql, rs2
        Do Until rs2.EOF ' x eventos x estructuras
            
            ' EMPLEADOS
            If tipo3 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON cap_candidato.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est3.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON cap_candidato.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est2.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
            Else
                StrSql = "SELECT cap_candidato.ternro, empleado.empleg, empleado.terape, empleado.terape2, empleado.ternom, empleado.ternom2 "
                StrSql = StrSql & " FROM cap_candidato "
                StrSql = StrSql & " INNER JOIN empleado ON cap_candidato.ternro = empleado.ternro "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON cap_candidato.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(rs2!evefecini) & " AND (est1.htethasta >= " & ConvFecha(rs2!evefecini) & " OR est1.htethasta IS NULL)"
                'StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(fecdesdeA) & " AND (est1.htethasta >= " & ConvFecha(fecdesdeA) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & rs!estrnro
                StrSql = StrSql & " WHERE cap_candidato.evenro = " & rs2!evenro & " AND cap_candidato.conf = -1"
                'Debug.Print rs!estrnro & " " & rs2!evenro
            End If
            OpenRecordset StrSql, rs3
            Do Until rs3.EOF
                
                estrnro_aux = IIf(IsNull(rs!estrnro), 0, rs!estrnro)
                estrdabr_aux = IIf(IsNull(rs!estrdabr), 0, rs!estrdabr)
                evenro_aux = IIf(IsNull(rs2!evenro), 0, rs2!evenro)
                ternro_aux = IIf(IsNull(rs3!ternro), 0, rs3!ternro)
                empleg_aux = IIf(IsNull(rs3!empleg), 0, rs3!empleg)
                ternom_aux = IIf(IsNull(rs3!ternom), "", rs3!ternom)
                terape_aux = IIf(IsNull(rs3!terape), "", rs3!terape)
                terape2_aux = IIf(IsNull(rs3!terape2), "", rs3!terape2)
                ternom2_aux = IIf(IsNull(rs3!ternom2), "", rs3!ternom2)
                'evecodext_aux = IIf(IsNull(rs2!evecodext), "", rs2!evecodext)
                'evedesabr_aux = IIf(IsNull(rs2!evedesabr), "", rs2!evedesabr)
                
                Call guardarInd_G(estrnro_aux, estrdabr_aux, evenro_aux, ternro_aux, empleg_aux, ternom_aux, terape_aux, ternom2_aux, terape2_aux, AnioIni, lista_itenro_conf, fecdesdeA)
                
                rs3.MoveNext
            Loop
            rs3.Close
            
            rs2.MoveNext
        Loop
        rs2.Close
        
        
        TiempoAcumulado = GetTickCount
        cantRegistros = cantRegistros - 1
        Progreso = Progreso + IncPorc
        
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso
        StrSql = StrSql & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "'"
        StrSql = StrSql & ", bprcempleados ='" & CStr(cantRegistros) & "' WHERE bpronro = " & NroProceso
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
        
        rs.MoveNext
    Loop
    rs.Close

    
Fin:
    'Cierro y libero todo
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs2.State = adStateOpen Then rs2.Close
    Set rs2 = Nothing
    If rs3.State = adStateOpen Then rs3.Close
    Set rs3 = Nothing
    'If rs4.State = adStateOpen Then rs4.Close
    'Set rs4 = Nothing
    'If rs5.State = adStateOpen Then rs5.Close
    'Set rs5 = Nothing
  
Exit Sub

ME_Local:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline


End Sub

Private Sub guardarInd_A_cant_empl(ByVal estrnro As Long, ByVal Anio As Integer, ByVal tipo1 As Integer, ByVal estruc1 As Long, ByVal tipo2 As Integer, ByVal estruc2 As Long, ByVal tipo3 As Integer, ByVal estruc3 As Long)
Dim I
Dim totEmpl As Double
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
    
    On Error GoTo ME_guardarInd_A_cant_empl
    
    For I = 1 To 3
        ' Para cada uno de los 3 posibles cuatrimestres
        StrSql = " SELECT * FROM rep_ind_A WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND anio = " & Anio
        StrSql = StrSql & " AND cuatrimestre = " & I
        StrSql = StrSql & " AND estrnro = " & estrnro
        OpenRecordset StrSql, rsConsult
        If Not rsConsult.EOF Then
            ' Si se generaron datos, debo contar la cantidad de empleados a la fecha de inicio del cuatrimestre, respetando los niveles de esctructuras
            If I = 1 Then
                Fecha = CDate("01/01/" & CStr(Anio))
            End If
            If I = 2 Then
                Fecha = CDate("01/05/" & CStr(Anio))
            End If
            If I = 3 Then
                Fecha = CDate("01/09/" & CStr(Anio))
            End If
            
            If tipo3 <> 0 Then
                StrSql = "SELECT est1.ternro "
                StrSql = StrSql & " FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON empleado.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(Fecha) & " AND (est1.htethasta >= " & ConvFecha(Fecha) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON empleado.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(Fecha) & " AND (est2.htethasta >= " & ConvFecha(Fecha) & " OR est2.htethasta IS NULL)"
                If estruc2 <> 0 And estruc2 <> -1 Then
                    StrSql = StrSql & " AND est2.estrnro = " & estruc2
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est3 ON empleado.ternro = est3.ternro AND est3.tenro = " & tipo3
                StrSql = StrSql & " AND est3.htetdesde <= " & ConvFecha(Fecha) & " AND (est3.htethasta >= " & ConvFecha(Fecha) & " OR est3.htethasta IS NULL)"
                StrSql = StrSql & " AND est3.estrnro = " & estrnro
            ElseIf tipo2 <> 0 Then
                StrSql = "SELECT est1.ternro "
                StrSql = StrSql & " FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON empleado.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(Fecha) & " AND (est1.htethasta >= " & ConvFecha(Fecha) & " OR est1.htethasta IS NULL)"
                If estruc1 <> 0 And estruc1 <> -1 Then
                    StrSql = StrSql & " AND est1.estrnro = " & estruc1
                End If
                StrSql = StrSql & " INNER JOIN his_estructura est2 ON empleado.ternro = est2.ternro AND est2.tenro = " & tipo2
                StrSql = StrSql & " AND est2.htetdesde <= " & ConvFecha(Fecha) & " AND (est2.htethasta >= " & ConvFecha(Fecha) & " OR est2.htethasta IS NULL)"
                StrSql = StrSql & " AND est2.estrnro = " & estrnro
            Else
                StrSql = "SELECT est1.ternro "
                StrSql = StrSql & " FROM empleado "
                StrSql = StrSql & " INNER JOIN his_estructura est1 ON empleado.ternro = est1.ternro AND est1.tenro = " & tipo1
                StrSql = StrSql & " AND est1.htetdesde <= " & ConvFecha(Fecha) & " AND (est1.htethasta >= " & ConvFecha(Fecha) & " OR est1.htethasta IS NULL)"
                StrSql = StrSql & " AND est1.estrnro = " & estrnro
            End If
            
            OpenRecordset StrSql, rsConsult2
            If Not rsConsult2.EOF Then
                totEmpl = rsConsult2.RecordCount
                
                StrSql = "UPDATE rep_ind_A SET cantemp = " & totEmpl
                StrSql = StrSql & " WHERE bpronro = " & NroProceso
                StrSql = StrSql & " AND anio = " & Anio
                StrSql = StrSql & " AND cuatrimestre = " & I
                StrSql = StrSql & " AND estrnro = " & estrnro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            rsConsult2.Close
            
        End If
        
        rsConsult.Close
    Next
    
ME_Fin_A_cant_empl:
    Set rsConsult = Nothing
    Set rsConsult2 = Nothing
    Exit Sub

ME_guardarInd_A_cant_empl:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo ME_Fin_A_cant_empl
End Sub

Private Sub guardarInd_A(ByVal estrnro As Long, ByVal estrdabr As String, ByVal ternro As Long, ByVal empleg As Long, ByVal terape As String, ByVal terape2 As String, ByVal ternom As String, ByVal ternom2 As String, ByVal Anio As Integer, ByVal Fecha As Date, ByVal horini As String, ByVal horfin As String)
Dim totHoras As Double
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
'Dim sumar_emp As Integer
Dim cuatr As Integer
    
    On Error GoTo ME_guardarInd_A
    
    totHoras = DateDiff("n", CDate(Mid(horini, 1, 2) & ":" & Mid(horini, 3, 2)), CDate(Mid(horfin, 1, 2) & ":" & Mid(horfin, 3, 2)))
    totHoras = Replace(CStr(CSng(totHoras / 60)), ",", ".")
    
    If Month(Fecha) / 4 <= 1 Then
        cuatr = 1
    Else
        If Month(Fecha) / 4 > 1 And Month(Fecha) / 4 <= 2 Then
            cuatr = 2
        Else
            cuatr = 3
        End If
    End If

    StrSql = " SELECT * FROM rep_ind_A WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND anio = " & Anio
    StrSql = StrSql & " AND cuatrimestre = " & cuatr
    StrSql = StrSql & " AND estrnro = " & estrnro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        StrSql = "INSERT INTO rep_ind_A_det (bpronro, anio, cuatrimestre, estrnro, ternro, empleg, terape, terape2, ternom, ternom2, canths)"
        StrSql = StrSql & " VALUES (" & NroProceso
        StrSql = StrSql & ", " & Anio
        StrSql = StrSql & ", " & cuatr
        StrSql = StrSql & ", " & estrnro
        StrSql = StrSql & ", " & ternro
        StrSql = StrSql & ", " & empleg
        StrSql = StrSql & ", '" & terape & "'"
        StrSql = StrSql & ", '" & terape2 & "'"
        StrSql = StrSql & ", '" & ternom & "'"
        StrSql = StrSql & ", '" & ternom2 & "'"
        StrSql = StrSql & ", " & totHoras
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "INSERT INTO rep_ind_A (bpronro, anio, cuatrimestre, estrnro, estrdabr, canths, cantemp)"
        StrSql = StrSql & " VALUES (" & NroProceso
        StrSql = StrSql & ", " & Anio
        StrSql = StrSql & ", " & cuatr
        StrSql = StrSql & ", " & estrnro
        StrSql = StrSql & ", '" & estrdabr & "'"
        StrSql = StrSql & ", " & totHoras
        StrSql = StrSql & ", 0"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
    Else
        StrSql = " SELECT * FROM rep_ind_A_det WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND anio = " & Anio
        StrSql = StrSql & " AND cuatrimestre = " & cuatr
        StrSql = StrSql & " AND estrnro = " & estrnro
        StrSql = StrSql & " AND ternro = " & ternro
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            StrSql = "INSERT INTO rep_ind_A_det (bpronro, anio, cuatrimestre, estrnro, ternro, empleg, terape, terape2, ternom, ternom2, canths)"
            StrSql = StrSql & " VALUES (" & NroProceso
            StrSql = StrSql & ", " & Anio
            StrSql = StrSql & ", " & cuatr
            StrSql = StrSql & ", " & estrnro
            StrSql = StrSql & ", " & ternro
            StrSql = StrSql & ", " & empleg
            StrSql = StrSql & ", '" & terape & "'"
            StrSql = StrSql & ", '" & terape2 & "'"
            StrSql = StrSql & ", '" & ternom & "'"
            StrSql = StrSql & ", '" & ternom2 & "'"
            StrSql = StrSql & ", " & totHoras
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
'            sumar_emp = 1
        Else
            StrSql = "UPDATE rep_ind_A_det SET canths = " & rsConsult2!canths + totHoras
            StrSql = StrSql & " WHERE bpronro = " & NroProceso
            StrSql = StrSql & " AND anio = " & Anio
            StrSql = StrSql & " AND cuatrimestre = " & cuatr
            StrSql = StrSql & " AND estrnro = " & estrnro
            StrSql = StrSql & " AND ternro = " & ternro
            objConn.Execute StrSql, , adExecuteNoRecords
'            sumar_emp = 0
        End If
        rsConsult2.Close
        
        StrSql = "UPDATE rep_ind_A SET canths = " & rsConsult!canths + totHoras
'        StrSql = StrSql & ", cantemp = " & rsConsult!cantEmp + sumar_emp
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND anio = " & Anio
        StrSql = StrSql & " AND cuatrimestre = " & cuatr
        StrSql = StrSql & " AND estrnro = " & estrnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
    rsConsult.Close
    
ME_Fin_A:
    Set rsConsult = Nothing
    Set rsConsult2 = Nothing
    Exit Sub

ME_guardarInd_A:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo ME_Fin_A
End Sub

Private Sub guardarInd_B_cant_hs_trab(ByVal estrnro As Long, ByVal Anio As Integer, ByVal tenro_conf As String, ByVal lista_estrnro_conf As String) ', ByVal tipo1 As Integer, ByVal estruc1 As Long, ByVal tipo2 As Integer, ByVal estruc2 As Long, ByVal tipo3 As Integer, ByVal estruc3 As Long)
Dim I
Dim totEmpl As Double
Dim fecha_Calculo As Date
Dim cuatr_ant As Integer
Dim hs_Trab
Dim total_hs_Trab
Dim dias_trab
Dim dias_cuatr
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim rsConsult3 As New ADODB.Recordset
Dim fecha_ini As Date
Dim Fecha_Fin As Date
Dim Encontro As Boolean
Dim formaliq As String

    On Error GoTo ME_guardarInd_B_cant_hs_trab
    
    StrSql = " SELECT DISTINCT empleado.ternro, empleado.empleg, cuatrimestre FROM rep_ind_B_det "
    StrSql = StrSql & " INNER JOIN empleado ON rep_ind_B_det.ternro = empleado.ternro "
    StrSql = StrSql & "WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND anio = " & Anio
    StrSql = StrSql & " AND estrnro = " & estrnro
    StrSql = StrSql & " ORDER BY cuatrimestre "
    OpenRecordset StrSql, rsConsult
    If Not rsConsult.EOF Then
        cuatr_ant = rsConsult!cuatrimestre
    End If
    total_hs_Trab = 0
    Do Until rsConsult.EOF
        If cuatr_ant <> rsConsult!cuatrimestre Then
            StrSql = " UPDATE rep_ind_B SET "
            StrSql = StrSql & " canthstrab = " & total_hs_Trab
            StrSql = StrSql & " WHERE bpronro = " & NroProceso
            StrSql = StrSql & " AND anio = " & Anio
            StrSql = StrSql & " AND cuatrimestre = " & rsConsult!cuatrimestre
            StrSql = StrSql & " AND estrnro = " & estrnro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            total_hs_Trab = 0
        End If
        
        Select Case rsConsult!cuatrimestre
            Case 1
                fecha_ini = CDate("01/01/" & CStr(Anio))
                Fecha_Fin = CDate("30/04/" & CStr(Anio))
            Case 2
                fecha_ini = CDate("01/05/" & CStr(Anio))
                Fecha_Fin = CDate("31/08/" & CStr(Anio))
            Case 3
                fecha_ini = CDate("01/09/" & CStr(Anio))
                Fecha_Fin = CDate("31/12/" & CStr(Anio))
        End Select
        
        StrSql = "SELECT confrep.confetiq ,confrep.confval2, estructura.estrdabr "
        StrSql = StrSql & " FROM his_estructura est1 "
        StrSql = StrSql & " INNER JOIN estructura ON est1.estrnro = estructura.estrnro "
        StrSql = StrSql & " LEFT JOIN confrep ON est1.estrnro = confrep.confval "
        StrSql = StrSql & " WHERE est1.htetdesde <= " & ConvFecha(fecha_ini) & " AND (est1.htethasta >= " & ConvFecha(fecha_ini) & " OR est1.htethasta IS NULL)"
        StrSql = StrSql & " AND est1.estrnro IN (" & lista_estrnro_conf & ") AND est1.tenro = " & tenro_conf & " AND est1.ternro = " & rsConsult!ternro
        StrSql = StrSql & " AND confrep.conftipo = 'EST' "
        OpenRecordset StrSql, rsConsult2
        hs_Trab = 0
        formaliq = ""
        If rsConsult2.EOF Then
            Flog.writeline Espacios(Tabulador * 2) & " ** No se encontro el tipo de estructura " & tenro_conf & " para el empleado " & rsConsult!empleg & " para la fecha " & fecha_ini
        Else
            Encontro = False
            Do Until rsConsult2.EOF
                If UCase(Mid(rsConsult2!confetiq, 1, 1)) = "B" Then
                    formaliq = rsConsult2!estrdabr
                    hs_Trab = IIf(IsNull(rsConsult2!confval2), 0, rsConsult2!confval2)
                    
                    StrSql = " SELECT altfec, bajfec "
                    StrSql = StrSql & " FROM fases "
                    StrSql = StrSql & " WHERE empleado = " & rsConsult!ternro
                    StrSql = StrSql & " AND ((altfec >= " & ConvFecha(fecha_ini) & " AND altfec <= " & ConvFecha(Fecha_Fin) & ") "
                    StrSql = StrSql & " OR (bajfec >= " & ConvFecha(fecha_ini) & " AND bajfec <= " & ConvFecha(Fecha_Fin) & ")) "
                    StrSql = StrSql & " ORDER BY altfec "
                    OpenRecordset StrSql, rsConsult3
                    dias_trab = 0
                    Do Until rsConsult3.EOF
                        If CDate(rsConsult3!altfec) < fecha_ini Then
                            dias_trab = dias_trab + DateDiff("d", fecha_ini, CDate(rsConsult3!bajfec))
                        Else
                            If IsNull(rsConsult3!bajfec) Then
                                dias_trab = dias_trab + DateDiff("d", CDate(rsConsult3!altfec), Fecha_Fin)
                            Else
                                If CDate(rsConsult3!bajfec) > Fecha_Fin Then
                                    dias_trab = dias_trab + DateDiff("d", CDate(rsConsult3!altfec), Fecha_Fin)
                                Else
                                    dias_trab = dias_trab + DateDiff("d", CDate(rsConsult3!altfec), CDate(rsConsult3!bajfec))
                                End If
                            End If
                        End If
                        rsConsult3.MoveNext
                    Loop
                    rsConsult3.Close
                    
                    If dias_trab <> 0 Then
                        dias_cuatr = DateDiff("d", fecha_ini, Fecha_Fin)
                        hs_Trab = dias_trab * hs_Trab / dias_cuatr
                    End If
                    
                    Encontro = True
                End If
                rsConsult2.MoveNext
            Loop
            
            If Not Encontro Then
                Flog.writeline Espacios(Tabulador * 2) & " ** No se encontro el tipo de estructura " & tenro_conf & " para el empleado " & rsConsult!empleg & " para la fecha " & fecha_Calculo
            End If
        End If
        
        StrSql = " UPDATE rep_ind_B_det SET "
        StrSql = StrSql & " canthstrab = " & hs_Trab
        StrSql = StrSql & " ,formaliq = '" & formaliq & "'"
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND anio = " & Anio
        StrSql = StrSql & " AND cuatrimestre = " & rsConsult!cuatrimestre
        StrSql = StrSql & " AND estrnro = " & estrnro
        StrSql = StrSql & " AND ternro = " & rsConsult!ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        total_hs_Trab = total_hs_Trab + hs_Trab
        
        rsConsult2.Close
        
        cuatr_ant = rsConsult!cuatrimestre
        
        rsConsult.MoveNext
        
    Loop
    
    If total_hs_Trab <> 0 Then
        StrSql = " UPDATE rep_ind_B SET "
        StrSql = StrSql & " canthstrab = " & total_hs_Trab
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND anio = " & Anio
        StrSql = StrSql & " AND cuatrimestre = " & cuatr_ant
        StrSql = StrSql & " AND estrnro = " & estrnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        total_hs_Trab = 0
    End If
    
ME_Fin_guardarInd_B_cant_hs_trab:
    Set rsConsult = Nothing
    Set rsConsult2 = Nothing
    Exit Sub

ME_guardarInd_B_cant_hs_trab:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo ME_guardarInd_B_cant_hs_trab
End Sub

Private Sub guardarInd_B(ByVal estrnro As Long, ByVal estrdabr As String, ByVal ternro As Long, ByVal empleg As Long, ByVal terape As String, ByVal terape2 As String, ByVal ternom As String, ByVal ternom2 As String, ByVal Anio As Integer, ByVal Fecha As Date, ByVal horini As String, ByVal horfin As String, ByVal evenro As Integer, ByVal evecodext As String, ByVal evedesabr As String, ByVal evereqasi As Integer)
Dim totHoras As Double
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
'Dim sumar_emp As Integer
Dim cuatr As Integer
    
    On Error GoTo ME_guardarInd_B
    
    totHoras = DateDiff("n", CDate(Mid(horini, 1, 2) & ":" & Mid(horini, 3, 2)), CDate(Mid(horfin, 1, 2) & ":" & Mid(horfin, 3, 2)))
    totHoras = Replace(CStr(CSng(totHoras / 60)), ",", ".")
    
    If Month(Fecha) / 4 <= 1 Then
        cuatr = 1
    Else
        If Month(Fecha) / 4 > 1 And Month(Fecha) / 4 <= 2 Then
            cuatr = 2
        Else
            cuatr = 3
        End If
    End If
    
    StrSql = " SELECT * FROM rep_ind_B WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND anio = " & Anio
    StrSql = StrSql & " AND cuatrimestre = " & cuatr
    StrSql = StrSql & " AND estrnro = " & estrnro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        
        StrSql = "INSERT INTO rep_ind_B (bpronro, anio, cuatrimestre, estrnro, estrdabr, canthscap, canthstrab)"
        StrSql = StrSql & " VALUES (" & NroProceso
        StrSql = StrSql & ", " & Anio
        StrSql = StrSql & ", " & cuatr
        StrSql = StrSql & ", " & estrnro
        StrSql = StrSql & ", '" & estrdabr & "'"
        StrSql = StrSql & ", " & totHoras
        StrSql = StrSql & ", 0"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "INSERT INTO rep_ind_B_det (bpronro, anio, cuatrimestre, estrnro, ternro, empleg, terape, "
        StrSql = StrSql & "terape2, ternom, ternom2, evenro, evecodext, evedesabr, evereqasi, canthscap, canthstrab)"
        StrSql = StrSql & " VALUES (" & NroProceso
        StrSql = StrSql & ", " & Anio
        StrSql = StrSql & ", " & cuatr
        StrSql = StrSql & ", " & estrnro
        StrSql = StrSql & ", " & ternro
        StrSql = StrSql & ", " & empleg
        StrSql = StrSql & ", '" & terape & "'"
        StrSql = StrSql & ", '" & terape2 & "'"
        StrSql = StrSql & ", '" & ternom & "'"
        StrSql = StrSql & ", '" & ternom2 & "'"
        StrSql = StrSql & ", " & evenro
        StrSql = StrSql & ", '" & evecodext & "'"
        StrSql = StrSql & ", '" & evedesabr & "'"
        StrSql = StrSql & ", " & evereqasi
        StrSql = StrSql & ", " & totHoras
        StrSql = StrSql & ", 0"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = " SELECT * FROM rep_ind_B_det WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND anio = " & Anio
        StrSql = StrSql & " AND cuatrimestre = " & cuatr
        StrSql = StrSql & " AND estrnro = " & estrnro
        StrSql = StrSql & " AND ternro = " & ternro
        StrSql = StrSql & " AND evenro = " & evenro
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            StrSql = "INSERT INTO rep_ind_B_det (bpronro, anio, cuatrimestre, estrnro, ternro, empleg, terape, "
            StrSql = StrSql & "terape2, ternom, ternom2, evenro, evecodext, evedesabr, evereqasi, canthscap, canthstrab)"
            StrSql = StrSql & " VALUES (" & NroProceso
            StrSql = StrSql & ", " & Anio
            StrSql = StrSql & ", " & cuatr
            StrSql = StrSql & ", " & estrnro
            StrSql = StrSql & ", " & ternro
            StrSql = StrSql & ", " & empleg
            StrSql = StrSql & ", '" & terape & "'"
            StrSql = StrSql & ", '" & terape2 & "'"
            StrSql = StrSql & ", '" & ternom & "'"
            StrSql = StrSql & ", '" & ternom2 & "'"
            StrSql = StrSql & ", " & evenro
            StrSql = StrSql & ", '" & evecodext & "'"
            StrSql = StrSql & ", '" & evedesabr & "'"
            StrSql = StrSql & ", " & evereqasi
            StrSql = StrSql & ", " & totHoras
            StrSql = StrSql & ", 0"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
'            StrSql = "INSERT INTO rep_ind_B_det (bpronro, anio, cuatrimestre, estrnro, empleg,ternro, terape, terape2, ternom, ternom2, canths)"
'            StrSql = StrSql & " VALUES (" & NroProceso
'            StrSql = StrSql & ", " & Anio
'            StrSql = StrSql & ", " & cuatr
'            StrSql = StrSql & ", " & estrnro
'            StrSql = StrSql & ", " & empleg
'            StrSql = StrSql & ", " & ternro
'            StrSql = StrSql & ", '" & terape & "'"
'            StrSql = StrSql & ", '" & terape2 & "'"
'            StrSql = StrSql & ", '" & ternom & "'"
'            StrSql = StrSql & ", '" & ternom2 & "'"
'            StrSql = StrSql & ", " & totHoras
'            StrSql = StrSql & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE rep_ind_B_det SET canthscap = " & rsConsult2!canthscap + totHoras
            StrSql = StrSql & " WHERE bpronro = " & NroProceso
            StrSql = StrSql & " AND anio = " & Anio
            StrSql = StrSql & " AND cuatrimestre = " & cuatr
            StrSql = StrSql & " AND estrnro = " & estrnro
            StrSql = StrSql & " AND ternro = " & ternro
            StrSql = StrSql & " AND evenro = " & evenro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rsConsult2.Close
        
        StrSql = "UPDATE rep_ind_B SET canthscap = " & rsConsult!canthscap + totHoras
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND anio = " & Anio
        StrSql = StrSql & " AND cuatrimestre = " & cuatr
        StrSql = StrSql & " AND estrnro = " & estrnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
    rsConsult.Close
    
ME_Fin_B:
    Set rsConsult = Nothing
    Set rsConsult2 = Nothing
    Exit Sub

ME_guardarInd_B:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo ME_Fin_B
End Sub

Private Sub guardarInd_C(ByVal estrnro As Long, ByVal estrdabr As String, ByVal evenro As Long, ByVal evecodext As String, ByVal evedesabr As String, ByVal estevedesabr As String, ByVal Anio As Integer, ByVal eveej As Boolean)
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
    
    On Error GoTo ME_guardarInd_C
    

    StrSql = " SELECT * FROM rep_ind_C WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND anio = " & Anio
    StrSql = StrSql & " AND estrnro = " & estrnro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        StrSql = "INSERT INTO rep_ind_C (bpronro, anio, estrnro, estrdabr, canteveej, canteve)"
        StrSql = StrSql & " VALUES (" & NroProceso
        StrSql = StrSql & ", " & Anio
        StrSql = StrSql & ", " & estrnro
        StrSql = StrSql & ", '" & estrdabr & "'"
        If eveej Then
            StrSql = StrSql & ", 1"
        Else
            StrSql = StrSql & ", 0"
        End If
        StrSql = StrSql & ", 1"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = "INSERT INTO rep_ind_C_det (bpronro, anio, estrnro, estrdabr, evenro, evecodext, evedesabr, estevedesabr)"
        StrSql = StrSql & " VALUES (" & NroProceso
        StrSql = StrSql & ", " & Anio
        StrSql = StrSql & ", " & estrnro
        StrSql = StrSql & ", '" & estrdabr & "'"
        StrSql = StrSql & ", " & evenro
        StrSql = StrSql & ", '" & evecodext & "'"
        StrSql = StrSql & ", '" & evedesabr & "'"
        StrSql = StrSql & ", '" & estevedesabr & "'"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = " SELECT * FROM rep_ind_C_det WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND anio = " & Anio
        StrSql = StrSql & " AND estrnro = " & estrnro
        StrSql = StrSql & " AND evenro = " & evenro
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            StrSql = "INSERT INTO rep_ind_C_det (bpronro, anio, estrnro, estrdabr, evenro, evecodext, evedesabr, estevedesabr)"
            StrSql = StrSql & " VALUES (" & NroProceso
            StrSql = StrSql & ", " & Anio
            StrSql = StrSql & ", " & estrnro
            StrSql = StrSql & ", '" & estrdabr & "'"
            StrSql = StrSql & ", " & evenro
            StrSql = StrSql & ", '" & evecodext & "'"
            StrSql = StrSql & ", '" & evedesabr & "'"
            StrSql = StrSql & ", '" & estevedesabr & "'"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rsConsult2.Close
        
        StrSql = "UPDATE rep_ind_C SET canteveej = "
        If eveej Then
            StrSql = StrSql & rsConsult!canteveej + 1
        Else
            StrSql = StrSql & rsConsult!canteveej
        End If
        StrSql = StrSql & ", canteve = " & rsConsult!canteve + 1
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND anio = " & Anio
        StrSql = StrSql & " AND estrnro = " & estrnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
    rsConsult.Close
    
ME_Fin_C:
    Set rsConsult = Nothing
    Set rsConsult2 = Nothing
    Exit Sub

ME_guardarInd_C:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo ME_Fin_C
End Sub

Private Sub guardarInd_D(ByVal estrnro As Long, ByVal estrdabr As String, ByVal ternro As Long, ByVal Anio As Integer, ByVal Fecha As Date, ByVal evenro As Integer, ByVal evecodext As String, ByVal evedesabr As String)
    
    Dim rsConsult As New ADODB.Recordset
    Dim rsConsult2 As New ADODB.Recordset
    'Dim sumar_emp As Integer
    Dim cuatr As Integer
    Dim l_min, l_med, l_max As Double
    
    On Error GoTo ME_guardarInd_D
    
    If Month(Fecha) / 4 <= 1 Then
        cuatr = 1
    Else
        If Month(Fecha) / 4 > 1 And Month(Fecha) / 4 <= 2 Then
            cuatr = 2
        Else
            cuatr = 3
        End If
    End If


    StrSql = " SELECT * FROM rep_ind_D WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND anio = " & Anio
    StrSql = StrSql & " AND cuatrimestre = " & cuatr
    StrSql = StrSql & " AND estrnro = " & estrnro
    StrSql = StrSql & " AND evenro = " & evenro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        StrSql = " SELECT cap_factor.grufacnro, facdesabr, facnro, evedesabr  "
        StrSql = StrSql & ", SUM(resval) valor, COUNT(evaresnro) cant, grufacdesabr, (SUM(resval) / COUNT(evaresnro)) promedio "
        StrSql = StrSql & ", MIN(resval) minimo, MAX(resval) maximo "
        StrSql = StrSql & " FROM cap_evaluacion "
        StrSql = StrSql & " INNER JOIN cap_factor ON cap_factor.facnro = cap_evaluacion.evaentnro "
        StrSql = StrSql & " INNER JOIN cap_grupofactor ON cap_factor.grufacnro = cap_grupofactor.grufacnro "
        StrSql = StrSql & " INNER JOIN cap_resultado ON cap_resultado.resnro = cap_evaluacion.evaresnro "
        StrSql = StrSql & " INNER JOIN cap_evento ON cap_evaluacion.evenro = cap_evento.evenro "
        StrSql = StrSql & " WHERE cap_evaluacion.evatipo = 3 "
        StrSql = StrSql & " AND cap_evaluacion.evaorigen = 2 "
        StrSql = StrSql & " AND cap_evento.evenro = " & evenro
        StrSql = StrSql & " GROUP BY facdesabr, facnro, cap_factor.grufacnro, grufacdesabr, evedesabr ORDER BY facnro "
        OpenRecordset StrSql, rsConsult2
        If Not rsConsult2.EOF Then
            'StrSql = "INSERT INTO rep_ind_D (bpronro , Anio, cuatrimestre, estrnro, estrdabr, evenro, evedesabr, grufacnro, grufacdesabr, evasatmin, evasatmed, evasatmax)"
            StrSql = "INSERT INTO rep_ind_D (bpronro , Anio, cuatrimestre, estrnro, estrdabr, evenro, evedesabr, evasatmin, evasatmed, evasatmax)"
            StrSql = StrSql & " VALUES (" & NroProceso
            StrSql = StrSql & ", " & Anio
            StrSql = StrSql & ", " & cuatr
            StrSql = StrSql & ", " & estrnro
            StrSql = StrSql & ", '" & estrdabr & "'"
            StrSql = StrSql & ", " & evenro
            StrSql = StrSql & ", '" & rsConsult2("evedesabr") & "'"
            'StrSql = StrSql & ", '" & rsConsult2("grufacnro") & "'"
            'StrSql = StrSql & ", '" & rsConsult2("grufacdesabr") & "'"
            StrSql = StrSql & ", 0" ' & rsConsult2("minimo")
            StrSql = StrSql & ", 0" ' & rsConsult2("promedio")
            StrSql = StrSql & ", 0" ' & rsConsult2("maximo")
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
                        
            Do While Not rsConsult2.EOF
                'Inserto los valores de de los factores de satisfaccion
                StrSql = "INSERT INTO rep_ind_D_det ( bpronro , Anio, cuatrimestre, estrnro, estrdabr, evenro, grufacnro, grufacdesabr, facnro, facdesabr, evasatmin, evasatmed, evasatmax)"
                StrSql = StrSql & " VALUES (" & NroProceso
                StrSql = StrSql & ", " & Anio
                StrSql = StrSql & ", " & cuatr
                StrSql = StrSql & ", " & estrnro
                StrSql = StrSql & ", '" & estrdabr & "'"
                StrSql = StrSql & ", " & evenro
                StrSql = StrSql & ", " & rsConsult2("grufacnro")
                StrSql = StrSql & ", '" & rsConsult2("grufacdesabr") & "'"
                StrSql = StrSql & ", " & rsConsult2("facnro")
                StrSql = StrSql & ", '" & rsConsult2("facdesabr") & "'"
                StrSql = StrSql & ", " & rsConsult2("minimo")
                StrSql = StrSql & ", " & rsConsult2("promedio")
                StrSql = StrSql & ", " & rsConsult2("maximo")
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                rsConsult2.MoveNext
            Loop
            
        End If
        rsConsult2.Close
        
        'Actualizo los totales
        StrSql = " SELECT min(evasatmin) min, AVG(evasatmed) med ,MAX(evasatmax) max "
        StrSql = StrSql & " FROM rep_ind_D_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND cuatrimestre = " & cuatr
        StrSql = StrSql & " AND estrnro = " & estrnro
        StrSql = StrSql & " AND evenro = " & evenro
        OpenRecordset StrSql, rsConsult2
        If Not rsConsult2.EOF Then
            If IsNull(rsConsult2!Min) Then
                l_min = 0
            Else
                l_min = CDbl(rsConsult2!Min)
            End If
            If IsNull(rsConsult2!Med) Then
                l_med = 0
            Else
                l_med = CDbl(rsConsult2!Med)
            End If
            If IsNull(rsConsult2!Max) Then
                l_max = 0
            Else
                l_max = CDbl(rsConsult2!Max)
            End If
            
            StrSql = "UPDATE rep_ind_D SET "
            StrSql = StrSql & " evasatmin = " & l_min
            StrSql = StrSql & ", evasatmed = " & l_med
            StrSql = StrSql & ", evasatmax = " & l_max
            StrSql = StrSql & " WHERE bpronro = " & NroProceso
            StrSql = StrSql & " AND cuatrimestre = " & cuatr
            StrSql = StrSql & " AND estrnro = " & estrnro
            StrSql = StrSql & " AND evenro = " & evenro
            objConn.Execute StrSql, , adExecuteNoRecords
        
        End If
    
    End If
    rsConsult.Close
    
ME_Fin_D:
    Set rsConsult = Nothing
    Set rsConsult2 = Nothing
    Exit Sub

ME_guardarInd_D:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo ME_Fin_D

End Sub

Private Sub guardarInd_E(ByVal estrnro As Long, ByVal estrdabr As String, ByVal ternro As Long, ByVal Anio As Integer, ByVal Fecha As Date, ByVal evenro As Integer, ByVal evecodext As String, ByVal evedesabr As String)
    Dim rsConsult As New ADODB.Recordset
    Dim rsConsult2 As New ADODB.Recordset
    Dim rsConsult3 As New ADODB.Recordset
    'Dim sumar_emp As Integer
    Dim cuatr As Integer
    
    On Error GoTo ME_guardarInd_E
    
    If Month(Fecha) / 4 <= 1 Then
        cuatr = 1
    Else
        If Month(Fecha) / 4 > 1 And Month(Fecha) / 4 <= 2 Then
            cuatr = 2
        Else
            cuatr = 3
        End If
    End If

    If evenro = 75 Then
        Debug.Print "" & ternro
    End If
    StrSql = " SELECT * FROM rep_ind_E WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND anio = " & Anio
    StrSql = StrSql & " AND cuatrimestre = " & cuatr
    StrSql = StrSql & " AND estrnro = " & estrnro
    StrSql = StrSql & " AND evenro = " & evenro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        StrSql = "INSERT INTO rep_ind_E (bpronro , Anio, cuatrimestre, estrnro, estrdabr, evenro, resultmin, resultmed, resultmax)"
        StrSql = StrSql & " VALUES (" & NroProceso
        StrSql = StrSql & ", " & Anio
        StrSql = StrSql & ", " & cuatr
        StrSql = StrSql & ", " & estrnro
        StrSql = StrSql & ", '" & estrdabr & "'"
        StrSql = StrSql & ", " & evenro
        StrSql = StrSql & ", 0" ' minimo
        StrSql = StrSql & ", 0" ' promedio
        StrSql = StrSql & ", 0" ' maximo
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rsConsult.Close
    
    StrSql = " SELECT  e.ternro, e.empleg, e.ternom, e.terape, e.ternom2, e.terape2, min(resval) resultmin, max(resval) resultmax, avg(resval) resultmed "
    StrSql = StrSql & " From cap_evaluacion "
    StrSql = StrSql & " INNER JOIN v_empleado e ON e.ternro = cap_evaluacion.evaparticipante "
    StrSql = StrSql & " INNER JOIN cap_resultado ON cap_resultado.resnro = cap_evaluacion.evaresnro "
    StrSql = StrSql & " Where cap_evaluacion.evatipo = 4 "
    StrSql = StrSql & " AND cap_evaluacion.evenro =  " & evenro
    StrSql = StrSql & " AND ternro = " & ternro
    StrSql = StrSql & " GROUP BY e.ternro, e.empleg, e.ternom, e.terape, e.ternom2, e.terape2 "
    OpenRecordset StrSql, rsConsult2
    If Not rsConsult2.EOF Then
        Do While Not rsConsult2.EOF
            StrSql = "INSERT INTO rep_ind_E_det ( bpronro , Anio, cuatrimestre, estrnro, estrdabr, evenro, evedesabr, ternro, empleg, ternom, ternom2, terape, terape2, resultmin, resultmed, resultmax)"
            StrSql = StrSql & " VALUES (" & NroProceso
            StrSql = StrSql & ", " & Anio
            StrSql = StrSql & ", " & cuatr
            StrSql = StrSql & ", " & estrnro
            StrSql = StrSql & ", '" & estrdabr & "'"
            StrSql = StrSql & ", " & evenro
            StrSql = StrSql & ", '" & evedesabr & "'"
            StrSql = StrSql & ", " & ternro
            StrSql = StrSql & ", " & rsConsult2("empleg")
            StrSql = StrSql & ", '" & rsConsult2("ternom") & "'"
            StrSql = StrSql & ", '" & rsConsult2("ternom2") & "'"
            StrSql = StrSql & ", '" & rsConsult2("terape") & "'"
            StrSql = StrSql & ", '" & rsConsult2("terape2") & "'"
            StrSql = StrSql & ", " & rsConsult2("resultmin")
            StrSql = StrSql & ", " & rsConsult2("resultmed")
            StrSql = StrSql & ", " & rsConsult2("resultmax")
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            rsConsult2.MoveNext
        Loop
        
        'Actualizo los totales
        StrSql = " SELECT min(resultmin) min, AVG(resultmed) med ,MAX(resultmax) max "
        StrSql = StrSql & " FROM rep_ind_E_det "
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND cuatrimestre = " & cuatr
        StrSql = StrSql & " AND estrnro = " & estrnro
        StrSql = StrSql & " AND evenro = " & evenro
        OpenRecordset StrSql, rsConsult2
        If Not rsConsult2.EOF Then
            StrSql = "UPDATE rep_ind_E SET "
            StrSql = StrSql & " resultmin = " & CDbl(rsConsult2!Min)
            StrSql = StrSql & " ,resultmed = " & CDbl(rsConsult2!Med)
            StrSql = StrSql & " ,resultmax = " & CDbl(rsConsult2!Max)
            StrSql = StrSql & " WHERE bpronro = " & NroProceso
            StrSql = StrSql & " AND cuatrimestre = " & cuatr
            StrSql = StrSql & " AND estrnro = " & estrnro
            StrSql = StrSql & " AND evenro = " & evenro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        rsConsult2.Close
    End If
    
    
    
ME_Fin_E:
    Set rsConsult = Nothing
    Set rsConsult2 = Nothing
    Exit Sub

ME_guardarInd_E:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo ME_Fin_E


End Sub

Private Sub guardarInd_F(ByVal estrnro As Long, ByVal estrdabr As String, ByVal ternro As Long, ByVal empleg As Long, ByVal terape As String, ByVal terape2 As String, ByVal ternom As String, ByVal ternom2 As String, ByVal Anio As Integer, ByVal Fecha As Date, ByVal control_asist As Integer, ByVal presente As Integer, ByVal evenro As Integer, ByVal evecodext As String, ByVal evedesabr As String)
Dim rsConsult As New ADODB.Recordset
Dim rsConsult2 As New ADODB.Recordset
Dim sumar_emp As Integer
    
    On Error GoTo ME_guardarInd_F
    
    
    StrSql = " SELECT * FROM rep_ind_F WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND anio = " & Anio
    StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
    StrSql = StrSql & " AND estrnro = " & estrnro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        StrSql = "INSERT INTO rep_ind_F (bpronro, anio, fecha, estrnro, estrdabr, cantasist, cantemp)"
        StrSql = StrSql & " VALUES (" & NroProceso
        StrSql = StrSql & ", " & Anio
        StrSql = StrSql & ", " & ConvFecha(Fecha)
        StrSql = StrSql & ", " & estrnro
        StrSql = StrSql & ", '" & estrdabr & "'"
        If presente = -1 Then
            StrSql = StrSql & ", 1"
        Else
            StrSql = StrSql & ", 0"
        End If
        StrSql = StrSql & ", 1"
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
    Else
        StrSql = "UPDATE rep_ind_F SET cantemp = " & rsConsult!cantEmp + 1
        If presente = -1 Then
            StrSql = StrSql & ", cantasist = " & rsConsult!cantasist + 1
        End If
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
        StrSql = StrSql & " AND estrnro = " & estrnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
    rsConsult.Close
    
    StrSql = "INSERT INTO rep_ind_F_det (bpronro, anio, fecha, estrnro, evenro, evedesabr, evereqasi, empleg, ternro, terape, terape2, ternom, ternom2, alcanzo)"
    StrSql = StrSql & " VALUES (" & NroProceso
    StrSql = StrSql & ", " & Anio
    StrSql = StrSql & ", " & ConvFecha(Fecha)
    StrSql = StrSql & ", " & estrnro
    StrSql = StrSql & ", " & evenro
    StrSql = StrSql & ", '" & evedesabr & "'"
    StrSql = StrSql & ", " & control_asist
    StrSql = StrSql & ", " & empleg
    StrSql = StrSql & ", " & ternro
    StrSql = StrSql & ", '" & terape & "'"
    StrSql = StrSql & ", '" & terape2 & "'"
    StrSql = StrSql & ", '" & ternom & "'"
    StrSql = StrSql & ", '" & ternom2 & "'"
    StrSql = StrSql & ", " & presente
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
ME_Fin_F:
    Set rsConsult = Nothing
    Set rsConsult2 = Nothing
    Exit Sub

ME_guardarInd_F:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo ME_Fin_F
End Sub

Private Sub guardarInd_G(ByVal estrnro As Long, ByVal estrdabr As String, ByVal evenro As Long, ByVal ternro As Long, ByVal empleg As Long, ByVal ternom As String, ByVal terape As String, ByVal ternom2 As String, ByVal terape2 As String, ByVal Anio As Integer, ByVal lista_itenro_conf As String, ByVal fecdesdeA As Date)

    Dim rsConsult As New ADODB.Recordset
    Dim rsConsult2 As New ADODB.Recordset
    Dim CostoInd As Double
    Dim evedesabr As String
    Dim cantEmp As Long
        
    On Error GoTo ME_guardarInd_G
    ' 0 = Evento
    ' 1 = costos
    ' 2 = formadores
    ' 3 = lugares
    ' 4 = materiales
    ' 5 = recursos
    ' 6 = Empleados
    
    '---- Primero inserto el reporte -----
    StrSql = " SELECT * FROM rep_ind_G WHERE bpronro = " & NroProceso
    StrSql = StrSql & " AND anio = " & Anio
    StrSql = StrSql & " AND estrnro = " & estrnro
    'StrSql = StrSql & " AND evenro = " & evenro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
    '---- Busco la cantidad de empleados para la estructura ----
        StrSql = " SELECT COUNT(empleado.ternro) cantemp "
        StrSql = StrSql & " FROM Empleado"
        StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = empleado.ternro "
        StrSql = StrSql & " WHERE estrnro = " & estrnro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(fecdesdeA)
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            cantEmp = 0
        Else
            cantEmp = rsConsult2!cantEmp
        End If
        
        'Creo el registro en la tabla
        StrSql = "INSERT INTO rep_ind_G (bpronro, anio, estrnro, estrdabr, costot, cantemp )"
        StrSql = StrSql & " VALUES (" & NroProceso
        StrSql = StrSql & ", " & Anio
        StrSql = StrSql & ", " & estrnro
        StrSql = StrSql & ", '" & estrdabr & "'"
        StrSql = StrSql & ", " & 0
        StrSql = StrSql & ", " & cantEmp
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        'Flog.writeline StrSql
    End If
    rsConsult.Close
    
    
    StrSql = " SELECT * FROM rep_ind_G_det "
    StrSql = StrSql & " WHERE rep_ind_G_det.bpronro = " & NroProceso
    StrSql = StrSql & " AND rep_ind_G_det.anio = " & Anio
    StrSql = StrSql & " AND estrnro = " & estrnro
    StrSql = StrSql & " AND evenro = " & evenro
    OpenRecordset StrSql, rsConsult
    If rsConsult.EOF Then
        
        '---------- Traigo los datos de la tabla de Eventos ----------'
        StrSql = "SELECT  evecostogral, * "
        StrSql = StrSql & " FROM cap_evento"
        StrSql = StrSql & " Where evenro =  " & evenro
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & " No se encontraron datos en la tabla EVENTOS"
        Else
            CostoInd = CDbl(rsConsult2("evecostoind"))
            evedesabr = CStr(rsConsult2!evedesabr)
            
            'Inserto los detalles del evento
            Do Until rsConsult2.EOF
                StrSql = "INSERT INTO rep_ind_G_det ("
                StrSql = StrSql & " bpronro , Anio, estrnro, estrdabr, evenro, evedesabr, costodesc, costo, empleg, ternom, terape, ternom2, terape2, tipdet "
                StrSql = StrSql & ") VALUES ("
                StrSql = StrSql & NroProceso & "," 'bpronro
                StrSql = StrSql & Anio & "," 'Anio
                StrSql = StrSql & estrnro & "," 'estrnro
                StrSql = StrSql & "'" & estrdabr & "'," 'estrdabr
                StrSql = StrSql & evenro & "," 'evenro
                StrSql = StrSql & "'" & evedesabr & "'," 'evedesabr
                StrSql = StrSql & "'" & evedesabr & "'," 'costrdesc
                StrSql = StrSql & rsConsult2!evecostogral & "," 'costo
                StrSql = StrSql & "null," 'empleg
                StrSql = StrSql & "null," 'ternom
                StrSql = StrSql & "null," 'terape
                StrSql = StrSql & "null," 'ternom2
                StrSql = StrSql & "null," 'terape2
                StrSql = StrSql & "0" 'tipdet
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                'Flog.writeline StrSql
                If Not IsNull(rsConsult2("evecostogral")) Then
                    StrSql = "UPDATE rep_ind_G set costot = costot + " & CDbl(rsConsult2!evecostogral)
                    StrSql = StrSql & " WHERE bpronro = " & NroProceso
                    StrSql = StrSql & " AND estrnro = " & estrnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'Flog.writeline StrSql
                End If
                
                rsConsult2.MoveNext
            Loop
        End If
        rsConsult2.Close
                
        
        '---------- Costos ----------'
        StrSql = "SELECT cosnro, costo.itenro, cosmonto, itedesabr "
        StrSql = StrSql & " FROM costo"
        StrSql = StrSql & " INNER JOIN gco_item ON gco_item.itenro = costo.itenro "
        StrSql = StrSql & " WHERE evenro =  " & evenro
        StrSql = StrSql & " AND gco_item.itenro NOT IN (" & lista_itenro_conf & ")"
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & " No se encontraron costos "
        Else
            Do Until rsConsult2.EOF
                'DescCostos = CStr(rs3("itedesabr"))
                'CostoCosto = CostoCosto + CDbl(rs3("cosmonto"))
                
                StrSql = "INSERT INTO rep_ind_G_det ("
                StrSql = StrSql & " bpronro , Anio, estrnro, estrdabr, evenro, evedesabr, costodesc, costo, ternro, empleg, ternom, terape, ternom2, terape2, tipdet "
                StrSql = StrSql & ") VALUES ("
                StrSql = StrSql & NroProceso & "," 'bpronro
                StrSql = StrSql & Anio & "," 'Anio
                StrSql = StrSql & estrnro & "," 'estrnro
                StrSql = StrSql & "'" & estrdabr & "'," 'estrdabr
                StrSql = StrSql & evenro & "," 'evenro
                StrSql = StrSql & "'" & evedesabr & "'," 'evedesabr
                StrSql = StrSql & "'" & CStr(rsConsult2!itedesabr) & "'," 'costodesc
                StrSql = StrSql & CDbl(rsConsult2!cosmonto) & "," 'costo
                StrSql = StrSql & "null," 'empleg
                StrSql = StrSql & "null," 'ternom
                StrSql = StrSql & "null," 'terape
                StrSql = StrSql & "null," 'ternom2
                StrSql = StrSql & "null," 'terape2
                StrSql = StrSql & "1" 'tipdet
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                'Flog.writeline StrSql
                If Not IsNull(rsConsult2("cosmonto")) Then
                    StrSql = "UPDATE rep_ind_G set costot = costot + " & CDbl(rsConsult2!cosmonto)
                    StrSql = StrSql & " WHERE bpronro = " & NroProceso
                    StrSql = StrSql & " AND estrnro = " & estrnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'Flog.writeline StrSql
                End If
                rsConsult2.MoveNext
            Loop
        End If
        rsConsult2.Close
        
            
        '---------- Formadores ----------'
        StrSql = "SELECT  monto, cap_eventoformador.ternro, ternom, ternom2, terape, terape2"
        StrSql = StrSql & " FROM cap_eventoformador"
        StrSql = StrSql & " INNER JOIN tercero ON tercero.ternro = cap_eventoformador.ternro"
        StrSql = StrSql & " WHERE evenro =  " & evenro
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & " No se encontraron costos en Formadores "
        Else
            Do Until rsConsult2.EOF
                StrSql = "INSERT INTO rep_ind_G_det ("
                StrSql = StrSql & " bpronro , Anio, estrnro, estrdabr, evenro, evedesabr, costodesc, costo, ternro, empleg, ternom, terape, ternom2, terape2, tipdet "
                StrSql = StrSql & ") VALUES ("
                StrSql = StrSql & NroProceso & "," 'bpronro
                StrSql = StrSql & Anio & "," 'Anio
                StrSql = StrSql & estrnro & "," 'estrnro
                StrSql = StrSql & "'" & estrdabr & "'," 'estrdabr
                StrSql = StrSql & evenro & "," 'evenro
                StrSql = StrSql & "'" & evedesabr & "'," 'evedesabr
                StrSql = StrSql & "'Formador'," 'costodesc
                StrSql = StrSql & CDbl(rsConsult2!Monto) & "," 'costo
                StrSql = StrSql & CStr(rsConsult2!ternro) & " ,"  'ternro
                StrSql = StrSql & "null," 'empleg
                StrSql = StrSql & "'" & CStr(rsConsult2!ternom) & "'," 'ternom
                StrSql = StrSql & "'" & CStr(rsConsult2!terape) & "'," 'terape
                StrSql = StrSql & "'" & CStr(rsConsult2!ternom2) & "'," 'ternom2
                StrSql = StrSql & "'" & CStr(rsConsult2!terape2) & "'," 'terape2
                StrSql = StrSql & "2" 'tipdet
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                'Flog.writeline StrSql
                If Not IsNull(rsConsult2("monto")) Then
                    StrSql = "UPDATE rep_ind_G set costot = costot + " & CDbl(rsConsult2!Monto)
                    StrSql = StrSql & " WHERE bpronro = " & NroProceso
                    StrSql = StrSql & " AND estrnro = " & estrnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'Flog.writeline StrSql
                End If
                rsConsult2.MoveNext
            Loop
        End If
        rsConsult2.Close
            
        '---------- Lugares ----------'
        StrSql = " SELECT monto, itedesabr, lugdesabr, cap_lugar.itenro "
        StrSql = StrSql & " FROM cap_eventolugar "
        StrSql = StrSql & " INNER JOIN cap_lugar ON cap_lugar.lugnro = cap_eventolugar.lugnro "
        StrSql = StrSql & " INNER JOIN gco_item ON gco_item.itenro = cap_lugar.itenro "
        StrSql = StrSql & " WHERE cap_eventolugar.evenro = " & evenro
        StrSql = StrSql & " AND gco_item.itenro NOT IN (" & lista_itenro_conf & ")"
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & " No se encontraron costos en Lugares"
        Else
            Do Until rsConsult2.EOF
                StrSql = "INSERT INTO rep_ind_G_det ("
                StrSql = StrSql & " bpronro , Anio, estrnro, estrdabr, evenro, evedesabr, costodesc, costo, ternro, empleg, ternom, terape, ternom2, terape2, tipdet "
                StrSql = StrSql & ") VALUES ("
                StrSql = StrSql & NroProceso & "," 'bpronro
                StrSql = StrSql & Anio & "," 'Anio
                StrSql = StrSql & estrnro & "," 'estrnro
                StrSql = StrSql & "'" & estrdabr & "'," 'estrdabr
                StrSql = StrSql & evenro & "," 'evenro
                StrSql = StrSql & "'" & evedesabr & "'," 'evedesabr
                StrSql = StrSql & "'" & CStr(rsConsult2!itedesabr) & "'," 'costodesc
                StrSql = StrSql & CDbl(rsConsult2!Monto) & "," 'costo
                StrSql = StrSql & " Null,"  'ternro
                StrSql = StrSql & " Null," 'empleg
                StrSql = StrSql & " Null," 'ternom
                StrSql = StrSql & " Null," 'terape
                StrSql = StrSql & " Null," 'ternom2
                StrSql = StrSql & " Null," 'terape2
                StrSql = StrSql & "3" 'tipdet
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                'Flog.writeline StrSql
                If Not IsNull(rsConsult2("monto")) Then
                    StrSql = "UPDATE rep_ind_G set costot = costot + " & CDbl(rsConsult2!Monto)
                    StrSql = StrSql & " WHERE bpronro = " & NroProceso
                    StrSql = StrSql & " AND estrnro = " & estrnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'Flog.writeline StrSql
                End If
                rsConsult2.MoveNext
            Loop
        End If
   
                
        '---------- Materiales ----------'
        StrSql = " SELECT monto, itedesabr, matdesabr, cap_material.itenro "
        StrSql = StrSql & " FROM cap_eventomaterial "
        StrSql = StrSql & " INNER JOIN cap_material ON cap_material.matnro = cap_eventomaterial.matnro "
        StrSql = StrSql & " INNER JOIN gco_item ON gco_item.itenro = cap_material.itenro "
        StrSql = StrSql & " WHERE cap_eventomaterial.evenro = " & evenro
        StrSql = StrSql & " AND gco_item.itenro NOT IN (" & lista_itenro_conf & ")"
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & " No se encontraron costos en Materiales "
        Else
            Do Until rsConsult2.EOF
                StrSql = "INSERT INTO rep_ind_G_det ("
                StrSql = StrSql & " bpronro , Anio, estrnro, estrdabr, evenro, evedesabr, costodesc, costo, ternro, empleg, ternom, terape, ternom2, terape2, tipdet "
                StrSql = StrSql & ") VALUES ("
                StrSql = StrSql & NroProceso & "," 'bpronro
                StrSql = StrSql & Anio & "," 'Anio
                StrSql = StrSql & estrnro & "," 'estrnro
                StrSql = StrSql & "'" & estrdabr & "'," 'estrdabr
                StrSql = StrSql & evenro & "," 'evenro
                StrSql = StrSql & "'" & evedesabr & "'," 'evedesabr
                StrSql = StrSql & "'" & CStr(rsConsult2!matdesabr) & "'," 'costodesc
                StrSql = StrSql & CDbl(rsConsult2!Monto) & "," 'costo
                StrSql = StrSql & " Null,"  'ternro
                StrSql = StrSql & " Null," 'empleg
                StrSql = StrSql & " Null," 'ternom
                StrSql = StrSql & " Null," 'terape
                StrSql = StrSql & " Null," 'ternom2
                StrSql = StrSql & " Null," 'terape2
                StrSql = StrSql & "4" 'tipdet
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                'Flog.writeline StrSql
                If Not IsNull(rsConsult2("monto")) Then
                    StrSql = "UPDATE rep_ind_G set costot = costot + " & CDbl(rsConsult2!Monto)
                    StrSql = StrSql & " WHERE bpronro = " & NroProceso
                    StrSql = StrSql & " AND estrnro = " & estrnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'Flog.writeline StrSql
                End If
                rsConsult2.MoveNext
            Loop
        End If
        rsConsult2.Close
            
        '---------- Recursos ----------'
        StrSql = " SELECT monto, itedesabr, recdesabr, cap_recurso.itenro "
        StrSql = StrSql & " FROM cap_eventorecurso "
        StrSql = StrSql & " INNER JOIN cap_recurso ON cap_recurso.recnro = cap_eventorecurso.recnro "
        StrSql = StrSql & " INNER JOIN gco_item ON gco_item.itenro = cap_recurso.itenro "
        StrSql = StrSql & " WHERE cap_eventorecurso.evenro = " & evenro
        StrSql = StrSql & " AND gco_item.itenro NOT IN (" & lista_itenro_conf & ")"
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & " No se encontraron costos en Recursos "
        Else
            Do Until rsConsult2.EOF
                StrSql = "INSERT INTO rep_ind_G_det ("
                StrSql = StrSql & " bpronro , Anio, estrnro, estrdabr, evenro, evedesabr, costodesc, costo, ternro, empleg, ternom, terape, ternom2, terape2, tipdet "
                StrSql = StrSql & ") VALUES ("
                StrSql = StrSql & NroProceso & "," 'bpronro
                StrSql = StrSql & Anio & "," 'Anio
                StrSql = StrSql & estrnro & "," 'estrnro
                StrSql = StrSql & "'" & estrdabr & "'," 'estrdabr
                StrSql = StrSql & evenro & "," 'evenro
                StrSql = StrSql & "'" & evedesabr & "'," 'evedesabr
                StrSql = StrSql & "'" & CStr(rsConsult2!recdesabr) & "'," 'costodesc
                StrSql = StrSql & CDbl(rsConsult2!Monto) & "," 'costo
                StrSql = StrSql & " Null,"  'ternro
                StrSql = StrSql & " Null," 'empleg
                StrSql = StrSql & " Null," 'ternom
                StrSql = StrSql & " Null," 'terape
                StrSql = StrSql & " Null," 'ternom2
                StrSql = StrSql & " Null," 'terape2
                StrSql = StrSql & "5" 'tipdet
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                'Flog.writeline StrSql
                If Not IsNull(rsConsult2("monto")) Then
                    StrSql = "UPDATE rep_ind_G set costot = costot + " & CDbl(rsConsult2!Monto)
                    StrSql = StrSql & " WHERE bpronro = " & NroProceso
                    StrSql = StrSql & " AND estrnro = " & estrnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'Flog.writeline StrSql
                End If
                rsConsult2.MoveNext
            Loop
        End If
        rsConsult2.Close
            
        '------- Empleado -------
        StrSql = "INSERT INTO rep_ind_G_det ("
        StrSql = StrSql & " bpronro , Anio, estrnro, estrdabr, evenro, evedesabr, costodesc, costo, ternro, empleg, ternom, terape, ternom2, terape2, tipdet "
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & NroProceso & "," 'bpronro
        StrSql = StrSql & Anio & "," 'Anio
        StrSql = StrSql & estrnro & "," 'estrnro
        StrSql = StrSql & "'" & estrdabr & "'," 'estrdabr
        StrSql = StrSql & evenro & "," 'evenro
        StrSql = StrSql & "'" & evedesabr & "'," 'evedesabr
        StrSql = StrSql & "'Empleado'," 'costodesc
        StrSql = StrSql & CDbl(CostoInd) & "," 'costo
        StrSql = StrSql & "'" & ternro & "'," 'ternro
        StrSql = StrSql & "'" & empleg & "'," 'empleg
        StrSql = StrSql & "'" & ternom & "'," 'ternom
        StrSql = StrSql & "'" & terape & "'," 'terape
        StrSql = StrSql & "'" & ternom2 & "'," 'ternom2
        StrSql = StrSql & "'" & terape2 & "'," 'terape2
        StrSql = StrSql & "6" 'tipdet
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        'Flog.writeline StrSql
        StrSql = "UPDATE rep_ind_G set costot = costot + " & CDbl(CostoInd)
        StrSql = StrSql & " WHERE bpronro = " & NroProceso
        StrSql = StrSql & " AND estrnro = " & estrnro
        objConn.Execute StrSql, , adExecuteNoRecords
        'Flog.writeline StrSql
    Else

        StrSql = "SELECT  evecostogral, * "
        StrSql = StrSql & " FROM cap_evento"
        StrSql = StrSql & " Where evenro =  " & evenro
        OpenRecordset StrSql, rsConsult2
        If rsConsult2.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & " No se encontraron datos en la tabla EVENTOS"
        Else
            CostoInd = CDbl(rsConsult2("evecostoind"))
            evedesabr = CStr(rsConsult2!evedesabr)
            
            '------- Empleado -------
            StrSql = "INSERT INTO rep_ind_G_det ("
            StrSql = StrSql & " bpronro , Anio, estrnro, estrdabr, evenro, evedesabr, costodesc, costo, ternro, empleg, ternom, terape, ternom2, terape2, tipdet "
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & NroProceso & "," 'bpronro
            StrSql = StrSql & Anio & "," 'Anio
            StrSql = StrSql & estrnro & "," 'estrnro
            StrSql = StrSql & "'" & estrdabr & "'," 'estrdabr
            StrSql = StrSql & evenro & "," 'evenro
            StrSql = StrSql & "'" & evedesabr & "'," 'evedesabr
            StrSql = StrSql & "'Empleado'," 'costodesc
            StrSql = StrSql & CDbl(CostoInd) & "," 'costo
            StrSql = StrSql & "'" & ternro & "'," 'ternro
            StrSql = StrSql & "'" & empleg & "'," 'empleg
            StrSql = StrSql & "'" & ternom & "'," 'ternom
            StrSql = StrSql & "'" & terape & "'," 'terape
            StrSql = StrSql & "'" & ternom2 & "'," 'ternom2
            StrSql = StrSql & "'" & terape2 & "'," 'terape2
            StrSql = StrSql & "6" 'tipdet
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            'Flog.writeline StrSql
            StrSql = "UPDATE rep_ind_G set costot = costot + " & CDbl(CostoInd)
            StrSql = StrSql & " WHERE bpronro = " & NroProceso
            StrSql = StrSql & " AND estrnro = " & estrnro
            objConn.Execute StrSql, , adExecuteNoRecords
            'Flog.writeline StrSql
        End If
        
    End If
    rsConsult.Close

ME_Fin_G:
    Set rsConsult = Nothing
    Set rsConsult2 = Nothing
    Exit Sub

ME_guardarInd_G:
    HuboErrores = True
    Flog.writeline
'    Resume Next
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "Ultimo SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
    Flog.writeline
    GoTo ME_Fin_G
End Sub

