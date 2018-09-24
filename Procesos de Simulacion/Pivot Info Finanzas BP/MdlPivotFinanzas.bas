Attribute VB_Name = "MdlPivotFinanzas"
Option Explicit
'Const Version = "1.00"
'Const FechaVersion = "29/09/2011"
'Const Autor = "Gonzalez Nicolás"
'Const Modificacion = "Version inicial"

'Const Version = "1.01"
'Const FechaVersion = "15/12/2011"
'Const Autor = "Gonzalez Nicolás"
'Const Modificacion = "Utiliza los datos del asiento contable."

'Const Version = "1.02"
'Const FechaVersion = "12/04/2012"
'Const Autor = "Deluchi Ezequiel"
'Const Modificacion = "Se mejoro los tiempos de la consulta de los asientos."

'Const Version = "1.03"
'Const FechaVersion = "15/05/2012"
'Const Autor = "Lisandro Moro"
'Const Modificacion = "Se agrego el filtro por modelos de asientos configurados como presupuesto conf_cont.cofcod = '1'." & vbCrLf & _
'                     " Se cambio el orden de los meses para que coincida con el reporte."

'Const Version = "1.04"
'Const FechaVersion = "15/06/2012"
'Const Autor = "Lisandro Moro"
'Const Modificacion = "Se agrego AND detalle_asi.ternro = sim_cabliq.empleado."

'Const Version = "1.05"
'Const FechaVersion = "28/09/2012"
'Const Autor = "Sebastian Stremel"
'Const Modificacion = "Se guarda dlmonto si es un concepto"

'Const Version = "1.06"
'Const FechaVersion = "04/12/2012"
'Const Autor = "Lisandro Moro"
'Const Modificacion = " Se agregaron condiciones por concepto o acumulador en sim_rep_pivotfinanzas " & vbCrLf & _
'                     " Se valida que solo sean procesos volcados."

'Const Version = "1.07"
'Const FechaVersion = "17/07/2015"
'Const Autor = "Miriam Ruiz - CAS-32019 - H&A - Errores GP"
'Const Modificacion = " Se modificó el progreso para que se pusiera en 100% cuando dá error "
           
Const Version = "1.08"
Const FechaVersion = "29/02/2016"
Const Autor = "Borrelli Facundo - CAS-33105 - H&A - Testeo R4 Patch Agosto 2015 - Cono Sur - GP - [Entrega 2]"
Const Modificacion = " Se modifico la consulta que busca los periodos de simulacion liquidados "
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------

Global Descripcion As String
Global Cantidad As Single
Dim I As Long
Global arregloTablas() As String
Global ListaProcesos As String
Global ListaCabeceras As String
Global ListaTerceros As String
Global pliqnroDesde As Long
Global pliqnroHasta As Long
Global AnioDesde As Long
Global AnioHasta As Long
Global MesDesde As Long
Global MesHasta As Long
Global fechadesde As String
Global fechahasta As String
Global CantPasos As Double
Global IncPasos As Double
Global Progreso As Double


'Global pliqdesde As String
'Global pliqhasta As String


Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial
' Autor      : Gonzalez Nicolás
' Fecha      : 29/09/2011
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
    
    Nombre_Arch = PathFLog & "ReportePivot" & "-" & NroProcesoBatch & ".log"
    
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
    CantPasos = 36 ' cantidad de tablas + 1
    IncPasos = 100 / CantPasos
    Progreso = 0
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 309 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Pivot(NroProcesoBatch, bprcparam)
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
        actualizarProgreso 100
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


Public Sub InsertarReg(ByVal bpronro, ByVal Ternro, ByVal Estrdabr, ByVal pliqnro, ByVal fdesde, ByVal fhasta, ByVal concnro, ByVal acuNro, ByVal num_mes, ByVal Valor, ByVal cuenta)
    'Inserta registros segun n° de mes
    Dim a
    Dim pivmes(11)
    
    
    For a = 0 To 11
        If a = num_mes - 1 Then
            pivmes(a) = Valor
        Else
            pivmes(a) = "Null"
        End If
    Next

    StrSql = "INSERT INTO sim_rep_pivotfinanzas"
    StrSql = StrSql & " ("
    StrSql = StrSql & " bpronro,ternro,estrdabr,pliqnro,pivfdesde,pivfhasta,concnro,acunro"
    StrSql = StrSql & ",pivmes1,pivmes2,pivmes3,pivmes4,pivmes5,pivmes6,pivmes7,pivmes8,pivmes9,pivmes10,pivmes11,pivmes12"
    StrSql = StrSql & ",cuenta"
    StrSql = StrSql & ")"
    
    StrSql = StrSql & " VALUES ("
    
    StrSql = StrSql & bpronro & "," & Ternro & ",'" & Estrdabr & "'"
    StrSql = StrSql & "," & pliqnro & "," & ConvFecha(fdesde) & "," & ConvFecha(fhasta)
    StrSql = StrSql & "," & concnro
    StrSql = StrSql & "," & acuNro
    StrSql = StrSql & "," & pivmes(0) & "," & pivmes(1) & "," & pivmes(2)
    StrSql = StrSql & "," & pivmes(3) & "," & pivmes(4) & "," & pivmes(5)
    StrSql = StrSql & "," & pivmes(6) & "," & pivmes(7)
    StrSql = StrSql & "," & pivmes(8) & "," & pivmes(9)
    StrSql = StrSql & "," & pivmes(10) & "," & pivmes(11)
    StrSql = StrSql & ",'" & cuenta & "'"
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    'Flog.writeline StrSql
    Flog.writeline Space(5) & "Inserto Mes: " & num_mes & " - Valor: " & Valor
    
End Sub


Public Sub ActualizarReg(ByVal id, ByVal num_mes, ByVal Valor)
    Dim rs As New ADODB.Recordset
    StrSql = " SELECT pivmes" & num_mes & " mes "
    StrSql = StrSql & " FROM sim_rep_pivotfinanzas "
    StrSql = StrSql & " WHERE reppivnro =" & id
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        If IsNull(rs("mes")) Then
            'UPDATE registros x ID
            StrSql = " UPDATE sim_rep_pivotfinanzas SET "
            StrSql = StrSql & " pivmes" & num_mes & " = " & Valor
            StrSql = StrSql & " WHERE reppivnro =" & id
        Else
            'UPDATE registros x ID (SUMO)
            StrSql = " UPDATE sim_rep_pivotfinanzas SET "
            StrSql = StrSql & " pivmes" & num_mes & "= pivmes" & num_mes & " + " & Valor
            StrSql = StrSql & " WHERE reppivnro =" & id
        End If
        'Flog.writeline StrSql
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        Flog.writeline Space(5) & "ERROR: esto no debe ocurrir."
    End If
    rs.Close
    
    ''UPDATE registros x ID
    'StrSql = " UPDATE sim_rep_pivotfinanzas SET "
    'StrSql = StrSql & " pivmes" & num_mes & "=" & Valor
    'StrSql = StrSql & " WHERE reppivnro =" & id
    ''Flog.writeline StrSql
    'objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline Space(5) & "Actualizo Mes: " & num_mes & " - Valor: " & Valor
    
End Sub

Public Sub Pivot(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : Gonzalez Nicolás
' Fecha      : 26/09/2011
' Modificacion:
' --------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset

Dim cantRegistros
Dim CantidadEmpleados As Long
Dim Sep As String
Dim simhisnro As Long
Dim tipo As String

Dim ArrParam
Dim ArrCabliq

'On Error GoTo ME_Main

TiempoAcumulado = GetTickCount

'Parametros de Entrada
Dim FechaInicio
Dim FechaFin
Dim legdesde
Dim leghasta
Dim filtro
Dim Orden

'Parametros de Confrep
Dim Ac
Dim Co
Dim tenro
'-------------------
Dim Ternro
Dim cuentaMes
'Dim CentroCosto
Dim concnro
Dim acuNro
Dim Valor
'Dim estrnro
Dim Ult_CoAc
Dim Estrdabr
cuentaMes = 0
'CentroCosto = ""
concnro = ""
acuNro = ""

Dim cuenta
cuenta = ""

'Actualizo progreso al 1%
actualizarProgreso (1)

'----------------------------------------------------------------------------
' Levanto cada parametro por separado
'----------------------------------------------------------------------------
Flog.writeline "Levantando parametros. " & Parametros
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        'Guardo parametros de entrada
         ArrParam = Split(Parametros, "@")
         
         FechaInicio = ArrParam(0)
         FechaFin = ArrParam(1)
         legdesde = ArrParam(2)
         leghasta = ArrParam(3)
         filtro = ArrParam(4)
         Orden = ArrParam(5)
    End If
Else
    Flog.writeline "Error - Parametros nulos"
    HuboError = True
    Exit Sub
End If

Flog.writeline "Terminó de levantar los parametros"

'----------------------------------------------------------------------------
' Levanto Parametros de Confrep
'----------------------------------------------------------------------------
StrSql = "SELECT confrep.conftipo,confrep.confval,confrep.confval2 FROM confrep"
StrSql = StrSql & " WHERE repnro = 353"
OpenRecordset StrSql, rs
If Not rs.EOF Then
    'tenro = "0"
    Ac = "0"
    Co = "0"
    Do While Not rs.EOF
        Select Case rs!conftipo
            Case "ES":
                tenro = rs!confval
            Case "AC":
            
                If rs!confval2 <> "-1" Then
                    Ac = Ac & "," & rs!confval2
                Else
                    Ac = "-1"
                End If
                
            Case "CO":
                
                If rs!confval2 <> "-1" Then
                    Co = Co & ",'" & rs!confval2 & "'" 'licho
                Else
                    Co = -1
                End If
        
        End Select
        rs.MoveNext
    Loop
Else
    Flog.writeline "Error - Parametros de confrep nulos"
    HuboError = True
    Exit Sub
End If
rs.Close


'Actualizo progreso al 2%
actualizarProgreso (2)


'----------------------------------------------------------------------------
'TRAE LA LISTA DE PERIODOS A BUSCAR
'----------------------------------------------------------------------------
StrSql = "SELECT sim_proceso.pronro , sim_proceso.prodesc, periodo.pliqdesc,periodo.pliqnro "
StrSql = StrSql & " ,periodo.pliqdesde,periodo.pliqhasta "
StrSql = StrSql & " ,sim_cabliq.cliqnro,sim_cabliq.empleado "
StrSql = StrSql & " ,proc_vol.vol_cod "
StrSql = StrSql & " ,detalle_asi.cuenta , detalle_asi.Ternro, detalle_asi.dlmontoacum, detalle_asi.tipoorigen, detalle_asi.Origen, detalle_asi.dlmonto " 'sebastian stremel se agrego dlmonto
StrSql = StrSql & " ,pliqmes "

StrSql = StrSql & " FROM periodo "
StrSql = StrSql & " INNER JOIN sim_proceso ON sim_proceso.pliqnro = periodo.pliqnro "
StrSql = StrSql & " INNER JOIN sim_cabliq on sim_cabliq.pronro = sim_proceso.pronro "
StrSql = StrSql & " INNER JOIN sim_empleado ON sim_empleado.ternro = sim_cabliq.empleado "
StrSql = StrSql & " INNER JOIN proc_vol ON proc_vol.pliqnro = periodo.pliqnro "
'StrSql = StrSql & " INNER JOIN proc_vol_pl ON proc_vol.vol_cod = proc_vol_pl.vol_cod and proc_vol_pl.pronro = sim_proceso.pronro "
StrSql = StrSql & " INNER JOIN proc_vol_pl ON proc_vol.vol_cod = proc_vol_pl.vol_cod "
                    
StrSql = StrSql & " INNER JOIN detalle_asi ON detalle_asi.vol_cod = proc_vol.vol_cod AND detalle_asi.ternro =  sim_cabliq.empleado "
StrSql = StrSql & " INNER JOIN mod_asiento ON detalle_asi.masinro = mod_asiento.masinro "
StrSql = StrSql & " INNER JOIN conf_cont ON mod_asiento.cofcnro = conf_cont.cofcnro AND conf_cont.cofcod = '1' "

StrSql = StrSql & " WHERE pliqdesde >= " & ConvFecha(FechaInicio)
StrSql = StrSql & " AND pliqhasta <= " & ConvFecha(FechaFin)
StrSql = StrSql & " AND " & filtro
'StrSql = StrSql & " AND sim_empleado.empleg >= " & legdesde & " AND sim_empleado.empleg <= " & leghasta
StrSql = StrSql & " ORDER BY periodo.pliqmes,sim_empleado.empleg ASC"
OpenRecordset StrSql, rs
'Flog.writeline StrSql
Flog.writeline "Fecha Desde: " & FechaInicio & " - Hasta:" & FechaFin
Flog.writeline "Filtro : " & filtro
Flog.writeline "Strsql :" & StrSql
If rs.EOF Then
    Flog.writeline "Error - No existen períodos liquidados entre las fechas " & FechaInicio & " y " & FechaFin
    HuboError = True
    Exit Sub
Else
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = rs.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (97 / CEmpleadosAProc)

    
    Do While Not rs.EOF
        
        
        'Incrementa El progreso----------------------------------
        Progreso = Progreso + IncPorc
        '--------------------------------------------------------
        
        'Incrementa N° de mes para insertar en X columna----------
        cuentaMes = DateDiff("m", FechaInicio, rs!pliqhasta) + 1
        cuentaMes = rs!pliqmes
        '---------------------------------------------------------
        Ternro = rs!Empleado 'Licho - No actualizaba el ternro
        Flog.writeline "Periodo Mes: " & cuentaMes
        Flog.writeline "Empleado (ternro): " & Ternro
        '_________________________________________________________________
        'BUSCO el detalle del asiento
        'StrSql = "SELECT cuenta,ternro,dlmontoacum"
        'StrSql = StrSql & " ,tipoorigen , Origen "
        'StrSql = StrSql & " FROM detalle_asi where vol_cod = " & rs!vol_cod
        
        'StrSql = StrSql & " AND ternro = " & rs!Empleado
        '-------------------------------------------------------------------
        
'        StrSql = " SELECT "
'        StrSql = StrSql & " sim_detliq.concnro,sim_detliq.dlimonto"
'        StrSql = StrSql & " , 'CO' tipo"
'        StrSql = StrSql & " FROM sim_detliq"
'        StrSql = StrSql & " LEFT JOIN concepto ON concepto.concnro = sim_detliq.concnro "
'        StrSql = StrSql & " WHERE"
'        StrSql = StrSql & " sim_detliq.cliqnro = " & rs!cliqnro
'
'        If Co <> "-1" Then
'            StrSql = StrSql & " AND concepto.conccod IN (" & Co & ")"
'            'StrSql = StrSql & " AND  sim_detliq.concnro IN (" & Co & ")"
'        End If
'        '----------------------------
'        StrSql = StrSql & " UNION"
'        '----------------------------
'        StrSql = StrSql & " SELECT"
'        StrSql = StrSql & " sim_acu_liq.acunro,sim_acu_liq.almonto"
'        StrSql = StrSql & " ,'AC' tipo"
'        StrSql = StrSql & " FROM sim_acu_liq"
'        StrSql = StrSql & " WHERE"
'        StrSql = StrSql & " sim_acu_liq.cliqnro = " & rs!cliqnro
'
'        If Ac <> "-1" Then
'            StrSql = StrSql & " AND  sim_acu_liq.acunro IN (" & Ac & ")"
'        End If

        'OpenRecordset StrSql, rs2
        'If Not rs2.EOF Then
            'Do While Not rs2.EOF
               
                'Devuelve Estructura del empleado/NN
                Estrdabr = PerteneceEstruc(rs!Empleado, tenro, DateAdd("m", cuentaMes, PrimerDiaDeMes(Month(rs!pliqdesde), Year(rs!pliqdesde))), DateAdd("m", cuentaMes, UltimoDiaDeMes(Month(rs!pliqdesde), Year(rs!pliqdesde))))
                
                'Si la estructura es vacía, sale y ejecuta el prox. período
                If Not Trim(Estrdabr) = "" Then
                    
                    'Exit Do
                'End If
                
                Flog.writeline Space(5) & "Cuenta: " & rs!cuenta & " - Origen:" & rs!Origen & " - Valor:" & rs!dlmontoacum & " - Procesonro:" & rs!pronro
                
                'If rs2!tipo = "AC" Then
                'If rs2!tipoorigen = 2 And filtraAcum(rs2!Origen, Ac) = True Then
                If rs!tipoorigen = 2 And filtraAcum(rs!Origen, Ac) = True Then
                    acuNro = rs!Origen
                    Ult_CoAc = rs!cuenta
                    cuenta = rs!cuenta
                    concnro = "Null"
                    Valor = rs!dlmontoacum
                    'acuNro = rs2!concnro
                    'Ult_CoAc = rs2!concnro
                    'concnro = "Null"
                    'Valor = rs2!dlimonto
                    
                    'Si existe el AC, hago update
                    StrSql = "SELECT sim_rep_pivotfinanzas.reppivnro,sim_rep_pivotfinanzas.concnro,sim_rep_pivotfinanzas.acunro"
                    StrSql = StrSql & " FROM sim_rep_pivotfinanzas"
                    StrSql = StrSql & " WHERE cuenta = '" & cuenta & "' AND ternro = " & rs!Empleado & " And bpronro = " & bpronro & " AND estrdabr='" & Estrdabr & "'"
                    StrSql = StrSql & " AND acunro = " & acuNro
                    'StrSql = StrSql & " WHERE acunro = " & acuNro & " AND ternro = " & rs!Empleado & " And bpronro = " & bpronro & " AND estrdabr='" & Estrdabr & "'"
                    'StrSql = StrSql & " WHERE acunro = " & acunro & " AND ternro = " & rs!Empleado & " And bpronro = " & bpronro & " AND pliqnro = " & rs!pliqnro
                    'Ejecuto
                    OpenRecordset StrSql, rs3
                    If Not rs3.EOF Then
                        'Update registro
                        Call ActualizarReg(rs3!reppivnro, cuentaMes, Valor)
                    Else
                        'Inserto registro
                        Call InsertarReg(bpronro, rs!Empleado, Estrdabr, rs!pliqnro, FechaInicio, FechaFin, concnro, acuNro, cuentaMes, Valor, cuenta)
                    End If
                    rs3.Close
                'ElseIf rs2!tipoorigen = 1 And filtraConc(rs2!Origen, Co) = True Then
                ElseIf rs!tipoorigen = 1 And filtraConc(rs!Origen, Co) = True Then
                    concnro = rs!Origen
                    Ult_CoAc = rs!cuenta
                    cuenta = rs!cuenta
                    acuNro = "Null"
                    'Valor = rs!dlmontoacum
                    Valor = rs!dlmonto 'sebastian stremel 28/09/2012
                    
                    'concnro = rs2!concnro
                    'Ult_CoAc = rs2!concnro
                    'acuNro = "Null"
                    'Valor = rs2!dlimonto
                    
                    'Si existe el CO, hago update
                    StrSql = "SELECT sim_rep_pivotfinanzas.reppivnro,sim_rep_pivotfinanzas.concnro,sim_rep_pivotfinanzas.acunro"
                    StrSql = StrSql & " FROM sim_rep_pivotfinanzas"
                    StrSql = StrSql & " WHERE cuenta = '" & cuenta & "' AND ternro = " & rs!Empleado & " And bpronro = " & bpronro & " AND estrdabr='" & Estrdabr & "'"
                    StrSql = StrSql & " AND concnro = " & concnro
                    'StrSql = StrSql & " WHERE concnro = " & concnro & " AND ternro = " & rs!Empleado & " And bpronro = " & bpronro & " AND estrdabr='" & Estrdabr & "'"
                    'StrSql = StrSql & " WHERE concnro = " & concnro & " AND ternro = " & rs!Empleado & " And bpronro = " & bpronro & " AND pliqnro = " & rs!pliqnro
                    'Ejecuto
                    OpenRecordset StrSql, rs3
                    If Not rs3.EOF Then
                        'Update registro
                        Call ActualizarReg(rs3!reppivnro, cuentaMes, Valor)
                    Else
                        'Inserto registro
                        Call InsertarReg(bpronro, rs!Empleado, Estrdabr, rs!pliqnro, FechaInicio, FechaFin, concnro, acuNro, cuentaMes, Valor, cuenta)
                    End If
                    rs3.Close
                
                End If

                End If 'no es estructura vacia
            'rs2.MoveNext
            'Loop
        
                    
        'End If
        'rs2.Close
        actualizarProgreso (Progreso)
        rs.MoveNext
    Loop
End If
rs.Close
Flog.writeline "Proceso finalizado con éxito"



End Sub
'Obtiene la ultima fecha del mes y anio dado
Function UltimoDiaDeMes(ByVal Mes As Integer, ByVal Anio As Integer) As Date
Dim l_SDate
Dim l_MesProx

    l_SDate = DateSerial(Anio, Mes, "01")
    l_MesProx = DateAdd("m", 1, l_SDate)
    UltimoDiaDeMes = Day(DateAdd("d", -1, l_MesProx)) & "/" & Mes & "/" & Anio
    
End Function
Function filtraConc(concnro, Co)
    If Co <> "-1" Then
        Dim rs As New ADODB.Recordset
        StrSql = " SELECT concnro,conccod FROM concepto"
        StrSql = StrSql & " WHERE concnro = " & concnro
        StrSql = StrSql & " AND conccod IN(" & Co & ")"
        Flog.writeline StrSql
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            filtraConc = True
        Else
            filtraConc = False
        End If
        rs.Close
    Else
       filtraConc = True
    End If
End Function

Function filtraAcum(acuNro, Ac)
    If Ac <> "-1" Then
        Dim rs As New ADODB.Recordset
        StrSql = " SELECT acunro FROM acumulador "
        StrSql = StrSql & " WHERE acunro = " & acuNro
        StrSql = StrSql & " AND acunro IN (" & Ac & ")"
        Flog.writeline StrSql
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            filtraAcum = True
        Else
            filtraAcum = False
        End If
        rs.Close
    Else
       filtraAcum = True
    End If
End Function

Function PrimerDiaDeMes(Mes, Anio)
    Dim l_SDate
    Dim l_MesProx
    l_SDate = DateSerial(Anio, Mes, "01")
    l_MesProx = DateAdd("m", 1, l_SDate)
    PrimerDiaDeMes = Day(DateAdd("d", 0, l_MesProx)) & "/" & Mes & "/" & Anio
End Function


Function PerteneceEstruc(Ternro, tenro, fechadesde, fechahasta)
    Dim rs As New ADODB.Recordset
    StrSql = "SELECT estructura.estrdabr FROM sim_his_estructura"
    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = sim_his_estructura.estrnro "
    StrSql = StrSql & " WHERE sim_his_estructura.ternro =" & Ternro
    'StrSql = StrSql & " AND estrnro = " & estrnro
    StrSql = StrSql & " AND sim_his_estructura.tenro = " & tenro
    StrSql = StrSql & " AND (sim_his_estructura.htethasta >= " & ConvFecha(fechahasta) & " OR sim_his_estructura.htethasta >=" & ConvFecha(fechadesde) & " OR sim_his_estructura.htethasta is null)"
    StrSql = StrSql & " AND (sim_his_estructura.htetdesde <= " & ConvFecha(fechahasta) & ")"
    
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        PerteneceEstruc = rs!Estrdabr
    Else
        PerteneceEstruc = ""
        Flog.writeline Space(5) & "El empleado no pertenence al tipo estructura : " & tenro
    End If
    rs.Close
End Function










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
