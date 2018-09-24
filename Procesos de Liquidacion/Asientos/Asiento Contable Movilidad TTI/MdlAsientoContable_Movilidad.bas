Attribute VB_Name = "MdlAsientoContable"
Option Explicit

Const Version = 1.01
Const FechaVersion = "14/11/2005"




'-------------------------------------------------------------------------------
Const EsConcepto = 1
Const EsAcumulador = 2

Private Type TRegEstructura
    TE As Long
    Estructura As Long
    Porcentaje As Single
End Type

Global Inx             As Integer
Global Inxfin          As Integer
Global LI_1 As Integer
Global LI_2 As Integer
Global LI_3 As Integer

Global Inx_1 As Integer
Global Inx_2 As Integer
Global Inx_3 As Integer

Global vec_testr1(50)  As TRegEstructura
Global vec_testr2(50)  As TRegEstructura
Global vec_testr3(50)  As TRegEstructura

Global Descripcion As String
Global Cantidad As Single
Global CatidadVueltas As Long

Global rs_Proc_Vol As New ADODB.Recordset
Global rs_Mod_Linea As New ADODB.Recordset
Global rs_Empleado As New ADODB.Recordset
Global rs_Mod_Asiento As New ADODB.Recordset

Global BUF_mod_linea As New ADODB.Recordset
Global BUF_temp As New ADODB.Recordset

Global CantidadEmpleados As Long
Global PrimeraVez As Boolean

Global Mascara As String
Global Orden_CuentaContrapartida As Long
Global Descripcion_CuentaContrapartida As String




Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Asientos Contables.
' Autor      : FGZ
' Fecha      : 16/01/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim ArrParametros

    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    Nombre_Arch = PathFLog & "Asiento_Contable_Movilidad" & "-" & NroProcesoBatch & ".log"
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
    TiempoInicialProceso = GetTickCount
    
    'Abro la conexion
    On Error Resume Next
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
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & ",bprctiempo = 0 WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 52 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call batAsi00(NroProcesoBatch, bprcparam)
    End If
    
    TiempoFinalProceso = GetTickCount
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords


    'Cierro y libero todo
    If rs_batch_proceso.State = adStateOpen Then rs_batch_proceso.Close
    Set rs_batch_proceso = Nothing
    objConn.Close
    objconnProgreso.Close
    Flog.Close
    Exit Sub

ME_Main:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
End Sub


Public Sub batAsi00(ByVal bpronro As Long, ByVal Parametros As String) ', ByVal FechaDesde As Date, ByVal FechaHasta As Date)
' --------------------------------------------------------------------------------------------
' Descripcion: Programa que se ejecuta para generar Asiento Contable
'              Configurado en el tipo de proceso batch ?
' Autor      : Maximiliano Breglia
' Fecha      : 01/12/01
' Traduccion : FGZ
' Fecha      : 09/01/2004
' --------------------------------------------------------------------------------------------
Dim Mantener_Liq As Boolean
Dim Analisis_Detallado As Boolean
Dim Todos As Boolean

Dim pos1 As Integer
Dim pos2 As Integer

Dim NroVol   As Long
Dim Fecha    As Date
Dim Total    As Long
Dim NroAsientos As Long
Dim NroLineas As Long
Dim NroAsi As Long
Dim NroLin As Long

Dim rs_ProcVol As New ADODB.Recordset
Dim rs_Proc_V_modasi As New ADODB.Recordset
Dim rs_Mod_Asi As New ADODB.Recordset
Dim rs_Conf_cont As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset
Dim Abortado As Boolean


PrimeraVez = True
' El formato del mismo es (pronro.mantener Liq Ant.Guardar Nov.Analisis Det.Todos)
' Levanto cada parametro por separado, el separador de parametros es "."
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        NroVol = CLng(Mid(Parametros, pos1, pos2))
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        HACE_TRAZA = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
    End If
End If

StrSql = "SELECT * FROM proc_vol WHERE vol_cod =" & NroVol
OpenRecordset StrSql, rs_ProcVol

StrSql = "SELECT * FROM proc_v_modasi WHERE proc_v_modasi.vol_cod =" & NroVol
OpenRecordset StrSql, rs_Proc_V_modasi

'seteo las variables de progreso
Progreso = 0
CatidadVueltas = rs_Proc_V_modasi.RecordCount

Do While Not rs_Proc_V_modasi.EOF
    StrSql = "SELECT * FROM mod_asiento "
    StrSql = StrSql & " WHERE masinro =" & rs_Proc_V_modasi!Asi_Cod
    OpenRecordset StrSql, rs_Mod_Asi
    
    'seteo las variables de progreso
    CConceptosAProc = rs_Proc_V_modasi.RecordCount
    CEmpleadosAProc = rs_Mod_Asi.RecordCount
    'IncPorc = ((95 / CEmpleadosAProc) * (95 / CConceptosAProc)) / 95
    'IncPorc = 95 / (CEmpleadosAProc * CConceptosAProc)
    
    Do While Not rs_Mod_Asi.EOF
        StrSql = "SELECT * FROM conf_cont WHERE conf_cont.cofcnro =" & rs_Mod_Asi!cofcnro
        OpenRecordset StrSql, rs_Conf_cont
        
        If rs_Conf_cont.EOF Then
            Flog.writeline "No existe Proceso de Configuraci¢n asociado al Modelo de Asiento."
        Else
            If rs_Conf_cont!cofcacum = "" Then
                Flog.writeline "Falta ingresar el Archivo de Acumulación."
                Exit Sub
            End If
     
            'Comienzo la transaccion
            MyBeginTrans
            Abortado = False
             ' Ejecutar el proceso de acumulaci¢n por modelo de asiento
             Call Proc_Acumulacion(NroVol, rs_Proc_V_modasi!Asi_Cod, NroAsientos, NroLineas, Abortado, rs_Conf_cont!cofcacum)
            
            'Fin de la transaccion
            If Not Abortado Then
                MyCommitTrans
                NroAsi = NroAsi + 1
                NroLin = NroLin + NroLineas
            Else
                MyRollbackTrans
            End If
        End If
    
        rs_Mod_Asi.MoveNext
    Loop
    rs_Proc_V_modasi.MoveNext
Loop

' Actualizo
MyBeginTrans
    'Cuento la cantidad de lineas generadas
    StrSql = "SELECT count(*) as Lineas FROM linea_asi WHERE vol_cod =" & NroVol
    If rs_Aux.State = adStateOpen Then rs_Aux.Close
    OpenRecordset StrSql, rs_Aux
    If Not rs_Aux.EOF Then
        NroLin = rs_Aux!Lineas
    End If
    
    'Cuento la cantidad de asientos generados
    StrSql = "SELECT COUNT(DISTINCT masinro) AS Asientos FROM linea_asi WHERE vol_cod =" & NroVol
    If rs_Aux.State = adStateOpen Then rs_Aux.Close
    OpenRecordset StrSql, rs_Aux
    If Not rs_Aux.EOF Then
        NroAsi = rs_Aux!Asientos
    End If

    StrSql = "UPDATE proc_vol SET vol_reg_cab = " & NroAsi
    StrSql = StrSql & ", vol_reg_det =" & NroLin
    StrSql = StrSql & " WHERE proc_vol.vol_cod =" & NroVol
    objConn.Execute StrSql, , adExecuteNoRecords
MyCommitTrans

If rs_ProcVol.State = adStateOpen Then rs_ProcVol.Close
If rs_Proc_V_modasi.State = adStateOpen Then rs_Proc_V_modasi.Close
If rs_Mod_Asi.State = adStateOpen Then rs_Mod_Asi.Close
If rs_Conf_cont.State = adStateOpen Then rs_Conf_cont.Close
If rs_Aux.State = adStateOpen Then rs_Aux.Close

Set rs_Conf_cont = Nothing
Set rs_Mod_Asi = Nothing
Set rs_ProcVol = Nothing
Set rs_Proc_V_modasi = Nothing
Set rs_Aux = Nothing

Exit Sub

CE:
    MyRollbackTrans
    HuboError = True
    'Flog.writeline " Empleado abortado: " & buliq_empleado!empleg
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
End Sub

Public Sub Proc_Acumulacion(ByVal NroVol As Long, ByVal Asi_Cod As Long, ByRef NroAsientos As Long, ByRef NroLineas As Long, ByRef Abortado As Boolean, ByVal ProcesoGeneral As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Proceso que acumula las liquidaciones de Personal en un workfile temp de
'   asientos y lineas para luego ser volcado a la Interface Contable
' Autor      : Maximiliano Breglia
' Fecha      : 01/12/01
' Traduccion : FGZ
' Fecha      : 09/01/2004
' --------------------------------------------------------------------------------------------
Dim procesados As Integer
Dim signo_con     As Integer
Dim contador      As Integer
Dim nro_cuenta      As String
Dim monto_siento    As Single
Dim Aux_monto       As Single
Dim vestr1  As Integer
Dim vestr2  As Integer
Dim vestr3  As Integer
Dim distri_fijo   As Boolean

Dim rs_Conf_cont As New ADODB.Recordset
Dim rs_Proc As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Proceso As New ADODB.Recordset

Dim TipoE1 As Long
Dim TipoE2 As Long
Dim TipoE3 As Long

Dim Masinivternro1 As Long
Dim Masinivternro2 As Long
Dim Masinivternro3 As Long
Dim Recalculo As Boolean

Dim Fecha_Desde As Date
Dim Fecha_Hasta As Date

'Activo el manejador de errores local
On Error GoTo ME_Acumulacion

'Inicializacion de las variables
Inx = 1
Inxfin = 0

' para qué quiero esto ?
StrSql = "SELECT * FROM proc_vol WHERE proc_vol.vol_cod =" & NroVol
OpenRecordset StrSql, rs_Proc_Vol

StrSql = "SELECT * FROM mod_asiento where mod_asiento.masinro =" & Asi_Cod
OpenRecordset StrSql, rs_Mod_Asiento
If rs_Mod_Asiento.EOF Then
    Flog.writeline "No se encontró Mod_Asiento"
    Abortado = True
    Exit Sub
Else
    TipoE1 = IIf(Not IsNull(rs_Mod_Asiento!Masinivternro1), rs_Mod_Asiento!Masinivternro1, 0)
    TipoE2 = IIf(Not IsNull(rs_Mod_Asiento!Masinivternro2), rs_Mod_Asiento!Masinivternro2, 0)
    TipoE3 = IIf(Not IsNull(rs_Mod_Asiento!Masinivternro3), rs_Mod_Asiento!Masinivternro3, 0)
End If
    
StrSql = "SELECT * FROM conf_cont where conf_cont.cofcnro =" & rs_Mod_Asiento!cofcnro
OpenRecordset StrSql, rs_Conf_cont
        
If rs_Conf_cont.EOF Then
    Flog.writeline "No se encontró el conf_cont"
    Abortado = True
    Exit Sub
End If

StrSql = "SELECT * FROM proc_vol_pl" & _
         " INNER JOIN proc_vol_emp ON proc_vol_emp.pronro  = proc_vol_pl.pronro" & _
         " WHERE proc_vol_pl.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
         " AND proc_vol_emp.vol_cod = " & rs_Proc_Vol!Vol_Cod & _
         " ORDER BY proc_vol_emp.ternro"
OpenRecordset StrSql, rs_Proc
 
If PrimeraVez Then
    PrimeraVez = False
    CantidadEmpleados = rs_Proc.RecordCount
    Flog.writeline "Cantidad de Empleados = " & CantidadEmpleados
    If CantidadEmpleados = 0 Then
        CantidadEmpleados = 1
    End If
    IncPorc = 95 / (CatidadVueltas * CantidadEmpleados)
End If

Do While Not rs_Proc.EOF ' (1)
    contador = contador + 1
    
    StrSql = "SELECT * FROM empleado where empleado.ternro = " & rs_Proc!Ternro
    OpenRecordset StrSql, rs_Empleado
    If rs_Empleado.EOF Then
        Flog.writeline "No se encontro el legajo"
        Exit Sub
    Else
        Flog.writeline " ************* Legajo : " & rs_Empleado!empleg
    End If
       
    'Imputar los movimientos
    For Inx = 1 To 50
        vec_testr1(Inx).TE = 0
        vec_testr1(Inx).Estructura = 0
        vec_testr1(Inx).Porcentaje = 0
        
        vec_testr2(Inx).TE = 0
        vec_testr2(Inx).Estructura = 0
        vec_testr2(Inx).Porcentaje = 0
        
        vec_testr3(Inx).TE = 0
        vec_testr3(Inx).Estructura = 0
        vec_testr3(Inx).Porcentaje = 0
    Next Inx
    Inx = 1
    Masinivternro1 = IIf(Not EsNulo(rs_Mod_Asiento!Masinivternro1), rs_Mod_Asiento!Masinivternro1, 0)
    Masinivternro2 = IIf(Not EsNulo(rs_Mod_Asiento!Masinivternro2), rs_Mod_Asiento!Masinivternro2, 0)
    Masinivternro3 = IIf(Not EsNulo(rs_Mod_Asiento!Masinivternro3), rs_Mod_Asiento!Masinivternro3, 0)
    
    'Busco las fecha desde y hasta del proceso de Liq
    StrSql = "SELECT * FROM proceso WHERE pronro = " & rs_Proc!pronro
    If rs_Proceso.State = adStateOpen Then rs_Proceso.Close
    OpenRecordset StrSql, rs_Proceso
    If Not rs_Proceso.EOF Then
        If Not EsNulo(rs_Proceso!profecini) And Not EsNulo(rs_Proceso!profecfin) Then
            Fecha_Desde = rs_Proceso!profecini
            Fecha_Hasta = rs_Proceso!profecfin
            Call Movilidad(rs_Proc!Ternro, Asi_Cod, Fecha_Desde, Fecha_Hasta, rs_Mod_Asiento!Masinro, rs_Proc_Vol!Vol_Fec_Asiento, TipoE1, TipoE2, TipoE3, Masinivternro1, Masinivternro2, Masinivternro3, NroAsientos, NroLineas, Abortado)
            Recalculo = False
        Else
            Flog.writeline "Valor Nulo no Permitido en las fechas desde o hasta del proceso nro " & rs_Proc!pronro
        End If
    Else
        Flog.writeline "No se encuentran los datos del proceso. Siguiente."
    End If
    
    'Actualizar el progreso
    TiempoFinalProceso = GetTickCount
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    rs_Proc.MoveNext
Loop '(1)


'BALANCEO
Call Lin(NroAsientos, NroLineas)

StrSql = "DELETE linea_asi WHERE linea_asi.monto = 0"
objConn.Execute StrSql, , adExecuteNoRecords

'Puede que tenga que recalcular los valores como en el caso de promedio
If Recalculo Then
    Call Recalcular_lineas(Asi_Cod, NroVol)
End If

'Cierro y libero
If rs_Proceso.State = adStateOpen Then rs_Proceso.Close
Set rs_Proceso = Nothing

Exit Sub

ME_Acumulacion:
    Abortado = True
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
End Sub


Public Sub Acu_tmp(ByVal monto_a_imputar As Single, ByVal CuentaNiveladora As Boolean, ByVal signo As Integer, ByVal Descripcion As String, ByVal NroCuenta As String, ByVal Masinro As Long, ByVal Orden As Integer, ByVal Descripcion2 As String, ByVal Porcentaje As Single, ByVal Origen, ByVal TipoOrigen)
' --------------------------------------------------------------------------------------------
' Descripcion: Acumula en el tmp el monto dado como primer parametro
' Parametro 1 : monto
'             2 : INDICA SI ES CUENTA NIVELADORA
'             3 : INDICA EL SIGNO
' Autor      : Maximiliano Breglia
' Fecha      : 01/12/01
' Traduccion : FGZ
' Fecha      : 09/01/2004
' --------------------------------------------------------------------------------------------
Dim Aux_monto As Single
Dim NroCuentaContrapartida As String
Dim Inserta As Boolean
Dim rs_Linea_asi As New ADODB.Recordset


'Si es una linea nivelador, salir
'IF NOT {2} AND mod_linea.lin_desc = "Niveladora" THEN NEXT.
If Not CuentaNiveladora And Descripcion = "Niveladora" Then Exit Sub

'asignar el parametro (campo) a una variable para poder cambiarlo
Aux_monto = Redon(monto_a_imputar)
If Aux_monto = 0 Then
    Flog.writeline "El monto es 0. SALIR "
    Exit Sub
End If

'Guardo la mascara para buscar la cuenta de contrapartida
NroCuentaContrapartida = Mascara
Call ArmarCuenta(NroCuentaContrapartida, Masinro, Orden, Inserta)


Call ArmarCuenta(NroCuenta, Masinro, Orden, Inserta)

If Inserta Then
    StrSql = "SELECT * FROM linea_asi " & _
             " WHERE linea_asi.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
             " AND linea_asi.cuenta  ='" & NroCuenta & "'" & _
             " AND linea_asi.masinro =" & Masinro
    OpenRecordset StrSql, rs_Linea_asi
    
    If rs_Linea_asi.EOF Then
        StrSql = "INSERT INTO linea_asi (cuenta,vol_cod,masinro,linea,desclinea,monto)" & _
                 " VALUES ('" & NroCuenta & _
                 "'," & rs_Proc_Vol!Vol_Cod & _
                 "," & Masinro & _
                 "," & Orden & _
                 ",'" & Descripcion & _
                 "',0" & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    'Actualizo el Monto
    'Si tiene signo definido, tomar valor absoluto
    If signo <> 3 Then
        Aux_monto = IIf(Aux_monto >= 0, Aux_monto, -(Aux_monto))
    End If
    If signo = 3 Then 'como no lo toque lo dejo como viene si esta positivo suma sino resta
        Aux_monto = Aux_monto
    Else ' ya esta en valor absoluto
        If signo = 1 Then 'como esta en valor abs, si signo es 1 suma sino resta
            Aux_monto = Aux_monto
        Else
            Aux_monto = -Aux_monto
        End If
    End If
    
    StrSql = "UPDATE linea_asi SET monto = monto + " & Aux_monto & _
             " WHERE linea_asi.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
             " AND linea_asi.cuenta  ='" & NroCuenta & "'" & _
             " AND linea_asi.masinro =" & Masinro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'levanto los datos recien actualizados
    If rs_Linea_asi.State = adStateOpen Then rs_Linea_asi.Close
    StrSql = "SELECT * FROM linea_asi " & _
             " WHERE linea_asi.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
             " AND linea_asi.cuenta  ='" & NroCuenta & "'" & _
             " AND linea_asi.masinro =" & Masinro
    OpenRecordset StrSql, rs_Linea_asi
    If Not rs_Linea_asi.EOF Then
        If HACE_TRAZA Then
            StrSql = "INSERT INTO detalle_asi (masinro, cuenta,dlcantidad,dlcosto1,dlcosto2,dlcosto3,dlcosto4,dldescripcion " & _
                     ",dlmonto,dlmontoacum,dlporcentaje,ternro,empleg,lin_orden,terape,vol_cod, origen, tipoorigen)" & _
                     " VALUES (" & rs_Linea_asi!Masinro & _
                     ",'" & rs_Linea_asi!cuenta & _
                     "'," & Cantidad & _
                     "," & vec_testr1(Inx).Estructura & _
                     "," & vec_testr2(Inx).Estructura & _
                     "," & vec_testr3(Inx).Estructura & _
                     ",0" & _
                     ",'" & Descripcion2 & _
                     "'," & Aux_monto & _
                     "," & rs_Linea_asi!Monto & _
                     "," & Porcentaje & _
                     "," & rs_Empleado!Ternro & _
                     "," & rs_Empleado!empleg & _
                     "," & Orden & _
                     ",'" & rs_Empleado!terape & _
                     "'," & rs_Linea_asi!Vol_Cod & _
                     "," & Origen & _
                     "," & TipoOrigen & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    
    '----------------------------------------------------------
    'Actualizo la contrapartida
    StrSql = "SELECT * FROM linea_asi " & _
             " WHERE linea_asi.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
             " AND linea_asi.cuenta  ='" & NroCuentaContrapartida & "'" & _
             " AND linea_asi.masinro =" & Masinro
    OpenRecordset StrSql, rs_Linea_asi
    
    If rs_Linea_asi.EOF Then
        StrSql = "INSERT INTO linea_asi (cuenta,vol_cod,masinro,linea,desclinea,monto)" & _
                 " VALUES ('" & NroCuentaContrapartida & _
                 "'," & rs_Proc_Vol!Vol_Cod & _
                 "," & Masinro & _
                 "," & Orden_CuentaContrapartida & _
                 ",'" & Descripcion_CuentaContrapartida & _
                 "',0" & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    'Actualizo el Monto
    'Si tiene signo definido, tomar valor absoluto
    If signo <> 3 Then
        Aux_monto = IIf(Aux_monto >= 0, Aux_monto, -(Aux_monto))
    End If
    If signo = 3 Then 'como no lo toque lo dejo como viene si esta positivo suma sino resta
        Aux_monto = Aux_monto
    Else ' ya esta en valor absoluto
        If signo = 1 Then 'como esta en valor abs, si signo es 1 suma sino resta
            Aux_monto = Aux_monto
        Else
            Aux_monto = -Aux_monto
        End If
    End If
    
    StrSql = "UPDATE linea_asi SET monto = monto + " & Aux_monto & _
             " WHERE linea_asi.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
             " AND linea_asi.cuenta  ='" & NroCuentaContrapartida & "'" & _
             " AND linea_asi.masinro =" & Masinro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'levanto los datos recien actualizados
    If rs_Linea_asi.State = adStateOpen Then rs_Linea_asi.Close
    StrSql = "SELECT * FROM linea_asi " & _
             " WHERE linea_asi.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
             " AND linea_asi.cuenta  ='" & NroCuentaContrapartida & "'" & _
             " AND linea_asi.masinro =" & Masinro
    OpenRecordset StrSql, rs_Linea_asi
    If Not rs_Linea_asi.EOF Then
        If HACE_TRAZA Then
            StrSql = "INSERT INTO detalle_asi (masinro, cuenta,dlcantidad,dlcosto1,dlcosto2,dlcosto3,dlcosto4,dldescripcion " & _
                     ",dlmonto,dlmontoacum,dlporcentaje,ternro,empleg,lin_orden,terape,vol_cod, origen, tipoorigen)" & _
                     " VALUES (" & rs_Linea_asi!Masinro & _
                     ",'" & rs_Linea_asi!cuenta & _
                     "'," & Cantidad & _
                     "," & vec_testr1(Inx).Estructura & _
                     "," & vec_testr2(Inx).Estructura & _
                     "," & vec_testr3(Inx).Estructura & _
                     ",0" & _
                     ",'" & Descripcion2 & _
                     "'," & Aux_monto & _
                     "," & rs_Linea_asi!Monto & _
                     "," & Porcentaje & _
                     "," & rs_Empleado!Ternro & _
                     "," & rs_Empleado!empleg & _
                     "," & Orden & _
                     ",'" & rs_Empleado!terape & _
                     "'," & rs_Linea_asi!Vol_Cod & _
                     "," & Origen & _
                     "," & TipoOrigen & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
Else
    Flog.writeline "ArmarCuenta_Con_Apertura. No Inserta"
End If


End Sub

Public Sub Acu_tmp_Con_Apertura(ByVal Tercero As Long, ByVal FechaAsiento As Date, ByVal TipoE1 As Long, ByVal TipoE2 As Long, ByVal TipoE3 As Long, ByVal monto_a_imputar As Single, ByVal CuentaNiveladora As Boolean, ByVal signo As Integer, ByVal Descripcion As String, ByVal NroCuenta As String, ByVal Masinro As Long, ByVal Orden As Integer, ByVal Descripcion2 As String, ByVal Porcentaje As Single, ByVal Origen, ByVal TipoOrigen)
' --------------------------------------------------------------------------------------------
' Descripcion: Acumula en el tmp el monto dado como primer parametro
' Parametro 1 : monto
'             2 : INDICA SI ES CUENTA NIVELADORA
'             3 : INDICA EL SIGNO
' Autor      : Maximiliano Breglia
' Fecha      : 01/12/01
' Traduccion : FGZ
' Fecha      : 09/01/2004
' --------------------------------------------------------------------------------------------
Dim Aux_monto As Single
Dim Total_Dias As Integer
Dim Total_Monto As Single
Dim Aux_Dias As Integer
Dim Inserta As Boolean
Dim Fecha_Desde As Date
Dim Fecha_Hasta As Date
Dim Aux_Fecha_Desde As Date
Dim Aux_Fecha_Hasta As Date
Dim Apertura As Boolean
Dim Estructura1 As Long
Dim Estructura2 As Long
Dim Estructura3 As Long
Dim Aux_NroCuenta As String

Dim rs_Linea_asi As New ADODB.Recordset
Dim rs_Estructura1 As New ADODB.Recordset
Dim rs_Estructura2 As New ADODB.Recordset
Dim rs_Estructura3 As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset

'Si es una linea nivelador, salir
If Not CuentaNiveladora And Descripcion = "Niveladora" Then Exit Sub

'asignar el parametro (campo) a una variable para poder cambiarlo
Aux_monto = Redon(monto_a_imputar)
If Aux_monto = 0 Then
    Flog.writeline "El monto es 0. SALIR "
    Exit Sub
End If
'en principio queda fijo para SAC
' busco  el mes del periodo que estamos liquidando para establecer las fechas desde y hasta
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & rs_Proc_Vol!pliqnro
OpenRecordset StrSql, rs_Periodo
If Not rs_Periodo.EOF Then
    If rs_Periodo!pliqmes > 6 Then
        Fecha_Desde = CDate("01/07/" & rs_Periodo!pliqanio)
        Fecha_Hasta = CDate("31/12/" & rs_Periodo!pliqanio)
    Else
        Fecha_Desde = CDate("01/01/" & rs_Periodo!pliqanio)
        Fecha_Hasta = CDate("30/06/" & rs_Periodo!pliqanio)
    End If
Else 'No se encontro el periodo
    Flog.writeline "No se encontro el periodo. Se usara la fecha actual para establecer las fechas desde y hasta"
    If Month(Date) > 6 Then
        Fecha_Desde = CDate("01/07/" & Year(Date))
        Fecha_Hasta = CDate("31/12/" & Year(Date))
    Else
        Fecha_Desde = CDate("01/01/" & Year(Date))
        Fecha_Hasta = CDate("30/06/" & Year(Date))
    End If
End If

Total_Dias = DateDiff("d", Fecha_Desde, Fecha_Hasta) + 1
Total_Monto = Redon(monto_a_imputar)

Apertura = False
Estructura1 = 0
Estructura2 = 0
Estructura3 = 0

Aux_Fecha_Desde = Fecha_Desde
Aux_Fecha_Hasta = Fecha_Hasta
'ciclo por los tres tipos de estructura
If TipoE1 <> 0 Then
    StrSql = " SELECT * FROM his_estructura " & _
             " WHERE ternro = " & Tercero & " AND " & _
             " tenro =" & TipoE1 & " AND " & _
             " (htetdesde <= " & ConvFecha(Fecha_Hasta) & ") AND " & _
             " ((" & ConvFecha(Fecha_Desde) & " <= htethasta) or (htethasta is null))" & _
             " ORDER BY htetdesde "
    OpenRecordset StrSql, rs_Estructura1
    
    Do While Not rs_Estructura1.EOF
        Estructura1 = rs_Estructura1!estrnro
        Aux_Fecha_Desde = IIf(rs_Estructura1!htetdesde < Fecha_Desde, Fecha_Desde, rs_Estructura1!htetdesde)
        If Not IsNull(rs_Estructura1!htethasta) Then
            Aux_Fecha_Hasta = IIf(rs_Estructura1!htethasta > Fecha_Hasta, Fecha_Hasta, rs_Estructura1!htethasta)
        Else
            Aux_Fecha_Hasta = Fecha_Hasta
        End If
        
        If TipoE2 <> 0 Then
            StrSql = " SELECT * FROM his_estructura " & _
                     " WHERE ternro = " & Tercero & " AND " & _
                     " tenro =" & TipoE2 & " AND " & _
                     " (htetdesde <= " & ConvFecha(Aux_Fecha_Hasta) & ") AND " & _
                     " ((" & ConvFecha(Aux_Fecha_Desde) & " <= htethasta) or (htethasta is null))" & _
                     " ORDER BY htetdesde "
            OpenRecordset StrSql, rs_Estructura2
    
            Do While Not rs_Estructura2.EOF
                Estructura2 = rs_Estructura2!estrnro
                Aux_Fecha_Desde = IIf(rs_Estructura2!htetdesde < Fecha_Desde, Fecha_Desde, rs_Estructura2!htetdesde)
                If Not IsNull(rs_Estructura2!htethasta) Then
                    Aux_Fecha_Hasta = IIf(rs_Estructura2!htethasta > Fecha_Hasta, Fecha_Hasta, rs_Estructura2!htethasta)
                Else
                    Aux_Fecha_Hasta = Fecha_Hasta
                End If
    
                If TipoE3 <> 0 Then
                    StrSql = " SELECT * FROM his_estructura " & _
                             " WHERE ternro = " & Tercero & " AND " & _
                             " tenro =" & TipoE3 & " AND " & _
                             " (htetdesde <= " & ConvFecha(Aux_Fecha_Hasta) & ") AND " & _
                             " ((" & ConvFecha(Aux_Fecha_Desde) & " <= htethasta) or (htethasta is null))" & _
                             " ORDER BY htetdesde "
                    OpenRecordset StrSql, rs_Estructura3
            
                    Do While Not rs_Estructura3.EOF
                        Estructura3 = rs_Estructura3!estrnro
                        Aux_Fecha_Desde = IIf(rs_Estructura3!htetdesde < Fecha_Desde, Fecha_Desde, rs_Estructura3!htetdesde)
                        If Not IsNull(rs_Estructura3!htethasta) Then
                            Aux_Fecha_Hasta = IIf(rs_Estructura3!htethasta > Fecha_Hasta, Fecha_Hasta, rs_Estructura3!htethasta)
                        Else
                            Aux_Fecha_Hasta = Fecha_Hasta
                        End If
                        
                        Aux_NroCuenta = NroCuenta
                        Call ArmarCuenta_Con_Apertura(Estructura1, Estructura2, Estructura3, Aux_NroCuenta, Masinro, Orden, Inserta)
                        If Inserta Then
                            Aux_Dias = CantidadDeDias(Fecha_Desde, Fecha_Hasta, Aux_Fecha_Desde, Aux_Fecha_Hasta)
                            Aux_monto = (Aux_Dias * Total_Monto) / Total_Dias
                    
                            Call Inserta_Linea_Asiento(Aux_monto, signo, Descripcion, Aux_NroCuenta, Masinro, Orden, Descripcion2, Porcentaje, Origen, TipoOrigen, Estructura1, Estructura2, Estructura3)
                        End If
            
                        rs_Estructura3.MoveNext
                    Loop
                Else
                    Aux_NroCuenta = NroCuenta
                    Call ArmarCuenta_Con_Apertura(Estructura1, Estructura2, Estructura3, Aux_NroCuenta, Masinro, Orden, Inserta)
                    If Inserta Then
                        Aux_Dias = CantidadDeDias(Fecha_Desde, Fecha_Hasta, Aux_Fecha_Desde, Aux_Fecha_Hasta)
                        Aux_monto = (Aux_Dias * Total_Monto) / Total_Dias
                
                        Call Inserta_Linea_Asiento(Aux_monto, signo, Descripcion, Aux_NroCuenta, Masinro, Orden, Descripcion2, Porcentaje, Origen, TipoOrigen, Estructura1, Estructura2, Estructura3)
                    Else
                        Flog.writeline "ArmarCuenta_Con_Apertura. No Inserta"
                    End If
                End If
                rs_Estructura2.MoveNext
            Loop
        Else
            Aux_NroCuenta = NroCuenta
            Call ArmarCuenta_Con_Apertura(Estructura1, Estructura2, Estructura3, Aux_NroCuenta, Masinro, Orden, Inserta)
            If Inserta Then
                Aux_Dias = CantidadDeDias(Fecha_Desde, Fecha_Hasta, Aux_Fecha_Desde, Aux_Fecha_Hasta)
                Aux_monto = (Aux_Dias * Total_Monto) / Total_Dias
        
                Call Inserta_Linea_Asiento(Aux_monto, signo, Descripcion, Aux_NroCuenta, Masinro, Orden, Descripcion2, Porcentaje, Origen, TipoOrigen, Estructura1, Estructura2, Estructura3)
            End If
        End If
        rs_Estructura1.MoveNext
    Loop
    Apertura = True
End If


If Not Apertura Then
    Aux_NroCuenta = NroCuenta
    Call ArmarCuenta_Con_Apertura(Estructura1, Estructura2, Estructura3, Aux_NroCuenta, Masinro, Orden, Inserta)
    
    If Inserta Then
        Aux_Dias = CantidadDeDias(Fecha_Desde, Fecha_Hasta, Aux_Fecha_Desde, Aux_Fecha_Hasta)
        Aux_monto = Redon((Aux_Dias * Total_Monto) / Total_Dias)

        Call Inserta_Linea_Asiento(Aux_monto, signo, Descripcion, Aux_NroCuenta, Masinro, Orden, Descripcion2, Porcentaje, Origen, TipoOrigen, Estructura1, Estructura2, Estructura3)
    End If
End If


'cerrar todo
If rs_Estructura1.State = adStateOpen Then rs_Estructura1.Close
Set rs_Estructura1 = Nothing
If rs_Estructura2.State = adStateOpen Then rs_Estructura2.Close
Set rs_Estructura2 = Nothing
If rs_Estructura3.State = adStateOpen Then rs_Estructura3.Close
Set rs_Estructura3 = Nothing
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
Set rs_Periodo = Nothing

End Sub

Private Sub Inserta_Linea_Asiento(ByVal Aux_monto As Single, ByVal signo As Integer, ByVal Descripcion As String, ByVal NroCuenta As String, ByVal Masinro As Long, ByVal Orden As Integer, ByVal Descripcion2 As String, ByVal Porcentaje As Single, ByVal Origen, ByVal TipoOrigen, ByVal Estr1 As Long, ByVal Estr2 As Long, ByVal Estr3 As Long)
Dim rs_Linea_asi As New ADODB.Recordset

        Aux_monto = Redon(Aux_monto)

        StrSql = "SELECT * FROM linea_asi " & _
                 " WHERE linea_asi.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
                 " AND linea_asi.cuenta  ='" & NroCuenta & "'" & _
                 " AND linea_asi.masinro =" & Masinro
        OpenRecordset StrSql, rs_Linea_asi
        
        If rs_Linea_asi.EOF Then
            StrSql = "INSERT INTO linea_asi (cuenta,vol_cod,masinro,linea,desclinea,monto)" & _
                     " VALUES ('" & NroCuenta & _
                     "'," & rs_Proc_Vol!Vol_Cod & _
                     "," & Masinro & _
                     "," & Orden & _
                     ",'" & Descripcion & _
                     "',0" & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        'Actualizo el Monto
        'Si tiene signo definido, tomar valor absoluto
        If signo <> 3 Then
            Aux_monto = IIf(Aux_monto >= 0, Aux_monto, -(Aux_monto))
        End If
        If signo = 3 Then 'como no lo toque lo dejo como viene si esta positivo suma sino resta
            Aux_monto = Aux_monto
        Else ' ya esta en valor absoluto
            If signo = 1 Then 'como esta en valor abs, si signo es 1 suma sino resta
                Aux_monto = Aux_monto
            Else
                Aux_monto = -Aux_monto
            End If
        End If
        
        StrSql = "UPDATE linea_asi SET monto = monto + " & Aux_monto & _
                 " WHERE linea_asi.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
                 " AND linea_asi.cuenta  ='" & NroCuenta & "'" & _
                 " AND linea_asi.masinro =" & Masinro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'levanto los datos recien actualizados
        If rs_Linea_asi.State = adStateOpen Then rs_Linea_asi.Close
        StrSql = "SELECT * FROM linea_asi " & _
                 " WHERE linea_asi.vol_cod =" & rs_Proc_Vol!Vol_Cod & _
                 " AND linea_asi.cuenta  ='" & NroCuenta & "'" & _
                 " AND linea_asi.masinro =" & Masinro
        OpenRecordset StrSql, rs_Linea_asi
        
        If HACE_TRAZA Then
            'creaci¢n de un registro de traza
'            StrSql = "INSERT INTO detalle_asi (masinro, cuenta,dlcantidad,dlcosto1,dlcosto2,dlcosto3,dlcosto4,dldescripcion " & _
'                     ",dlmonto,dlmontoacum,dlporcentaje,ternro,empleg,lin_orden,terape,vol_cod,origen,tipoorigen)" & _
'                     " VALUES (" & rs_Linea_asi!Masinro & _
'                     ",'" & rs_Linea_asi!cuenta & _
'                     "'," & Cantidad & _
'                     ",0" & _
'                     ",0" & _
'                     ",0" & _
'                     ",0" & _
'                     ",'" & Descripcion2 & _
'                     "'," & Aux_monto & _
'                     "," & rs_Linea_asi!Monto & _
'                     "," & Porcentaje & _
'                     "," & rs_Empleado!ternro & _
'                     "," & rs_Empleado!empleg & _
'                     "," & Orden & _
'                     ",'" & rs_Empleado!terape & _
'                     "'," & rs_Linea_asi!Vol_Cod & _
'                     "," & Origen & _
'                     "," & TipoOrigen & _
'                     ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
            StrSql = "INSERT INTO detalle_asi (masinro, cuenta,dlcantidad,dlcosto1,dlcosto2,dlcosto3,dlcosto4,dldescripcion " & _
                     ",dlmonto,dlmontoacum,dlporcentaje,ternro,empleg,lin_orden,terape,vol_cod,origen,tipoorigen)" & _
                     " VALUES (" & rs_Linea_asi!Masinro & _
                     ",'" & rs_Linea_asi!cuenta & _
                     "'," & Cantidad & _
                     "," & Estr1 & _
                     "," & Estr2 & _
                     "," & Estr3 & _
                     ",0" & _
                     ",'" & Descripcion2 & _
                     "'," & Aux_monto & _
                     "," & rs_Linea_asi!Monto & _
                     "," & Porcentaje & _
                     "," & rs_Empleado!Ternro & _
                     "," & rs_Empleado!empleg & _
                     "," & Orden & _
                     ",'" & rs_Empleado!terape & _
                     "'," & rs_Linea_asi!Vol_Cod & _
                     "," & Origen & _
                     "," & TipoOrigen & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

If rs_Linea_asi.State = adStateOpen Then rs_Linea_asi.Close
Set rs_Linea_asi = Nothing
End Sub


Public Sub Lin(ByRef NroAsientos As Long, ByRef NroLineas As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Asignar D/H a cada linea y balancer los asientos
' Autor      :
' Fecha      : 01/12/01
' Traduccion : FGZ
' Fecha      : 15/01/2004
' --------------------------------------------------------------------------------------------
Dim monto_asiento As Double
Dim monto_asiento_dos As Double
Dim linea_asi_dh As Integer
Dim Ultimo_Asi_Cod As Long
Dim Actualizo As Boolean

Dim rs_Mod_Linea_Balanceo As New ADODB.Recordset
Dim rs_asiento As New ADODB.Recordset
Dim rs_Linea_asi As New ADODB.Recordset

StrSql = "SELECT * FROM linea_asi "
StrSql = StrSql & " WHERE linea_asi.masinro = " & rs_Mod_Asiento!Masinro
StrSql = StrSql & " AND linea_asi.Vol_Cod = " & rs_Proc_Vol!Vol_Cod
OpenRecordset StrSql, BUF_temp

Do While Not BUF_temp.EOF
    StrSql = "SELECT * FROM mod_linea " & _
             " WHERE mod_linea.masinro = " & BUF_temp!Masinro & _
             " AND mod_linea.linaorden =" & BUF_temp!linea & _
             " ORDER BY masinro,linaorden"
    OpenRecordset StrSql, BUF_mod_linea
    
    Ultimo_Asi_Cod = -1
    Do While Not BUF_mod_linea.EOF
        If BUF_mod_linea!Masinro <> Ultimo_Asi_Cod Then 'es el primero
            Ultimo_Asi_Cod = BUF_mod_linea!Masinro
        
            'Creo el asiento
            StrSql = "SELECT * FROM asiento " & _
                     " WHERE masinro = " & BUF_mod_linea!Masinro & _
                     " AND vol_cod =" & rs_Proc_Vol!Vol_Cod
            OpenRecordset StrSql, rs_asiento
            
            If rs_asiento.EOF Then
                StrSql = "INSERT INTO asiento (masinro,asidebe,asihaber,vol_cod) " & _
                         " VALUES (" & BUF_mod_linea!Masinro & _
                         ",0" & _
                         ",0" & _
                         "," & rs_Proc_Vol!Vol_Cod & _
                         ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                monto_asiento = 0
                monto_asiento_dos = 0
                NroAsientos = NroAsientos + 1
            End If
        End If
                
        If rs_asiento.State = adStateOpen Then rs_asiento.Close
        StrSql = "SELECT * FROM asiento " & _
                 " WHERE masinro = " & BUF_mod_linea!Masinro & _
                 " AND vol_cod =" & rs_Proc_Vol!Vol_Cod
        OpenRecordset StrSql, rs_asiento

        
        StrSql = "SELECT * FROM linea_asi " & _
                 " WHERE masinro = " & BUF_temp!Masinro & _
                 " AND vol_cod =" & rs_Proc_Vol!Vol_Cod & _
                 " AND cuenta ='" & BUF_temp!cuenta & "'"
        OpenRecordset StrSql, rs_Linea_asi
        
        NroLineas = NroLineas + 1
        
        If BUF_mod_linea!linaD_H = 2 Then
            If (BUF_temp!Monto) > 0 Then
                linea_asi_dh = -1
            Else
                linea_asi_dh = 0
            End If
        Else
            If BUF_mod_linea!linaD_H = 0 Then
                linea_asi_dh = -1
            Else
                linea_asi_dh = 0
            End If
        End If
        StrSql = "UPDATE linea_asi SET dh = " & linea_asi_dh & _
                 ",monto =" & Abs(rs_Linea_asi!Monto) & _
                 " WHERE masinro = " & BUF_temp!Masinro & _
                 " AND vol_cod =" & rs_Proc_Vol!Vol_Cod & _
                 " AND cuenta ='" & BUF_temp!cuenta & "'"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - leo nuevamente porque actualicé
        StrSql = "SELECT * FROM linea_asi " & _
                 " WHERE masinro = " & BUF_temp!Masinro & _
                 " AND vol_cod =" & rs_Proc_Vol!Vol_Cod & _
                 " AND cuenta ='" & BUF_temp!cuenta & "'"
        OpenRecordset StrSql, rs_Linea_asi
        
        'monto_asiento = monto_asiento + IIf(CBool(linea_asi_dh), rs_Linea_asi!Monto, -rs_Linea_asi!Monto)
        'monto_asiento_dos = monto_asiento_dos + IIf(CBool(linea_asi_dh), Round(rs_Linea_asi!Monto, 2), Round(-rs_Linea_asi!Monto, 2))
        monto_asiento = monto_asiento + IIf(CBool(linea_asi_dh), rs_Linea_asi!Monto, 0)
        monto_asiento_dos = monto_asiento_dos + IIf(Not CBool(linea_asi_dh), rs_Linea_asi!Monto, 0)

       If CBool(linea_asi_dh) Then
            StrSql = "UPDATE asiento SET asidebe = asidebe + " & rs_Linea_asi!Monto & _
                     " WHERE masinro = " & BUF_mod_linea!Masinro & _
                     " AND vol_cod =" & rs_Proc_Vol!Vol_Cod
         Else
            StrSql = "UPDATE asiento SET asihaber = asihaber + " & rs_Linea_asi!Monto & _
                     " WHERE masinro = " & BUF_mod_linea!Masinro & _
                     " AND vol_cod =" & rs_Proc_Vol!Vol_Cod
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
'        'si es el ultimo ==> actualizo
'        BUF_mod_linea.MoveNext
'        If BUF_mod_linea.EOF Then
'            Actualizo = True
'        Else
'            If BUF_mod_linea!Masinro <> Ultimo_Asi_Cod And monto_asiento <> 0 Then
'                Actualizo = True
'            Else
'                Actualizo = False
'            End If
'        End If
'        BUF_mod_linea.MovePrevious
       
       
        'si es el ultimo ==> actualizo
        BUF_temp.MoveNext
        If BUF_temp.EOF Then
            If (Abs(monto_asiento) - Abs(monto_asiento_dos)) <> 0 Then
                Flog.writeline "Diferencia = " & Abs(monto_asiento) - Abs(monto_asiento_dos)
                Actualizo = True
            Else
                Actualizo = False
                Flog.writeline "No hay Diferencia. Debe = " & monto_asiento & " y Haber = " & monto_asiento_dos
            End If
        Else
            If BUF_mod_linea!Masinro <> Ultimo_Asi_Cod And (Abs(monto_asiento) - Abs(monto_asiento_dos)) <> 0 Then
'            If BUF_mod_linea!Masinro <> Ultimo_Asi_Cod And monto_asiento <> 0 Then
                Actualizo = True
                Flog.writeline "Diferencia = " & Abs(monto_asiento) - Abs(monto_asiento_dos)
            Else
                Actualizo = False
                Flog.writeline "No hay Diferencia. Debe = " & monto_asiento & " y Haber = " & monto_asiento_dos
            End If
        End If
        BUF_temp.MovePrevious
       
       
        If Actualizo Then
            ' Buscar la linea de Balanceo del asiento
            StrSql = "SELECT * FROM mod_linea " & _
                     " WHERE masinro = " & BUF_mod_linea!Masinro & _
                     " AND upper(linadesc) = 'NIVELADORA'"
            OpenRecordset StrSql, rs_Mod_Linea_Balanceo
            
            If Not rs_Mod_Linea_Balanceo.EOF Then
                monto_asiento = Abs(monto_asiento) - Abs(monto_asiento_dos)
                Flog.writeline "Genera niveladora por " & monto_asiento
                Call Acu_tmp(monto_asiento, True, 3, "Niveladora", rs_Mod_Linea_Balanceo!linacuenta, rs_Mod_Linea_Balanceo!Masinro, rs_Mod_Linea_Balanceo!LinaOrden, "", 100, 0, EsConcepto)

                ' ASIGNAR EL D/H
                If monto_asiento > 0 Then
                    linea_asi_dh = 0
                Else
                    linea_asi_dh = -1
                End If
                
'                StrSql = "UPDATE linea_asi SET dh = " & linea_asi_dh & _
'                         ",monto =" & Abs(rs_Linea_asi!Monto) & _
'                         " WHERE masinro = " & BUF_temp!Masinro & _
'                         " AND vol_cod =" & rs_Proc_Vol!Vol_Cod & _
'                         " AND cuenta ='" & rs_Mod_Linea_Balanceo!linacuenta & "'"
'                objConn.Execute StrSql, , adExecuteNoRecords
                
                StrSql = "UPDATE linea_asi SET dh = " & linea_asi_dh & _
                         ",monto =" & Abs(monto_asiento) & _
                         " WHERE masinro = " & BUF_temp!Masinro & _
                         " AND vol_cod =" & rs_Proc_Vol!Vol_Cod & _
                         " AND cuenta ='" & rs_Mod_Linea_Balanceo!linacuenta & "'"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                If CBool(linea_asi_dh) Then
                     StrSql = "UPDATE asiento SET asidebe = asidebe + " & Abs(monto_asiento) & _
                              " WHERE masinro = " & BUF_temp!Masinro & _
                              " AND vol_cod =" & rs_Proc_Vol!Vol_Cod
                  Else
                     StrSql = "UPDATE asiento SET asihaber = asihaber + " & Abs(monto_asiento) & _
                              " WHERE masinro = " & BUF_temp!Masinro & _
                              " AND vol_cod =" & rs_Proc_Vol!Vol_Cod
                 End If
                 objConn.Execute StrSql, , adExecuteNoRecords
                
            End If
        End If
        
        BUF_mod_linea.MoveNext
    Loop
    BUF_temp.MoveNext
Loop


End Sub


Public Sub ArmarCuenta(ByRef NroCuenta As String, ByVal Masinro As Long, ByVal LinaOrden As Long, ByRef Genera As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : Maximiliano Breglia
' Fecha      : 01/12/01
' Traduccion : FGZ
' Fecha      : 09/01/2004
' --------------------------------------------------------------------------------------------
Dim Aux_Cuenta As String
Dim Aux_Legajo As String
Dim Estructura1 As String
Dim Estructura2 As String
Dim Estructura3 As String

Dim i As Integer
Dim ch As String
Dim CantL As Integer
Dim CantE As Integer
Dim TipoE As String
Dim TipoE_Actual As String
Dim EsEstructura As Boolean
Dim Termino As Boolean

Dim PosE1 As Integer
Dim PosE2 As Integer
Dim PosE3 As Integer

Dim rs_Estructura1 As New ADODB.Recordset
Dim rs_Estructura2 As New ADODB.Recordset
Dim rs_Estructura3 As New ADODB.Recordset
Dim rs_Filtro As New ADODB.Recordset

Aux_Cuenta = NroCuenta
'aux_Cuenta = rs_Mod_Linea!linacuenta
Aux_Legajo = rs_Empleado!empleg
Genera = True

If IsNull(vec_testr1(Inx_1).Estructura) Or vec_testr1(Inx_1).Estructura = 0 Then
    Estructura1 = "00000000000000000000"
Else
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & vec_testr1(Inx_1).Estructura
    OpenRecordset StrSql, rs_Estructura1
    
    If Not rs_Estructura1.EOF Then
        'reviso que tenga un filtro
        StrSql = "SELECT * FROM mod_lin_filtro "
        StrSql = StrSql & " WHERE masinro = " & Masinro
        StrSql = StrSql & " AND linaorden = " & LinaOrden
        StrSql = StrSql & " AND tenro = " & rs_Estructura1!tenro
        If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
        OpenRecordset StrSql, rs_Filtro
        If Not rs_Filtro.EOF Then
            'tiene filtro
            StrSql = "SELECT * FROM mod_lin_filtro "
            StrSql = StrSql & " WHERE masinro = " & Masinro
            StrSql = StrSql & " AND linaorden = " & LinaOrden
            StrSql = StrSql & " AND estrnro = " & rs_Estructura1!estrnro
            If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
            OpenRecordset StrSql, rs_Filtro
            
            If Not rs_Filtro.EOF Then
                If IsNumeric(rs_Estructura1!estrcodext) Then
                    Estructura1 = rs_Estructura1!estrcodext & "00000000000000000000"
                Else
                    Estructura1 = IIf(IsNull(rs_Estructura1!estrcodext), "00000000000000000000", rs_Estructura1!estrcodext & "00000000000000000000")
                End If
            Else
                Estructura1 = "00000000000000000000"
                Genera = False
            End If
        Else
            'no tiene filtro
            If IsNumeric(rs_Estructura1!estrcodext) Then
                Estructura1 = rs_Estructura1!estrcodext & "00000000000000000000"
            Else
                Estructura1 = IIf(IsNull(rs_Estructura1!estrcodext), "00000000000000000000", rs_Estructura1!estrcodext & "00000000000000000000")
            End If
        End If
    Else
        Estructura1 = "00000000000000000000"
    End If
End If
Estructura1 = Left(Estructura1, 20)

If IsNull(vec_testr2(Inx_2).Estructura) Or vec_testr2(Inx_2).Estructura = 0 Then
    Estructura2 = "00000000000000000000"
Else
    StrSql = " SELECT estrcodext,estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & vec_testr2(Inx_2).Estructura
    OpenRecordset StrSql, rs_Estructura2
    
    If Not rs_Estructura2.EOF Then
        'reviso que tenga un filtro
        StrSql = "SELECT * FROM mod_lin_filtro "
        StrSql = StrSql & " WHERE masinro = " & Masinro
        StrSql = StrSql & " AND linaorden = " & LinaOrden
        StrSql = StrSql & " AND tenro = " & rs_Estructura2!tenro
        If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
        OpenRecordset StrSql, rs_Filtro
        If Not rs_Filtro.EOF Then
            'tiene filtro
            StrSql = "SELECT * FROM mod_lin_filtro "
            StrSql = StrSql & " WHERE masinro = " & Masinro
            StrSql = StrSql & " AND linaorden = " & LinaOrden
            StrSql = StrSql & " AND estrnro = " & rs_Estructura2!estrnro
            If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
            OpenRecordset StrSql, rs_Filtro
            
            If Not rs_Filtro.EOF Then
                If IsNumeric(rs_Estructura2!estrcodext) Then
                    'Estructura2 = Format(rs_Estructura2!estrcodext, "00000000000000000000")
                    Estructura2 = rs_Estructura2!estrcodext & "00000000000000000000"
                Else
                    Estructura2 = IIf(IsNull(rs_Estructura2!estrcodext), "00000000000000000000", rs_Estructura2!estrcodext & "00000000000000000000")
                End If
            Else
                Estructura2 = "00000000000000000000"
                Genera = False
            End If
        Else
            'no tiene filtro
            If IsNumeric(rs_Estructura2!estrcodext) Then
                Estructura2 = rs_Estructura2!estrcodext & "00000000000000000000"
            Else
                Estructura2 = IIf(IsNull(rs_Estructura2!estrcodext), "00000000000000000000", rs_Estructura2!estrcodext & "00000000000000000000")
            End If
        End If
    Else
        Estructura2 = "00000000000000000000"
    End If
End If
Estructura2 = Left(Estructura2, 20)

If IsNull(vec_testr3(Inx_3).Estructura) Or vec_testr3(Inx_3).Estructura = 0 Then
    Estructura3 = "00000000000000000000"
Else
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & vec_testr3(Inx_3).Estructura
    OpenRecordset StrSql, rs_Estructura3
    
    If Not rs_Estructura3.EOF Then
        'reviso que tenga un filtro
        StrSql = "SELECT * FROM mod_lin_filtro "
        StrSql = StrSql & " WHERE masinro = " & Masinro
        StrSql = StrSql & " AND linaorden = " & LinaOrden
        StrSql = StrSql & " AND tenro = " & rs_Estructura3!tenro
        If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
        OpenRecordset StrSql, rs_Filtro
        If Not rs_Filtro.EOF Then
            'tiene filtro
            StrSql = "SELECT * FROM mod_lin_filtro "
            StrSql = StrSql & " WHERE masinro = " & Masinro
            StrSql = StrSql & " AND linaorden = " & LinaOrden
            StrSql = StrSql & " AND estrnro = " & rs_Estructura3!estrnro
            If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
            OpenRecordset StrSql, rs_Filtro
            
            If Not rs_Filtro.EOF Then
                If IsNumeric(rs_Estructura3!estrcodext) Then
                    Estructura3 = rs_Estructura3!estrcodext & "00000000000000000000"
                Else
                    Estructura3 = IIf(IsNull(rs_Estructura3!estrcodext), "00000000000000000000", rs_Estructura3!estrcodext & "00000000000000000000")
                End If
            Else
                Estructura3 = "00000000000000000000"
                Genera = False
            End If
        Else
            'no tiene filtro
            If IsNumeric(rs_Estructura3!estrcodext) Then
                Estructura3 = rs_Estructura3!estrcodext & "00000000000000000000"
            Else
                Estructura3 = IIf(IsNull(rs_Estructura3!estrcodext), "00000000000000000000", rs_Estructura3!estrcodext & "00000000000000000000")
            End If
        End If
    Else
        Estructura3 = "00000000000000000000"
    End If
End If
Estructura3 = Left(Estructura3, 20)

'Estructura1 = vec_testr1(inx)
'Estructura2 = vec_testr2(inx)
'Estructura3 = vec_testr3(inx)

PosE1 = 1
PosE2 = 1
PosE3 = 1

If Genera Then
    'Voy recorriendo de Izquierda a Derecha el aux_cuenta y voy generando el NroCuenta
    i = 1
    NroCuenta = ""
    CantL = 0
    CantE = 0
    Termino = False
    Do While Not (i > Len(Aux_Cuenta))
        ch = UCase(Mid(Aux_Cuenta, i, 1))
    
        Select Case ch
        Case "_", "-", ".":
            NroCuenta = NroCuenta & ch
            i = i + 1
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            NroCuenta = NroCuenta & ch
            i = i + 1
        Case "E": 'Estrcutura
            EsEstructura = True
            CantE = 1
            'leo el nro de la estructura
            i = i + 1
            ch = UCase(Mid(Aux_Cuenta, i, 1))
            TipoE = ch
            Termino = False
            
            Do While EsEstructura And Not Termino
                'leo el siguiente
                i = i + 1
                If Not (i > Len(Aux_Cuenta)) Then
                    ch = UCase(Mid(Aux_Cuenta, i, 1))
                Else
                    Termino = True
                End If
                
                If ch = "E" And Not Termino Then
                    'leo lel nro de la estructura
                    i = i + 1
                    ch = UCase(Mid(Aux_Cuenta, i, 1))
                    TipoE_Actual = ch
                    
                    Do While TipoE = TipoE_Actual And EsEstructura And Not Termino
                        CantE = CantE + 1
        
                        i = i + 1
                        If Not (i > Len(Aux_Cuenta)) Then
                            ch = UCase(Mid(Aux_Cuenta, i, 1))
                        Else
                            Termino = True
                        End If
                        
                        If ch = "E" Then
                            'leo el nro de la estructura
                            i = i + 1
                            ch = UCase(Mid(Aux_Cuenta, i, 1))
                            TipoE_Actual = ch
                        Else
                            Termino = True
                        End If
                    Loop
                    
                    
                Else
                    EsEstructura = False
                End If
                
                'reemplazo por la estructura correspondiente
                Select Case TipoE
                Case 1:
                    'NroCuenta = NroCuenta & Right(Estructura1, CantE)
                    NroCuenta = NroCuenta & Mid(Estructura1, PosE1, CantE)
                    PosE1 = PosE1 + CantE
                    If PosE1 >= 20 Then PosE1 = 1
                Case 2:
                    'NroCuenta = NroCuenta & Right(Estructura2, CantE)
                    NroCuenta = NroCuenta & Mid(Estructura2, PosE2, CantE)
                    PosE2 = PosE2 + CantE
                    If PosE2 >= 20 Then PosE2 = 1
                Case 3:
                    'NroCuenta = NroCuenta & Right(Estructura3, CantE)
                    NroCuenta = NroCuenta & Mid(Estructura3, PosE3, CantE)
                    PosE3 = PosE3 + CantE
                    If PosE3 >= 20 Then PosE3 = 1
                End Select
                
                TipoE = TipoE_Actual
                CantE = 1
            Loop
            
        Case "L": 'Legajo
            Termino = False
            CantL = 1
            i = i + 1
            If i <= Len(Aux_Cuenta) Then
                ch = UCase(Mid(Aux_Cuenta, i, 1))
            End If
            
            Do While ch = "L" And Not Termino
                CantL = CantL + 1
                i = i + 1
                If i <= Len(Aux_Cuenta) Then
                    ch = UCase(Mid(Aux_Cuenta, i, 1))
                Else
                    Termino = True
                End If
            Loop
            
            NroCuenta = NroCuenta & Right(Format(Aux_Legajo, "0000000000"), CantL)
        Case "a" To "z", "A" To "Z":
            NroCuenta = NroCuenta & ch
            i = i + 1
        Case Else:
            i = i + 1
        End Select
    Loop
End If

'cierro todo
If rs_Estructura1.State = adStateOpen Then rs_Estructura1.Close
If rs_Estructura2.State = adStateOpen Then rs_Estructura2.Close
If rs_Estructura3.State = adStateOpen Then rs_Estructura3.Close

Set rs_Estructura1 = Nothing
Set rs_Estructura2 = Nothing
Set rs_Estructura3 = Nothing

End Sub


Public Sub ArmarCuenta_Con_Apertura(ByVal testr1 As Long, ByVal testr2 As Long, ByVal testr3 As Long, ByRef NroCuenta As String, ByVal Masinro As Long, ByVal LinaOrden As Long, ByRef Genera As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion   :
' Autor         : FGZ
' Fecha         : 01/09/2004
' Modificacion  :
' --------------------------------------------------------------------------------------------
Dim Aux_Cuenta As String
Dim Aux_Legajo As String
Dim Estructura1 As String
Dim Estructura2 As String
Dim Estructura3 As String

Dim i As Integer
Dim ch As String
Dim CantL As Integer
Dim CantE As Integer
Dim TipoE As String
Dim TipoE_Actual As String
Dim EsEstructura As Boolean
Dim Termino As Boolean

Dim PosE1 As Integer
Dim PosE2 As Integer
Dim PosE3 As Integer

Dim rs_Estructura1 As New ADODB.Recordset
Dim rs_Estructura2 As New ADODB.Recordset
Dim rs_Estructura3 As New ADODB.Recordset
Dim rs_Filtro As New ADODB.Recordset

Aux_Cuenta = NroCuenta
Aux_Legajo = rs_Empleado!empleg
Genera = True

If IsNull(testr1) Or testr1 = 0 Then
    Estructura1 = "00000000000000000000"
Else
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & testr1
    OpenRecordset StrSql, rs_Estructura1
    
    If Not rs_Estructura1.EOF Then
        'reviso que tenga un filtro
        StrSql = "SELECT * FROM mod_lin_filtro "
        StrSql = StrSql & " WHERE masinro = " & Masinro
        StrSql = StrSql & " AND linaorden = " & LinaOrden
        StrSql = StrSql & " AND tenro = " & rs_Estructura1!tenro
        If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
        OpenRecordset StrSql, rs_Filtro
        If Not rs_Filtro.EOF Then
            'tiene filtro
            StrSql = "SELECT * FROM mod_lin_filtro "
            StrSql = StrSql & " WHERE masinro = " & Masinro
            StrSql = StrSql & " AND linaorden = " & LinaOrden
            StrSql = StrSql & " AND estrnro = " & rs_Estructura1!estrnro
            If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
            OpenRecordset StrSql, rs_Filtro
            
            If Not rs_Filtro.EOF Then
                If IsNumeric(rs_Estructura1!estrcodext) Then
                    Estructura1 = rs_Estructura1!estrcodext & "00000000000000000000"
                Else
                    Estructura1 = IIf(IsNull(rs_Estructura1!estrcodext), "00000000000000000000", rs_Estructura1!estrcodext & "00000000000000000000")
                End If
            Else
                Estructura1 = "00000000000000000000"
                Genera = False
            End If
        Else
            'no tiene filtro
            If IsNumeric(rs_Estructura1!estrcodext) Then
                Estructura1 = rs_Estructura1!estrcodext & "00000000000000000000"
            Else
                Estructura1 = IIf(IsNull(rs_Estructura1!estrcodext), "00000000000000000000", rs_Estructura1!estrcodext & "00000000000000000000")
            End If
        End If
    Else
        Estructura1 = "00000000000000000000"
    End If
End If
Estructura1 = Left(Estructura1, 20)

If IsNull(testr2) Or testr2 = 0 Then
    Estructura2 = "00000000000000000000"
Else
    StrSql = " SELECT estrcodext,estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & testr2
    OpenRecordset StrSql, rs_Estructura2
    
    If Not rs_Estructura2.EOF Then
        'reviso que tenga un filtro
        StrSql = "SELECT * FROM mod_lin_filtro "
        StrSql = StrSql & " WHERE masinro = " & Masinro
        StrSql = StrSql & " AND linaorden = " & LinaOrden
        StrSql = StrSql & " AND tenro = " & rs_Estructura2!tenro
        If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
        OpenRecordset StrSql, rs_Filtro
        If Not rs_Filtro.EOF Then
            'tiene filtro
            StrSql = "SELECT * FROM mod_lin_filtro "
            StrSql = StrSql & " WHERE masinro = " & Masinro
            StrSql = StrSql & " AND linaorden = " & LinaOrden
            StrSql = StrSql & " AND estrnro = " & rs_Estructura2!estrnro
            If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
            OpenRecordset StrSql, rs_Filtro
            
            If Not rs_Filtro.EOF Then
                If IsNumeric(rs_Estructura2!estrcodext) Then
                    'Estructura2 = Format(rs_Estructura2!estrcodext, "00000000000000000000")
                    Estructura2 = rs_Estructura2!estrcodext & "00000000000000000000"
                Else
                    Estructura2 = IIf(IsNull(rs_Estructura2!estrcodext), "00000000000000000000", rs_Estructura2!estrcodext & "00000000000000000000")
                End If
            Else
                Estructura2 = "00000000000000000000"
                Genera = False
            End If
        Else
            'no tiene filtro
            If IsNumeric(rs_Estructura2!estrcodext) Then
                Estructura2 = rs_Estructura2!estrcodext & "00000000000000000000"
            Else
                Estructura2 = IIf(IsNull(rs_Estructura2!estrcodext), "00000000000000000000", rs_Estructura2!estrcodext & "00000000000000000000")
            End If
        End If
    Else
        Estructura2 = "00000000000000000000"
    End If
End If
Estructura2 = Left(Estructura2, 20)

If IsNull(testr3) Or testr3 = 0 Then
    Estructura3 = "00000000000000000000"
Else
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & testr3
    OpenRecordset StrSql, rs_Estructura3
    
    If Not rs_Estructura3.EOF Then
        'reviso que tenga un filtro
        StrSql = "SELECT * FROM mod_lin_filtro "
        StrSql = StrSql & " WHERE masinro = " & Masinro
        StrSql = StrSql & " AND linaorden = " & LinaOrden
        StrSql = StrSql & " AND tenro = " & rs_Estructura3!tenro
        If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
        OpenRecordset StrSql, rs_Filtro
        If Not rs_Filtro.EOF Then
            'tiene filtro
            StrSql = "SELECT * FROM mod_lin_filtro "
            StrSql = StrSql & " WHERE masinro = " & Masinro
            StrSql = StrSql & " AND linaorden = " & LinaOrden
            StrSql = StrSql & " AND estrnro = " & rs_Estructura3!estrnro
            If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
            OpenRecordset StrSql, rs_Filtro
            
            If Not rs_Filtro.EOF Then
                If IsNumeric(rs_Estructura3!estrcodext) Then
                    Estructura3 = rs_Estructura3!estrcodext & "00000000000000000000"
                Else
                    Estructura3 = IIf(IsNull(rs_Estructura3!estrcodext), "00000000000000000000", rs_Estructura3!estrcodext & "00000000000000000000")
                End If
            Else
                Estructura3 = "00000000000000000000"
                Genera = False
            End If
        Else
            'no tiene filtro
            If IsNumeric(rs_Estructura3!estrcodext) Then
                Estructura3 = rs_Estructura3!estrcodext & "00000000000000000000"
            Else
                Estructura3 = IIf(IsNull(rs_Estructura3!estrcodext), "00000000000000000000", rs_Estructura3!estrcodext & "00000000000000000000")
            End If
        End If
    Else
        Estructura3 = "00000000000000000000"
    End If
End If
Estructura3 = Left(Estructura3, 20)

PosE1 = 1
PosE2 = 1
PosE3 = 1

If Genera Then
    'Voy recorriendo de Izquierda a Derecha el aux_cuenta y voy generando el NroCuenta
    i = 1
    NroCuenta = ""
    CantL = 0
    CantE = 0
    Termino = False
    Do While Not (i > Len(Aux_Cuenta))
        ch = UCase(Mid(Aux_Cuenta, i, 1))
    
        Select Case ch
        Case "_", "-", ".":
            NroCuenta = NroCuenta & ch
            i = i + 1
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            NroCuenta = NroCuenta & ch
            i = i + 1
        Case "E": 'Estrcutura
            EsEstructura = True
            CantE = 1
            'leo el nro de la estructura
            i = i + 1
            ch = UCase(Mid(Aux_Cuenta, i, 1))
            TipoE = ch
            Termino = False
            
            Do While EsEstructura And Not Termino
                'leo el siguiente
                i = i + 1
                If Not (i > Len(Aux_Cuenta)) Then
                    ch = UCase(Mid(Aux_Cuenta, i, 1))
                Else
                    Termino = True
                End If
                
                If ch = "E" And Not Termino Then
                    'leo lel nro de la estructura
                    i = i + 1
                    ch = UCase(Mid(Aux_Cuenta, i, 1))
                    TipoE_Actual = ch
                    
                    Do While TipoE = TipoE_Actual And EsEstructura And Not Termino
                        CantE = CantE + 1
        
                        i = i + 1
                        If Not (i > Len(Aux_Cuenta)) Then
                            ch = UCase(Mid(Aux_Cuenta, i, 1))
                        Else
                            Termino = True
                        End If
                        
                        If ch = "E" Then
                            'leo el nro de la estructura
                            i = i + 1
                            ch = UCase(Mid(Aux_Cuenta, i, 1))
                            TipoE_Actual = ch
                        Else
                            Termino = True
                        End If
                    Loop
                    
                    
                Else
                    EsEstructura = False
                End If
                
                'reemplazo por la estructura correspondiente
                Select Case TipoE
                Case 1:
                    'NroCuenta = NroCuenta & Right(Estructura1, CantE)
                    NroCuenta = NroCuenta & Mid(Estructura1, PosE1, CantE)
                    PosE1 = PosE1 + CantE
                    If PosE1 >= 20 Then PosE1 = 1
                Case 2:
                    'NroCuenta = NroCuenta & Right(Estructura2, CantE)
                    NroCuenta = NroCuenta & Mid(Estructura2, PosE2, CantE)
                    PosE2 = PosE2 + CantE
                    If PosE2 >= 20 Then PosE2 = 1
                Case 3:
                    'NroCuenta = NroCuenta & Right(Estructura3, CantE)
                    NroCuenta = NroCuenta & Mid(Estructura3, PosE3, CantE)
                    PosE3 = PosE3 + CantE
                    If PosE3 >= 20 Then PosE3 = 1
                End Select
                
                TipoE = TipoE_Actual
                CantE = 1
            Loop
            
        Case "L": 'Legajo
            Termino = False
            CantL = 1
            i = i + 1
            If i <= Len(Aux_Cuenta) Then
                ch = UCase(Mid(Aux_Cuenta, i, 1))
            End If
            
            Do While ch = "L" And Not Termino
                CantL = CantL + 1
                i = i + 1
                If i <= Len(Aux_Cuenta) Then
                    ch = UCase(Mid(Aux_Cuenta, i, 1))
                Else
                    Termino = True
                End If
            Loop
            NroCuenta = NroCuenta & Right(Format(Aux_Legajo, "0000000000"), CantL)
        Case "a" To "z", "A" To "Z":
            NroCuenta = NroCuenta & ch
            i = i + 1
        Case Else:
            i = i + 1
        End Select
    Loop
End If

'cierro todo
If rs_Estructura1.State = adStateOpen Then rs_Estructura1.Close
If rs_Estructura2.State = adStateOpen Then rs_Estructura2.Close
If rs_Estructura3.State = adStateOpen Then rs_Estructura3.Close

Set rs_Estructura1 = Nothing
Set rs_Estructura2 = Nothing
Set rs_Estructura3 = Nothing

End Sub




Public Sub Movilidad(ByVal Ternro As Long, ByVal Asi_Cod As Long, ByVal FechaDesde As Date, ByVal FechaHasta As Date, ByVal Masinro As Long, ByVal Vol_Fec_Asiento As Date, ByVal TipoE1 As Long, ByVal TipoE2 As Long, ByVal TipoE3 As Long, ByVal Masinivternro1 As Long, ByVal Masinivternro2 As Long, ByVal Masinivternro3 As Long, ByRef NroAsientos As Long, ByRef NroLineas As Long, ByRef Abortado As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion: Imputacion por empleados
' Autor      : FGZ
' Fecha      : 01/04/2005
' --------------------------------------------------------------------------------------------
Dim rs_Imputacion As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Mod_Linea As New ADODB.Recordset
Dim rs_ee_Mov As New ADODB.Recordset
Dim rs_LineaContrapartida As New ADODB.Recordset

Dim Tot_Jor  As Single
Dim monto_a_imputar As Single
Dim HayImputaciones As Boolean
Dim distri_legajo As Boolean
Dim Ya_Imputo As Boolean

Dim HayImputaciones1 As Boolean
Dim HayImputaciones2 As Boolean
Dim HayImputaciones3 As Boolean

    'Distribucion en % Fijos para cada empleado
    StrSql = "SELECT * FROM imputacion "
    StrSql = StrSql & " WHERE imputacion.ternro = " & Ternro
    StrSql = StrSql & " AND imputacion.masinro = " & Masinro
    StrSql = StrSql & " ORDER BY imputacion.tenro"
    OpenRecordset StrSql, rs_Imputacion
    If rs_Imputacion.EOF Then
        Flog.writeline "No hay imputaciones para el legajo"
        HayImputaciones = False
    Else
        HayImputaciones = True
    End If
    HayImputaciones1 = False
    HayImputaciones2 = False
    HayImputaciones3 = False
    
    Inx = 1
    LI_1 = 48
    LI_2 = 48
    LI_3 = 48
    Do While Not rs_Imputacion.EOF '(2)
        If rs_Imputacion!tenro = Masinivternro1 Then
            If LI_1 > Inx Then
                'LI_1 = Inx - 1
                LI_1 = Inx
            End If
            vec_testr1(Inx).TE = Masinivternro1
            vec_testr1(Inx).Estructura = rs_Imputacion!estrnro
            vec_testr1(Inx).Porcentaje = vec_testr1(Inx).Porcentaje + rs_Imputacion!Porcentaje
            HayImputaciones1 = True
        End If
        If rs_Imputacion!tenro = Masinivternro2 Then
            If LI_2 > Inx Then
                'LI_2 = Inx - 1
                LI_2 = Inx
            End If
            vec_testr2(Inx).TE = Masinivternro2
            vec_testr2(Inx).Estructura = rs_Imputacion!estrnro
            vec_testr2(Inx).Porcentaje = vec_testr2(Inx).Porcentaje + rs_Imputacion!Porcentaje
            HayImputaciones2 = True
        End If
        If rs_Imputacion!tenro = Masinivternro3 Then
            If LI_3 > Inx Then
                'LI_3 = Inx - 1
                LI_3 = Inx
            End If
            vec_testr3(Inx).TE = Masinivternro3
            vec_testr3(Inx).Estructura = rs_Imputacion!estrnro
            vec_testr3(Inx).Porcentaje = vec_testr3(Inx).Porcentaje + rs_Imputacion!Porcentaje
            HayImputaciones3 = True
        End If

        Inx = Inx + 1

        rs_Imputacion.MoveNext
    Loop '(2)
    'FIN Distribucion en % Fijos para cada empleado
    Inxfin = Inx - 1
    
    
    Flog.writeline "DISTRIBUCION EN BASE AL LEGAJO DEL EMPLEADO"
    If Not HayImputaciones1 Then
        If Not IsNull(Masinivternro1) Then
            LI_1 = 1
            Flog.writeline "Tipo de Estructura de nivel 1: " & Masinivternro1
            StrSql = " SELECT estrnro FROM his_estructura " & _
                     " WHERE ternro = " & Ternro & " AND " & _
                     " tenro =" & Masinivternro1 & " AND " & _
                     " (htetdesde <= " & ConvFecha(Vol_Fec_Asiento) & ") AND " & _
                     " ((" & ConvFecha(Vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
    
            If Not rs_Estructura.EOF Then
                rs_Estructura.MoveLast
                vec_testr1(1).TE = Masinivternro1
                vec_testr1(1).Estructura = rs_Estructura!estrnro
                vec_testr1(1).Porcentaje = 100
                Flog.writeline "Estructura " & vec_testr1(1).Estructura
            Else
                Flog.writeline "No se encontró ningun tipo de Estructura de nivel 1"
            End If
        Else
            Flog.writeline "Tipo de Estructura de nivel 1 Nulo"
        End If
        If Inxfin < 1 Then
            Inxfin = 1
        End If
    End If
    
    If Not HayImputaciones2 Then
        If Not IsNull(Masinivternro2) Then
            LI_2 = 1
            Flog.writeline "Tipo de Estructura de nivel 2: " & Masinivternro2
            StrSql = " SELECT estrnro FROM his_estructura " & _
                     " WHERE ternro = " & Ternro & " AND " & _
                     " tenro =" & Masinivternro2 & " AND " & _
                     " (htetdesde <= " & ConvFecha(Vol_Fec_Asiento) & ") AND " & _
                     " ((" & ConvFecha(Vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
    
            If Not rs_Estructura.EOF Then
                rs_Estructura.MoveLast
                vec_testr2(1).TE = Masinivternro2
                vec_testr2(1).Estructura = rs_Estructura!estrnro
                vec_testr2(1).Porcentaje = 100
                Flog.writeline "Estructura " & vec_testr2(1).Estructura
            Else
                Flog.writeline "No se encontró ningun tipo de Estructura de nivel 2"
            End If
        Else
            Flog.writeline "Tipo de Estructura de nivel 2 Nulo"
        End If
        If Inxfin < 1 Then
            Inxfin = 1
        End If
    End If
    
    If Not HayImputaciones3 Then
        If Not IsNull(Masinivternro3) Then
            LI_3 = 1
            Flog.writeline "Tipo de Estructura de nivel 3: " & Masinivternro3
            StrSql = " SELECT estrnro FROM his_estructura " & _
                     " WHERE ternro = " & Ternro & " AND " & _
                     " tenro =" & Masinivternro3 & " AND " & _
                     " (htetdesde <= " & ConvFecha(Vol_Fec_Asiento) & ") AND " & _
                     " ((" & ConvFecha(Vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
    
            If Not rs_Estructura.EOF Then
                rs_Estructura.MoveLast
                vec_testr3(1).TE = Masinivternro3
                vec_testr3(1).Estructura = rs_Estructura!estrnro
                vec_testr3(1).Porcentaje = 100
                Flog.writeline "Estructura " & vec_testr3(1).Estructura
            Else
                Flog.writeline "No se encontró ningun tipo de Estructura de nivel 3"
            End If
        Else
            Flog.writeline "Tipo de Estructura de nivel 3 Nulo"
        End If
        If Inxfin < 1 Then
            Inxfin = 1
        End If
    End If
    
    'Busco la mascara de la cuenta de contrapartida
    'Se supone que hay una sola cuenta del haber que es la cuenta de contrapartida
    'por lo que busco la primer cuenta que encuentra del haber y tomo esa como la cuenta de contrapartida
    
    'linaD_H = 0    'Debe
    'linaD_H = 1    'Haber
    'linaD_H = 2    'Variable
    StrSql = "SELECT * FROM mod_linea "
    StrSql = StrSql & " WHERE masinro = " & Asi_Cod
    StrSql = StrSql & " AND linaD_H = 1"    'del haber
    If rs_LineaContrapartida.State = adStateOpen Then rs_LineaContrapartida.Close
    OpenRecordset StrSql, rs_LineaContrapartida
    If Not rs_LineaContrapartida.EOF Then
        Mascara = rs_LineaContrapartida!linacuenta
        Orden_CuentaContrapartida = rs_LineaContrapartida!LinaOrden
        Descripcion_CuentaContrapartida = rs_LineaContrapartida!linadesc
    Else
        Mascara = "Inexistente"
    End If
   
    
    distri_legajo = True
    Tot_Jor = 100
    NroLineas = 0
    StrSql = "SELECT * FROM mod_linea " & _
             " WHERE masinro = " & Asi_Cod
    OpenRecordset StrSql, rs_Mod_Linea
    Do While Not rs_Mod_Linea.EOF '(2)
        NroLineas = NroLineas + 1
        'Tabla de movilidad
        Flog.writeline "ee_Movilidad "
        StrSql = "SELECT * FROM ee_movilidad "
        StrSql = StrSql & " WHERE masinro = " & rs_Mod_Linea!Masinro
        StrSql = StrSql & " AND linaorden = " & rs_Mod_Linea!LinaOrden
        StrSql = StrSql & " AND ternro = " & Ternro
        StrSql = StrSql & " AND (fecha73 >= " & ConvFecha(FechaDesde)
        StrSql = StrSql & " AND fecha73 <= " & ConvFecha(FechaHasta) & ")"
        OpenRecordset StrSql, rs_ee_Mov

        Do While Not rs_ee_Mov.EOF '(4)
            'Ciclar por los tres niveles de estructura (las que haya)
            Inx_1 = LI_1
            Inx_2 = LI_2
            Inx_3 = LI_3
            
            If Not EsNulo(rs_Mod_Linea!lineanivternro1) Or Not EsNulo(rs_Mod_Linea!lineanivternro2) Or Not EsNulo(rs_Mod_Linea!lineanivternro3) Then
                For Inx = 1 To Inxfin
                    Ya_Imputo = False
                    If vec_testr1(Inx).Estructura <> 0 And vec_testr1(Inx).TE = rs_Mod_Linea!lineanivternro1 Then
                        Inx_1 = Inx
                        'Inx_2 = IIf(Inx + LI_2 <= UBound(vec_testr2), Inx + LI_2, 0)
                        'Inx_3 = IIf(Inx + LI_3 <= UBound(vec_testr3), Inx + LI_3, 0)
                        monto_a_imputar = rs_ee_Mov!imptot73 * vec_testr1(Inx).Porcentaje / Tot_Jor
                        Descripcion = CStr(rs_ee_Mov!movnro) + " - " + rs_ee_Mov!vehinro
                        Cantidad = rs_ee_Mov!dias73 * vec_testr1(Inx).Porcentaje / Tot_Jor
    
                        'acumular en la linea del asiento
                        '{conta\acu_hya.i monto_a_imputar FALSE asi_acu.signo}
                        Flog.writeline "Monto a imputar " & monto_a_imputar
                        Flog.writeline "Cantidad a imputar " & Cantidad
                        Call Acu_tmp(monto_a_imputar, False, 1, rs_Mod_Linea!linadesc, rs_Mod_Linea!linacuenta, rs_Mod_Linea!Masinro, rs_Mod_Linea!LinaOrden, Descripcion, vec_testr1(Inx).Porcentaje, rs_ee_Mov!movnro, 0)
                        Ya_Imputo = True
                    End If
                    If Not Ya_Imputo Then
                        If vec_testr2(Inx).Estructura <> 0 And vec_testr2(Inx).TE = rs_Mod_Linea!lineanivternro2 Then
                            'Inx_1 = IIf(Inx + LI_1 <= UBound(vec_testr1), Inx + LI_1, 0)
                            Inx_2 = Inx
                            'Inx_3 = IIf(Inx + LI_3 <= UBound(vec_testr3), Inx + LI_3, 0)
                            monto_a_imputar = rs_ee_Mov!imptot73 * vec_testr2(Inx).Porcentaje / Tot_Jor
                            Descripcion = CStr(rs_ee_Mov!movnro) + " - " + rs_ee_Mov!vehinro
                            Cantidad = rs_ee_Mov!dias73 * vec_testr2(Inx).Porcentaje / Tot_Jor
        
                            'acumular en la linea del asiento
                            '{conta\acu_hya.i monto_a_imputar FALSE asi_acu.signo}
                            Flog.writeline "Monto a imputar " & monto_a_imputar
                            Flog.writeline "Cantidad a imputar " & Cantidad
                            Call Acu_tmp(monto_a_imputar, False, 1, rs_Mod_Linea!linadesc, rs_Mod_Linea!linacuenta, rs_Mod_Linea!Masinro, rs_Mod_Linea!LinaOrden, Descripcion, vec_testr2(Inx).Porcentaje, rs_ee_Mov!movnro, 0)
                            Ya_Imputo = True
                        End If
                        If Not Ya_Imputo Then
                            If vec_testr3(Inx).Estructura <> 0 And vec_testr3(Inx).TE = rs_Mod_Linea!lineanivternro3 Then
                                'Inx_1 = IIf(Inx + LI_1 <= UBound(vec_testr1), Inx + LI_1, 0)
                                'Inx_2 = IIf(Inx + LI_2 <= UBound(vec_testr2), Inx + LI_2, 0)
                                Inx_3 = Inx
                                monto_a_imputar = rs_ee_Mov!imptot73 * vec_testr3(Inx).Porcentaje / Tot_Jor
                                Descripcion = CStr(rs_ee_Mov!movnro) + " - " + rs_ee_Mov!vehinro
                                Cantidad = rs_ee_Mov!dias73 * vec_testr3(Inx).Porcentaje / Tot_Jor
            
                                'acumular en la linea del asiento
                                '{conta\acu_hya.i monto_a_imputar FALSE asi_acu.signo}
                                Flog.writeline "Monto a imputar " & monto_a_imputar
                                Flog.writeline "Cantidad a imputar " & Cantidad
                                Call Acu_tmp(monto_a_imputar, False, 1, rs_Mod_Linea!linadesc, rs_Mod_Linea!linacuenta, rs_Mod_Linea!Masinro, rs_Mod_Linea!LinaOrden, Descripcion, vec_testr3(Inx).Porcentaje, rs_ee_Mov!movnro, 0)
                                Ya_Imputo = True
                            End If
                        End If
                    End If
                Next Inx
            Else
                monto_a_imputar = rs_ee_Mov!imptot73
                Descripcion = CStr(rs_ee_Mov!movnro) + " - " + rs_ee_Mov!vehinro
                Cantidad = rs_ee_Mov!dias73

                Flog.writeline "Monto a imputar " & monto_a_imputar
                Flog.writeline "Cantidad a imputar " & Cantidad
                Call Acu_tmp(monto_a_imputar, False, 1, rs_Mod_Linea!linadesc, rs_Mod_Linea!linacuenta, rs_Mod_Linea!Masinro, rs_Mod_Linea!LinaOrden, Descripcion, 100, rs_ee_Mov!movnro, 0)
            End If
            rs_ee_Mov.MoveNext
        Loop '(4)

        
        rs_Mod_Linea.MoveNext
    Loop '(2)

'cierro y libero
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Mod_Linea.State = adStateOpen Then rs_Mod_Linea.Close
If rs_Imputacion.State = adStateOpen Then rs_Imputacion.Close
If rs_ee_Mov.State = adStateOpen Then rs_ee_Mov.Close
If rs_LineaContrapartida.State = adStateOpen Then rs_LineaContrapartida.Close

Set rs_Estructura = Nothing
Set rs_Mod_Linea = Nothing
Set rs_Imputacion = Nothing
Set rs_ee_Mov = Nothing
Set rs_LineaContrapartida = Nothing
End Sub



Public Sub Recalcular_lineas(ByVal Masinro As Long, ByVal Vol_Cod As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Recalcula las lineas de asiento teniendo en cuenta el valor promedio de cada cuenta.
' Autor      : FGZ
' Fecha      : 02/05/2005
' --------------------------------------------------------------------------------------------
Dim rs_Linea_asi As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim CuentaAnterior As String

Dim MontoTotal As Single
Dim CantidadTotal As Long
Dim CantidadReal As Long
Dim Promedio As Single

    On Error GoTo ME_Local
    
    MontoTotal = 0
    CantidadTotal = 1
    CantidadReal = 0
    Promedio = 0
    CuentaAnterior = ""
    
    StrSql = "SELECT * FROM linea_asi "
    StrSql = StrSql & " WHERE linea_asi.vol_cod =" & Vol_Cod
    StrSql = StrSql & " AND linea_asi.masinro =" & Masinro
    StrSql = StrSql & " ORDER BY desclinea "
    OpenRecordset StrSql, rs_Linea_asi
     
    Do While Not rs_Linea_asi.EOF
        If CuentaAnterior <> rs_Linea_asi!desclinea Then
            CuentaAnterior = rs_Linea_asi!desclinea
            'Busco el monto acumulado de la cuenta sin aperturas
            StrSql = "SELECT sum(dlmonto) as MontoTotal FROM detalle_asi "
            StrSql = StrSql & " WHERE masinro = " & Masinro
            StrSql = StrSql & " AND Vol_Cod = " & Vol_Cod
            StrSql = StrSql & " AND dldescripcion = '" & rs_Linea_asi!desclinea & "'"
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                MontoTotal = rs!MontoTotal
            Else
                MontoTotal = 0
            End If
            
            'Busco la cantidad de legajos que suman a la cuenta sin aperturas
            StrSql = "SELECT count(distinct empleg) as Cantidad FROM detalle_asi "
            StrSql = StrSql & " WHERE masinro = " & Masinro
            StrSql = StrSql & " AND Vol_Cod = " & Vol_Cod
            StrSql = StrSql & " AND dldescripcion = '" & rs_Linea_asi!desclinea & "'"
            If rs.State = adStateOpen Then rs.Close
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                CantidadTotal = rs!Cantidad
            Else
                CantidadTotal = 1
            End If
            
            'Promedio
            Promedio = MontoTotal / CantidadTotal
        End If
        
        'Busco la cantidad de legajos que suman a la misma cuenta
        StrSql = "SELECT count(distinct empleg) as Cantidad FROM detalle_asi "
        StrSql = StrSql & " WHERE masinro = " & Masinro
        StrSql = StrSql & " AND Vol_Cod = " & Vol_Cod
        StrSql = StrSql & " AND cuenta = '" & rs_Linea_asi!cuenta & "'"
        If rs.State = adStateOpen Then rs.Close
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            CantidadReal = rs!Cantidad
        Else
            CantidadReal = 1
        End If
        
        'Actualizo los valores del linea_asi
        StrSql = "UPDATE linea_asi SET "
        StrSql = StrSql & " monto = " & CantidadReal * Promedio
        StrSql = StrSql & " WHERE vol_cod = " & Vol_Cod
        StrSql = StrSql & " AND masinro = " & Masinro
        StrSql = StrSql & " AND cuenta = '" & rs_Linea_asi!cuenta & "'"
        objConn.Execute StrSql, , adExecuteNoRecords
    
        rs_Linea_asi.MoveNext
    Loop

'cierro y libero
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

If rs_Linea_asi.State = adStateOpen Then rs_Linea_asi.Close
Set rs_Linea_asi = Nothing
Exit Sub

ME_Local:
    Flog.writeline "Error: " & Err.Description
End Sub


Function Redon(Valor)

   Redon = Round(Valor, 2)

End Function

