Attribute VB_Name = "MdlAsientoContable"
Option Explicit

'Const Version = 1.01
'Const FechaVersion = "15/08/2005"   'Martin Ferraro - Version Inicial

'Const Version = 1.02
'Const FechaVersion = "22/02/2007"   'Martin Ferraro - Permitir * en los nro de cuenta

'Const Version = 1.03
'Const FechaVersion = "12/05/2007"   'Martin Ferraro - Guardar la descripcion de la linea en el detalle y debe/haber
                                    'para la exportacion alcoa
'Const Version = 1.04
'Const FechaVersion = "11/05/2007"   'Martin Ferraro - En GuardarDetalleAsi se resetean los indicies al final

'Const Version = 1.05
'Const FechaVersion = "13/06/2007 - Custom SMT - Modelo TARJA"   'Fernando Favre - modelo TARJA custom para SMT

'Const Version = 1.06
'Const FechaVersion = "10/07/2007"  'Fernando Favre - Se agrego la proporcionalidad en las cantidades

Const Version = 1.07
Const FechaVersion = "12/08/2014"  'Fernando Favre - Se agrego la proporcionalidad en las cantidades
'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------

Global rs_Empleado As New ADODB.Recordset
Global rs_Mod_Asiento As New ADODB.Recordset

Global CantidadEmpleados As Long
Global CatidadVueltas As Long
Global corteDesbalance As Boolean
Global NroVol As Long
Global vol_Fec_Asiento As Date

'Vector de imputaciones -------------------------------------
Private Type TregImputacion
    te1 As Long
    Estructura1 As Long
    te2 As Long
    Estructura2 As Long
    te3 As Long
    Estructura3 As Long
    porcentaje As Double
End Type
Global vec_imputacion(100) As TregImputacion
Global ind_imp_act As Long
Const max_ind_imp = 100

'Tabla aux de linea asi -------------------------------------
Private Type TregLineaAsiAux
    cuenta As String
    Linea As Long
    desclinea As String
    dh As Integer
    Monto As Double
End Type
Global lineaAsiAux(1000) As TregLineaAsiAux
Global ind_lineaAsiAux As Long
Const max_ind_lineaAsiAux = 1000

'Tabla aux de detalle asi por linea-------------------------------------
Private Type TregDetalleAsiAux
    masinro As Long
    vol_cod As Long
    cuenta As String
    lin_orden As Long
    empleg As Long
    terape As String
    dldescripcion As String
    dlcantidad As Double
    dlmonto As Double
    porcentaje As Double
    Ternro As Long
    Origen As Long
    tipoorigen As Long
    moddesc As String
    linadesc As String
    linaD_H As Integer
End Type
Global detalleAsiAux(1000) As TregDetalleAsiAux
Global ind_detalleAsiAux As Long
Const max_ind_detalleAsiAux = 1000

'Tabla aux de detalle asi por empleado-------------------------------------
Private Type TregDetalleAsiAuxEmp
    masinro As Long
    vol_cod As Long
    cuenta As String
    lin_orden As Long
    empleg As Long
    terape As String
    dldescripcion As String
    dlcantidad As Double
    dlmonto As Double
    dlmontoacum As Double
    dlcosto1 As Long
    dlcosto2 As Long
    dlcosto3 As Long
    dlcosto4 As Long
    dlporcentaje As Double
    Ternro As Long
    Origen As Long
    tipoorigen As Long
    moddesc As String
    linadesc As String
    linaD_H As Integer
End Type
Global detalleAsiAuxEmp(1000) As TregDetalleAsiAuxEmp
Global ind_detalleAsiAuxEmp As Long
Const max_ind_detalleAsiAuxEmp = 1000

'Vector de concepto de Tarja --------------------------------
Private Type TvecConcepto
    ConcNro As Integer
    cuenta As String
    proyecto As String
    canthoras As Double
End Type
Global vec_con(100) As TvecConcepto
Global ind_con_act As Integer
Const max_ind_con = 100

Private Type TvecTotConcepto
    ConcNro As Integer
    canthoras As Double
End Type
Global vec_con_tot(100) As TvecTotConcepto
Global ind_con_tot_act As Integer

Global vec_jor(100) As Double
Global vec_cta(100) As String
Global vec_pro(100) As String

Global errorCorte As Boolean
Global primer_asi_cod As Integer
Global tot_jor As Double
Global asi_cod_ant As Integer
Global ternro_ant As Long



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Asientos Contables.
' Autor      : Martin Ferraro
' Fecha      : 07/08/2006
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
    
    
    Nombre_Arch = PathFLog & "Asiento_Contable" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline

    TiempoInicialProceso = GetTickCount
    
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
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "'"
    StrSql = StrSql & " , bprcfecinicioej = " & ConvFecha(Date)
    StrSql = StrSql & " , bprcestado = 'Procesando'"
    StrSql = StrSql & " , bprcpid = " & PID
    StrSql = StrSql & " , bprctiempo = 0 "
    StrSql = StrSql & " , bprcprogreso = 0 "
    StrSql = StrSql & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 6 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call GenerarAsiento(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "NO SE ENCONTRO EL PROCESO " & NroProcesoBatch
    End If
    
    TiempoFinalProceso = GetTickCount
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = 100, bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords


Fin:
    Flog.Close
    If objConn.State = adStateOpen Then objConn.Close
    If objconnProgreso.State = adStateOpen Then objconnProgreso.Close
Exit Sub

ME_Main:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " Error : " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
End Sub



Public Sub GenerarAsiento(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Programa que se ejecuta para generar Asiento Contable
'              Configurado en el tipo de proceso batch
' Autor      : Martin Ferraro
' Fecha      : 24/12/2006
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer

Dim total As Long
Dim NroAsientos As Long
Dim NroLineas As Long
Dim NroAsi As Long
Dim NroLin As Long

Dim rs_ProcVol As New ADODB.Recordset
Dim rs_Proc_V_modasi As New ADODB.Recordset
Dim rs_Aux As New ADODB.Recordset


On Error GoTo ME_GenerarAsiento

'Seteo la varible global de corte
errorCorte = False

' Levanto cada parametro por separado, el separador de parametros es "."
Flog.writeline Espacios(Tabulador * 0) & "Inicio del proceso de volcado."
If Not IsNull(Parametros) Then
    If Len(Parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        NroVol = CLng(Mid(Parametros, pos1, pos2))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        HACE_TRAZA = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        corteDesbalance = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
                
    End If
End If

Flog.writeline Espacios(Tabulador * 0) & "Parametros: "
Flog.writeline Espacios(Tabulador * 0) & "            Numero de Proceso = " & NroVol
Flog.writeline Espacios(Tabulador * 0) & "            Analisis Detallado = " & HACE_TRAZA
Flog.writeline Espacios(Tabulador * 0) & "            Corte Desbalance = " & corteDesbalance
Flog.writeline
Flog.writeline


Flog.writeline Espacios(Tabulador * 0) & "Buscando el proceso de volcado."
'Buscando el proceso de volcado
StrSql = "SELECT * FROM proc_vol WHERE proc_vol.vol_cod =" & NroVol
OpenRecordset StrSql, rs_ProcVol

If rs_ProcVol.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. Proceso de Volcado no encontrado."
    Exit Sub
End If

Flog.writeline Espacios(Tabulador * 0) & "Buscando modelos proceso de volcado."
'Buscando los modelos asociados al proceso
StrSql = "SELECT * FROM proc_v_modasi WHERE proc_v_modasi.vol_cod =" & NroVol
StrSql = StrSql & " ORDER BY asi_cod "
OpenRecordset StrSql, rs_Proc_V_modasi

If rs_Proc_V_modasi.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. El proceso de volcado no tiene modelos asociados."
    Exit Sub
End If

'Buscando empleados del proceso
Flog.writeline Espacios(Tabulador * 0) & "Buscando cabeceras a porcesar del proceso de volcado."
StrSql = "SELECT * FROM proc_vol_pl" & _
         " INNER JOIN proc_vol_emp ON proc_vol_emp.pronro  = proc_vol_pl.pronro" & _
         " WHERE proc_vol_pl.vol_cod =" & NroVol & _
         " AND proc_vol_emp.vol_cod = " & NroVol & _
         " ORDER BY proc_vol_emp.ternro"
OpenRecordset StrSql, rs_Empleado
If rs_Empleado.EOF Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. El proceso de volcado no tiene cabeceras asociados."
    Exit Sub
End If

'seteo las variables de progreso
Progreso = 0
CatidadVueltas = rs_Proc_V_modasi.RecordCount
CantidadEmpleados = rs_Empleado.RecordCount
IncPorc = 95 / (CatidadVueltas * CantidadEmpleados)

'variable global de fecha de asiento
vol_Fec_Asiento = rs_ProcVol!vol_Fec_Asiento


Flog.writeline Espacios(Tabulador * 0) & "Cantidad de modelos del proceso de volcado = " & CatidadVueltas
Flog.writeline Espacios(Tabulador * 0) & "Cantidad de cabeceras a procesar del proceso de volcado = " & CantidadEmpleados

'Por cada modelo del Proceso de volcado
Do While Not rs_Proc_V_modasi.EOF

    
    
    'Voy a la primera cabecera para procesar nuevamente todas en el siguiente modelo
    rs_Empleado.MoveFirst
    
    'Proceso todas las cabeceras
    Call ProcesarModelo(rs_Proc_V_modasi!asi_cod)
    
    'Paso al siguiente modelo
    rs_Proc_V_modasi.MoveNext
        
    'Verifico si debe cortar por error
    If corteDesbalance Then
        If errorCorte Then
            Flog.writeline "CORTE POR ERROR."
            Exit Sub
        End If
    End If
    
Loop

'Cuento la cantidad de lineas generadas
StrSql = "SELECT count(*) Lineas FROM linea_asi "
StrSql = StrSql & " WHERE vol_cod =" & NroVol
If rs_Aux.State = adStateOpen Then rs_Aux.Close
OpenRecordset StrSql, rs_Aux
If Not rs_Aux.EOF Then
    NroLin = rs_Aux!Lineas
End If

'Cuento la cantidad de asientos generados
StrSql = "SELECT COUNT(DISTINCT masinro) Asientos FROM linea_asi "
StrSql = StrSql & " WHERE vol_cod =" & NroVol
If rs_Aux.State = adStateOpen Then rs_Aux.Close
OpenRecordset StrSql, rs_Aux
If Not rs_Aux.EOF Then
    NroAsi = rs_Aux!Asientos
End If

StrSql = "UPDATE proc_vol SET vol_reg_cab = " & NroAsi & _
             ", vol_reg_det =" & NroLin & _
             " WHERE proc_vol.vol_cod =" & NroVol
objConn.Execute StrSql, , adExecuteNoRecords


If rs_ProcVol.State = adStateOpen Then rs_ProcVol.Close
If rs_Proc_V_modasi.State = adStateOpen Then rs_Proc_V_modasi.Close

Set rs_ProcVol = Nothing
Set rs_Proc_V_modasi = Nothing

Exit Sub

'Manejador de Errores del procedimiento
ME_GenerarAsiento:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: GenerarAsiento"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
End Sub


Public Sub BuscarConfModelo(ByVal nrocofc As Long, ByRef ProcesoGeneral As String, ByRef existe As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion: busca los datos de configuracion especificos del modelo
' Autor      : Martin Ferraro
' Fecha      : 20/07/2006
' --------------------------------------------------------------------------------------------
Dim rs_Conf_cont As New ADODB.Recordset

    existe = False
    ProcesoGeneral = ""
    
    StrSql = "SELECT * FROM conf_cont WHERE conf_cont.cofcnro =" & nrocofc
    OpenRecordset StrSql, rs_Conf_cont
        
    If rs_Conf_cont.EOF Then
        Flog.writeline "ERROR. No existe Proceso de Configuraci¢n asociado al Modelo de Asiento."
    Else
        If rs_Conf_cont!cofcacum = "" Then
            Flog.writeline "ERROR. Falta ingresar el Archivo de Acumulación."
        Else
            If EsNulo(rs_Conf_cont!cofcacum) Then
                Flog.writeline "ERROR. Falta ingresar el Archivo de Acumulación."
            Else
                ProcesoGeneral = UCase(rs_Conf_cont!cofcacum)
                existe = True
            End If
        End If
    End If

If rs_Conf_cont.State = adStateOpen Then rs_Conf_cont.Close
Set rs_Conf_cont = Nothing

End Sub


Public Sub ProcesarModelo(ByVal asi_cod As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: realiza el procesamiento de los empleados para un modelo
' Autor      : Martin Ferraro
' Fecha      : 27/07/2001
' --------------------------------------------------------------------------------------------

Dim rs_tercero As New ADODB.Recordset

Dim Masinivternro1 As Long
Dim Masinivternro2 As Long
Dim Masinivternro3 As Long
Dim ProcesoGeneral As String
Dim existeConf As Boolean
Dim balancea As Boolean

'Activo el manejador de errores local
On Error GoTo ME_ProcesarModelo

Flog.writeline
Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------"
Flog.writeline Espacios(Tabulador * 0) & "Procesando Modelo = " & asi_cod
Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------"

'Busco los datos del modelo
Flog.writeline Espacios(Tabulador * 0) & "Buscando datos del modelo "
StrSql = "SELECT * FROM mod_asiento where mod_asiento.masinro = " & asi_cod
OpenRecordset StrSql, rs_Mod_Asiento

If rs_Mod_Asiento.EOF Then
    'AVANZAR 1 PASO DE TODOS LOS EMPLEADOS
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontró Mod_Asiento."
    Exit Sub
End If

Flog.writeline Espacios(Tabulador * 0) & "Modelo = " & rs_Mod_Asiento!masidesc

'Niveles de apertura del modelo
Masinivternro1 = IIf(Not EsNulo(rs_Mod_Asiento!Masinivternro1), rs_Mod_Asiento!Masinivternro1, 0)
Masinivternro2 = IIf(Not EsNulo(rs_Mod_Asiento!Masinivternro2), rs_Mod_Asiento!Masinivternro2, 0)
Masinivternro3 = IIf(Not EsNulo(rs_Mod_Asiento!Masinivternro3), rs_Mod_Asiento!Masinivternro3, 0)
Flog.writeline Espacios(Tabulador * 0) & "Nivel 1 de estructura del modelo = " & Masinivternro1
Flog.writeline Espacios(Tabulador * 0) & "Nivel 2 de estructura del modelo = " & Masinivternro2
Flog.writeline Espacios(Tabulador * 0) & "Nivel 3 de estructura del modelo = " & Masinivternro3

'Busco la configuracion del modelo
Call BuscarConfModelo(rs_Mod_Asiento!cofcnro, ProcesoGeneral, existeConf)
      
If Not existeConf Then
    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No existe tipo de proceso del modelo = " & ProcesoGeneral
    Exit Sub
End If

Flog.writeline Espacios(Tabulador * 0) & "Tipo de proceso del modelo = " & ProcesoGeneral

'Limpia la tabla temporal de lineaAsi
Call InicializarVectorLineaAsiAux

'rs_empleado es global son todos los cabliq del proceso de volcado
Do While Not rs_Empleado.EOF

    Select Case ProcesoGeneral
        Case "ESTRUCTURAS":
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Proceso no desarrollado."
            Exit Sub
        Case "PORCENTAJES":
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Proceso no desarrollado."
            Exit Sub
        Case "ESTANDAR":
            Call AcumularEstandar(rs_Empleado!Ternro, rs_Empleado!cliqnro, rs_Empleado!pronro, rs_Mod_Asiento!masinro, Masinivternro1, Masinivternro2, Masinivternro3)
        Case "TARJA":
'            Call AcumularTarja(rs_Empleado!ternro, rs_Empleado!cliqnro, rs_Empleado!pronro, rs_Mod_Asiento!masinro, Masinivternro1, Masinivternro2, Masinivternro3)
            Call AcumularTarja(asi_cod, rs_Empleado!Ternro, rs_Empleado!cliqnro, rs_Empleado!pronro, rs_Mod_Asiento!masinro, Masinivternro1, Masinivternro2, Masinivternro3)
        Case "PROMEDIOS":
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Proceso no desarrollado."
            Exit Sub
        Case "GTI":
            Flog.writeline Espacios(Tabulador * 0) & "ERROR. Proceso no desarrollado."
            Exit Sub
        Case Else
            Exit Sub
            Flog.writeline "PROCESO DESCONOCIDO"
    End Select
    
    
    'Control balance de empleado
    balancea = True
    Call BalanceEmpleado(balancea)
    
    If Not balancea Then
        'Seteo la variable de desbalance
        errorCorte = True
    End If
    
    'Guardo linea_asi en tabla
    Call GuardarLineaAsi(NroVol, asi_cod)
    
    'Verifico si debe cortar por error
    If corteDesbalance Then
        If errorCorte Then
            Exit Sub
        End If
    End If
    
    rs_Empleado.MoveNext
    
    'Actualizar el progreso
    TiempoFinalProceso = GetTickCount
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    
Loop

'Balance
Call BalanceModelo(NroVol, asi_cod)

StrSql = "DELETE linea_asi WHERE linea_asi.monto = 0"
objConn.Execute StrSql, , adExecuteNoRecords


Exit Sub
'Manejador de Errores del procedimiento
ME_ProcesarModelo:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: ProcesarModelo"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
End Sub

Public Sub AcumularEstandar(ByVal Ternro As Long, ByVal cliqnro As Long, ByVal pronro As Long, ByVal masinro As Long, ByVal Masinivternro1 As Long, ByVal Masinivternro2 As Long, ByVal Masinivternro3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: realiza la acumulacion para el empleado, cabliq, pronro
' Autor      : Martin Ferraro
' Fecha      : 24/07/2006
' --------------------------------------------------------------------------------------------

Dim HayImputaciones As Boolean

Dim rs_tercero As New ADODB.Recordset
Dim rs_Imputacion As New ADODB.Recordset
Dim rs_Mod_Linea As New ADODB.Recordset
Dim rs_Asi_monto As New ADODB.Recordset
Dim rs_Asi_Acu_Con As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim monto_linea As Double
Dim monto_aux As Double
Dim signo As String
Dim estr1 As Long
Dim estr2 As Long
Dim estr3 As Long
Dim HayMasinivternro As Boolean
Dim HayMasinivternro1 As Boolean
Dim HayMasinivternro2 As Boolean
Dim HayMasinivternro3 As Boolean
Dim porcentaje As Double
Dim imputaTenro1 As Long
Dim imputaTenro2 As Long
Dim imputaTenro3 As Long
Dim imputaEstrnro1 As Long
Dim imputaEstrnro2 As Long
Dim imputaEstrnro3 As Long
Dim indice As Long
Dim Generar As Boolean
Dim cuenta As String
Dim MontoAImputar As Double
Dim generoAlguna As Boolean
Dim MontoRedondeo As Double

On Error GoTo ME_Acumular
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Acumulando para ternro = " & Ternro & " cliqnro = " & cliqnro & " pronro = " & pronro
    
    
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    'Busco que sea un empleado valido
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    StrSql = "SELECT * FROM empleado where empleado.ternro = " & Ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el legajo"
        Exit Sub
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Emplado : " & rs_tercero!empleg & " - " & rs_tercero!terape & ", " & rs_tercero!ternom
    End If
    
    
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    'Verifico que el empleado pertenezca a los tipos de estructuras del modelo
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    HayMasinivternro = False
    HayMasinivternro1 = False
    HayMasinivternro2 = False
    HayMasinivternro3 = False
    
    If Masinivternro1 <> 0 Then
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & Ternro & " AND " & _
                 " tenro =" & Masinivternro1 & " AND " & _
                 " (htetdesde <= " & ConvFecha(vol_Fec_Asiento) & ") AND " & _
                 " ((" & ConvFecha(vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            estr1 = rs_Estructura!Estrnro
            HayMasinivternro1 = True
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El empleado NO pertenece al PRIMER nivel de estructura del modelo a la fecha " & vol_Fec_Asiento
            'errorCorte = True
            Exit Sub
        End If
        rs_Estructura.Close
    End If
    
    If Masinivternro2 <> 0 Then
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & Ternro & " AND " & _
                 " tenro =" & Masinivternro2 & " AND " & _
                 " (htetdesde <= " & ConvFecha(vol_Fec_Asiento) & ") AND " & _
                 " ((" & ConvFecha(vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            estr2 = rs_Estructura!Estrnro
            HayMasinivternro2 = True
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El empleado NO pertenece al SEGUNDO nivel de estructura del modelo a la fecha " & vol_Fec_Asiento
            'errorCorte = True
            Exit Sub
        End If
        rs_Estructura.Close
    End If
    
    If Masinivternro3 <> 0 Then
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & Ternro & " AND " & _
                 " tenro =" & Masinivternro3 & " AND " & _
                 " (htetdesde <= " & ConvFecha(vol_Fec_Asiento) & ") AND " & _
                 " ((" & ConvFecha(vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            estr3 = rs_Estructura!Estrnro
            HayMasinivternro3 = True
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El empleado NO pertenece al TERCER nivel de estructura del modelo a la fecha " & vol_Fec_Asiento
            'errorCorte = True
            Exit Sub
        End If
        rs_Estructura.Close
    End If
    
    HayMasinivternro = HayMasinivternro1 Or HayMasinivternro2 Or HayMasinivternro3
    
    
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    'Armar el vector de imputacion
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    If Not HayMasinivternro Then
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene apertura"
    Else
        'El modelo tiene apertura
        Flog.writeline Espacios(Tabulador * 1) & "El modelo tiene apertura"
        
        'Borro el vector de imputacion
        Call BorrarVectorImputacion
        ind_imp_act = 0
        
        Flog.writeline Espacios(Tabulador * 1) & "Busqueda de Distribuion contable"
        
        'Distribucion en % Fijos para cada empleado
        StrSql = "SELECT * FROM imputacion where imputacion.ternro = " & Ternro & _
                 " AND imputacion.masinro = " & masinro & _
                 " AND imputacion.porcentaje <> 0 " & _
                 " ORDER BY imputacion.impnro"
        OpenRecordset StrSql, rs_Imputacion
    
        If Not rs_Imputacion.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "El empleado tiene Distribucion Contable"
            
            'ARMO EL VECTOR DE IMPUTACIONES EN BASE A LO CARGADO DESDE ADP
            porcentaje = 0
            Do While Not rs_Imputacion.EOF
                
                ind_imp_act = ind_imp_act + 1
                
                'Controlo desbordamiento
                If ind_imp_act > max_ind_imp Then
                    Flog.writeline Espacios(Tabulador * 1) & "Error. El indice del vector de imputaciones supero el max de " & max_ind_imp
                End If
                
                imputaTenro1 = IIf(EsNulo(rs_Imputacion!Tenro), 0, rs_Imputacion!Tenro)
                imputaTenro2 = IIf(EsNulo(rs_Imputacion!tenro2), 0, rs_Imputacion!tenro2)
                imputaTenro3 = IIf(EsNulo(rs_Imputacion!tenro3), 0, rs_Imputacion!tenro3)
                imputaEstrnro1 = IIf(EsNulo(rs_Imputacion!Estrnro), 0, rs_Imputacion!Estrnro)
                imputaEstrnro2 = IIf(EsNulo(rs_Imputacion!estrnro2), 0, rs_Imputacion!estrnro2)
                imputaEstrnro3 = IIf(EsNulo(rs_Imputacion!estrnro3), 0, rs_Imputacion!estrnro3)
                
                'Miro que componente tiene cargada
                
                'Si el modelo tiene apertura por tipo estructura 1
                If (Masinivternro1 <> 0) Then
                   'cargo el tipo de estructura (debe coincidir con la del modelo)
                    vec_imputacion(ind_imp_act).te1 = Masinivternro1
                    If imputaEstrnro1 <> 0 Then
                        'cargo con la estructura de la imputacion
                        vec_imputacion(ind_imp_act).Estructura1 = imputaEstrnro1
                    Else
                        'cargo con la estructura del empleado
                        vec_imputacion(ind_imp_act).Estructura1 = estr1
                    End If
                End If
                
                'Si el modelo tiene apertura por tipo estructura 2
                If (Masinivternro2 <> 0) Then
                    'cargo el tipo de estructura (debe coincidir con la del modelo)
                    vec_imputacion(ind_imp_act).te2 = Masinivternro2
                    If imputaEstrnro2 <> 0 Then
                        'cargo con la estructura de la imputacion
                        vec_imputacion(ind_imp_act).Estructura2 = imputaEstrnro2
                    Else
                        'cargo con la estructura del empleado
                        vec_imputacion(ind_imp_act).Estructura2 = estr2
                    End If
                End If
                
                'Si el modelo tiene apertura por tipo estructura 3
                If (Masinivternro3 <> 0) Then
                    'cargo el tipo de estructura (debe coincidir con la del modelo)
                    vec_imputacion(ind_imp_act).te3 = Masinivternro3
                    If imputaEstrnro3 <> 0 Then
                        'cargo con la estructura de la imputacion
                        vec_imputacion(ind_imp_act).Estructura3 = imputaEstrnro3
                    Else
                        'cargo con la estructura del empleado
                        vec_imputacion(ind_imp_act).Estructura3 = estr3
                    End If
                End If
                
                'Cargo el porcentaje
                vec_imputacion(ind_imp_act).porcentaje = rs_Imputacion!porcentaje
                
                Flog.writeline Espacios(Tabulador * 1) & "Vector de Imputacion( " & ind_imp_act & ") TipoEst1 = " & vec_imputacion(ind_imp_act).te1 & " Estr1 = " & vec_imputacion(ind_imp_act).Estructura1 & " TipoEst2 = " & vec_imputacion(ind_imp_act).te2 & " Estr2 = " & vec_imputacion(ind_imp_act).Estructura2 & " TipoEst3 = " & vec_imputacion(ind_imp_act).te3 & " Estr3 = " & vec_imputacion(ind_imp_act).Estructura3 & " Porcentaje = " & vec_imputacion(ind_imp_act).porcentaje
                
                porcentaje = porcentaje + rs_Imputacion!porcentaje
                
                rs_Imputacion.MoveNext
            Loop
            rs_Imputacion.Close
            
            'Si el porcentaje es < 100 debo completar
            If porcentaje < 100 Then
                'Si el porcentaje es menor o igual que 1 a la ultima imputacion la corrijo
                If CDbl(100 - porcentaje) <= 1 Then
                    'A la ultima imputacion la completo con lo faltante
                    vec_imputacion(ind_imp_act).porcentaje = vec_imputacion(ind_imp_act).porcentaje + (100 - porcentaje)
                    Flog.writeline Espacios(Tabulador * 1) & "Correccion de la componente " & ind_imp_act & " por error de redondeo."
                    Flog.writeline Espacios(Tabulador * 1) & "Vector de Imputacion( " & ind_imp_act & ") TipoEst1 = " & vec_imputacion(ind_imp_act).te1 & " Estr1 = " & vec_imputacion(ind_imp_act).Estructura1 & " TipoEst2 = " & vec_imputacion(ind_imp_act).te2 & " Estr2 = " & vec_imputacion(ind_imp_act).Estructura2 & " TipoEst3 = " & vec_imputacion(ind_imp_act).te3 & " Estr3 = " & vec_imputacion(ind_imp_act).Estructura3 & " Porcentaje = " & vec_imputacion(ind_imp_act).porcentaje
                Else
                    'sino inserto otra componente con el % faltante con la estructura del empleado
                    
                    ind_imp_act = ind_imp_act + 1
                    
                    If Masinivternro1 <> 0 Then
                        vec_imputacion(ind_imp_act).te1 = Masinivternro1
                        vec_imputacion(ind_imp_act).Estructura1 = estr1
                    End If
                    If Masinivternro2 <> 0 Then
                        vec_imputacion(ind_imp_act).te2 = Masinivternro2
                        vec_imputacion(ind_imp_act).Estructura2 = estr2
                    End If
                    If Masinivternro3 <> 0 Then
                        vec_imputacion(ind_imp_act).te3 = Masinivternro3
                        vec_imputacion(ind_imp_act).Estructura3 = estr3
                    End If
                    
                    vec_imputacion(ind_imp_act).porcentaje = (100 - porcentaje)
                    
                    Flog.writeline Espacios(Tabulador * 1) & "El % no es 100, completo con las estructuras del empleado."
                    Flog.writeline Espacios(Tabulador * 1) & "Vector de Imputacion( " & ind_imp_act & ") TipoEst1 = " & vec_imputacion(ind_imp_act).te1 & " Estr1 = " & vec_imputacion(ind_imp_act).Estructura1 & " TipoEst2 = " & vec_imputacion(ind_imp_act).te2 & " Estr2 = " & vec_imputacion(ind_imp_act).Estructura2 & " TipoEst3 = " & vec_imputacion(ind_imp_act).te3 & " Estr3 = " & vec_imputacion(ind_imp_act).Estructura3 & " Porcentaje = " & vec_imputacion(ind_imp_act).porcentaje
                End If
                
            End If
            
        Else
            rs_Imputacion.Close
            Flog.writeline Espacios(Tabulador * 1) & "El empleado NO posee Distribucion Contable"
            'Armo el vector de imputaciones con la distribucion del empleado al 100%
            
            ind_imp_act = ind_imp_act + 1
            
            If Masinivternro1 <> 0 Then
                vec_imputacion(ind_imp_act).te1 = Masinivternro1
                vec_imputacion(ind_imp_act).Estructura1 = estr1
            End If
            If Masinivternro2 <> 0 Then
                vec_imputacion(ind_imp_act).te2 = Masinivternro2
                vec_imputacion(ind_imp_act).Estructura2 = estr2
            End If
            If Masinivternro3 <> 0 Then
                vec_imputacion(ind_imp_act).te3 = Masinivternro3
                vec_imputacion(ind_imp_act).Estructura3 = estr3
            End If
            
            vec_imputacion(ind_imp_act).porcentaje = 100
            
            Flog.writeline Espacios(Tabulador * 1) & "Vector de Imputacion( " & ind_imp_act & ") TipoEst1 = " & vec_imputacion(ind_imp_act).te1 & " Estr1 = " & vec_imputacion(ind_imp_act).Estructura1 & " TipoEst2 = " & vec_imputacion(ind_imp_act).te2 & " Estr2 = " & vec_imputacion(ind_imp_act).Estructura2 & " TipoEst3 = " & vec_imputacion(ind_imp_act).te3 & " Estr3 = " & vec_imputacion(ind_imp_act).Estructura3 & " Porcentaje = " & vec_imputacion(ind_imp_act).porcentaje
            
        End If

    End If 'Si el modelo tiene distribucion contable
    
    'BORRO EL VECTOR QUE ACUMULA DETALLES DEL EMPLEADO
    If HACE_TRAZA Then
        Call BorrarDetalleAsiAuxEmp
    End If
    
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    'Calculo de las lineas del modelo
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    StrSql = "SELECT * FROM mod_linea WHERE masinro = " & masinro
    OpenRecordset StrSql, rs_Mod_Linea
    Do While Not rs_Mod_Linea.EOF
        
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 2) & "Procesando linea: " & rs_Mod_Linea!LinaOrden & " - " & rs_Mod_Linea!linadesc & " Cuenta: " & rs_Mod_Linea!linacuenta
        
        'Verifico que la cuenta no sea niveladora
        If UCase(rs_Mod_Linea!linadesc) = "NIVELADORA" Then
            'Cuenta Niveladora
            Flog.writeline Espacios(Tabulador * 3) & "Cuenta Niveladora. No se realiza acumulacion de la misma."
        Else
            'Analizo la cuenta
            
            'SI HACE TRAZA BORRO EL VECTOR QUE ACUMULA DETALLES DE EMPLEADO Y CUENTA
            If HACE_TRAZA Then
                Call BorrarDetalleAsiAux
            End If
            
            'Inicializo el monto de la linea
            monto_linea = 0
            
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            'BUSQUEDA DE ACUMULADORES QUE CONTRIBUYEN EN LA LINEA
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            StrSql = "SELECT * FROM asi_acu " & _
                     " WHERE asi_acu.masinro = " & rs_Mod_Linea!masinro & _
                     " AND asi_acu.linaorden =" & rs_Mod_Linea!LinaOrden
            OpenRecordset StrSql, rs_Asi_Acu_Con
    
            Do While Not rs_Asi_Acu_Con.EOF
            
                StrSql = "SELECT * FROM acu_liq " & _
                         " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro " & _
                         " WHERE acu_liq.cliqnro = " & cliqnro & _
                         " AND acu_liq.acunro =" & rs_Asi_Acu_Con!acuNro
                OpenRecordset StrSql, rs_Asi_monto
                
                If Not rs_Asi_monto.EOF Then
                    
                    monto_aux = IIf(EsNulo(rs_Asi_monto!almonto), 0, rs_Asi_monto!almonto)
                    signo = "(+/-)"
                    'Si signo + o - entonces tomar valor absoluto
                    If rs_Asi_Acu_Con!signo <> 3 Then
                        monto_aux = Abs(monto_aux)
                        signo = "(+)"
                        'Si signo - entonces lo hago restar
                        If rs_Asi_Acu_Con!signo = 2 Then
                            monto_aux = -monto_aux
                            signo = "(-)"
                        End If
                    End If
                    
                    Flog.writeline Espacios(Tabulador * 3) & "ACUMULADOR " & rs_Asi_monto!acuNro & " " & rs_Asi_monto!acudesabr & " - MONTO = " & rs_Asi_monto!almonto & " - SIGNO = " & signo
                    monto_linea = monto_linea + monto_aux
                    
                    'GUARDO LOS DETALLES DEL ACUMULADOR QUE CONTRIBUYE EN EL VECTOR DE DETALLE POR LINEA EMPLEADO
                    If HACE_TRAZA Then
                        Call InsertarDetalleAsiAux(rs_Mod_Linea!masinro, NroVol, rs_Mod_Linea!linacuenta, rs_Mod_Linea!LinaOrden, rs_tercero!empleg, rs_tercero!terape & " " & rs_tercero!ternom, rs_Asi_monto!acuNro & "-" & rs_Asi_monto!acudesabr, rs_Asi_monto!alcant, monto_aux, rs_tercero!Ternro, rs_Asi_monto!acuNro, 2, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H)
                    End If
                                    
                End If 'rs_Asi_monto
                rs_Asi_monto.Close
                    
                rs_Asi_Acu_Con.MoveNext
            Loop
            
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            'BUSQUEDA DE CONCEPTOS QUE CONTRIBUYEN EN LA LINEA
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            StrSql = "SELECT * FROM asi_con " & _
                     " WHERE asi_con.masinro = " & rs_Mod_Linea!masinro & _
                     " AND asi_con.linaorden =" & rs_Mod_Linea!LinaOrden
            OpenRecordset StrSql, rs_Asi_Acu_Con
    
            Do While Not rs_Asi_Acu_Con.EOF
            
                StrSql = "SELECT * FROM detliq " & _
                         " INNER JOIN concepto ON concepto.concnro = detliq.concnro " & _
                         " WHERE detliq.cliqnro = " & cliqnro & _
                         " AND detliq.concnro =" & rs_Asi_Acu_Con!ConcNro
                OpenRecordset StrSql, rs_Asi_monto
                
                If Not rs_Asi_monto.EOF Then
                    
                    monto_aux = IIf(EsNulo(rs_Asi_monto!dlimonto), 0, rs_Asi_monto!dlimonto)
                    signo = "(+/-)"
                    'Si signo + o - entonces tomar valor absoluto
                    If rs_Asi_Acu_Con!signo <> 3 Then
                        monto_aux = Abs(monto_aux)
                        signo = "(+)"
                        'Si signo - entonces lo hago restar
                        If rs_Asi_Acu_Con!signo = 2 Then
                            monto_aux = -monto_aux
                            signo = "(-)"
                        End If
                    End If
                    
                    Flog.writeline Espacios(Tabulador * 3) & "CONCEPTO " & rs_Asi_monto!ConcCod & " " & rs_Asi_monto!concabr & " - MONTO = " & rs_Asi_monto!dlimonto & " - SIGNO = " & signo
                    monto_linea = monto_linea + monto_aux
                    
                    'GUARDO LOS DETALLES DEL CONCEPTO QUE CONTRIBUYE EN EL VECTOR DE DETALLE POR LINEA Y EMPLEADO
                    If HACE_TRAZA Then
                        Call InsertarDetalleAsiAux(rs_Mod_Linea!masinro, NroVol, rs_Mod_Linea!linacuenta, rs_Mod_Linea!LinaOrden, rs_tercero!empleg, rs_tercero!terape & " " & rs_tercero!ternom, rs_Asi_monto!ConcCod & "-" & rs_Asi_monto!concabr, rs_Asi_monto!dlicant, monto_aux, rs_tercero!Ternro, rs_Asi_monto!ConcNro, 1, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H)
                    End If
                                    
                End If 'rs_Asi_monto
                rs_Asi_monto.Close
                    
                rs_Asi_Acu_Con.MoveNext
            Loop
            
            Flog.writeline Espacios(Tabulador * 2) & "MONTO LINEA: " & Round(monto_linea, 4)
            
            
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            'Insercion en la linea
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            'monto_linea = Round(monto_linea, 4)
            
            If Not HayMasinivternro Then
                'Si el modelo no tiene distribucion, no tengo vector de imputacion, el 100% de monto_linea va a la linea
                cuenta = rs_Mod_Linea!linacuenta
                Call ArmarCuenta(cuenta, rs_tercero!Ternro, rs_tercero!empleg, 0, 0, 0, 0, 0, 0)
                Flog.writeline Espacios(Tabulador * 3) & "ARMADO DE CUENTA: " & rs_Mod_Linea!linacuenta & " ----------> " & cuenta
                Call InsertarVectorLineaAsiAux(cuenta, rs_Mod_Linea!LinaOrden, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H, monto_linea)
                
                'SI HACE TRAZA ENTONCES RESOLVER DATOS FALTANTES
                If HACE_TRAZA Then
                    Call ResolverDetalleAsi(NroVol, masinro, cuenta, 100, 0, 0, 0)
                End If
                
            Else
                'Debo distribuir de acuerdo al vector de distribucion
                Flog.writeline Espacios(Tabulador * 2) & "Distribucion del monto de la linea por el vector de imputacion."

                'Para ver si la suma de los valores parciales de las lineas es igual al monto total de la linea
                'Sino corrijo por redondeo
                MontoRedondeo = 0
                generoAlguna = False
                
                For indice = 1 To ind_imp_act
                
                    'Calculo el monto a imputar de acuerdo al vector
                    'MontoAImputar = Round((monto_linea * vec_imputacion(indice).porcentaje / 100), 4)
                    MontoAImputar = (monto_linea * vec_imputacion(indice).porcentaje / 100)
                    Flog.writeline Espacios(Tabulador * 3) & vec_imputacion(indice).porcentaje & " % del monto de la linea = " & MontoAImputar
                    
                    Flog.writeline Espacios(Tabulador * 3) & "Aplicando los Filtros de la linea de orden " & rs_Mod_Linea!LinaOrden & " Para la componente " & indice & " del vector de imputacion."
                    Call FiltrosLinea(masinro, rs_Mod_Linea!LinaOrden, Masinivternro1, Masinivternro2, Masinivternro3, vec_imputacion(indice).Estructura1, vec_imputacion(indice).Estructura2, vec_imputacion(indice).Estructura3, Generar)
                    If Generar Then
                        
                        Flog.writeline Espacios(Tabulador * 3) & "Filtro OK "
                        generoAlguna = True
                        cuenta = rs_Mod_Linea!linacuenta
                        Call ArmarCuenta(cuenta, rs_tercero!Ternro, rs_tercero!empleg, Masinivternro1, Masinivternro2, Masinivternro3, vec_imputacion(indice).Estructura1, vec_imputacion(indice).Estructura2, vec_imputacion(indice).Estructura3)
                        Flog.writeline Espacios(Tabulador * 3) & "ARMADO DE CUENTA: " & rs_Mod_Linea!linacuenta & " ----------> " & cuenta
                                    
                        Call InsertarVectorLineaAsiAux(cuenta, rs_Mod_Linea!LinaOrden, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H, MontoAImputar)
                        If HACE_TRAZA Then
                            Call ResolverDetalleAsi(NroVol, masinro, cuenta, 100, vec_imputacion(indice).Estructura1, vec_imputacion(indice).Estructura2, vec_imputacion(indice).Estructura3)
                        End If

                    End If
                    
                    'Acumulo en el redondeo
                    MontoRedondeo = MontoRedondeo + MontoAImputar
                    
                Next
                
                If generoAlguna Then
                    'REDONDEO
                    If Round(MontoRedondeo, 4) <> Round(monto_linea, 4) Then
                        Flog.writeline Espacios(Tabulador * 3) & "DIFERENCIA DE REDONDEO " & FormatNumber(MontoRedondeo - monto_linea, 100)
                    End If
                End If
                
            End If
            
        End If 'No es cuenta niveladora
            
        'Paso a la siguiente linea
        rs_Mod_Linea.MoveNext
        
    Loop
    
    
If rs_tercero.State = adStateOpen Then rs_tercero.Close
If rs_Imputacion.State = adStateOpen Then rs_Imputacion.Close
If rs_Mod_Linea.State = adStateOpen Then rs_Mod_Linea.Close
If rs_Asi_monto.State = adStateOpen Then rs_Asi_monto.Close
If rs_Asi_Acu_Con.State = adStateOpen Then rs_Asi_Acu_Con.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_tercero = Nothing
Set rs_Imputacion = Nothing
Set rs_Mod_Linea = Nothing
Set rs_Asi_monto = Nothing
Set rs_Asi_Acu_Con = Nothing
Set rs_Estructura = Nothing


Exit Sub
'Manejador de Errores del procedimiento
ME_Acumular:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: Acumular"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
End Sub


'Public Sub AcumularTarja(ByVal asi_cod As Integer, ByVal ternro As Long, ByVal cliqnro As Long, ByVal pronro As Long, ByVal masinro As Long, ByVal Masinivternro1 As Long, ByVal Masinivternro2 As Long, ByVal Masinivternro3 As Long)
Public Sub AcumularTarja(ByVal asi_cod As Integer, ByVal Ternro As Long, ByVal cliqnro As Long, ByVal pronro As Long, ByVal masinro As Long, ByVal Masinivternro1 As Long, ByVal Masinivternro2 As Long, ByVal Masinivternro3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: realiza la acumulacion para el empleado, cabliq, pronro
' Autor      : Fernando Favre
' Fecha      : 19/04/2007
' --------------------------------------------------------------------------------------------

Dim HayImputaciones As Boolean

Dim rs_tercero As New ADODB.Recordset
Dim rs_Imputacion As New ADODB.Recordset
Dim rs_Mod_Linea As New ADODB.Recordset
Dim rs_Asi_monto As New ADODB.Recordset
Dim rs_Asi_Acu_Con As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_gti_acumdiario As New ADODB.Recordset
Dim rs_tiph_con  As New ADODB.Recordset
Dim rs_detliq As New ADODB.Recordset
Dim rs_tiph_con2 As New ADODB.Recordset
Dim rs_batch_empleado As New ADODB.Recordset

Dim monto_linea As Double
Dim monto_aux As Double
Dim cant_aux As Double
Dim signo As String
Dim estr1 As Long
Dim estr2 As Long
Dim estr3 As Long
Dim HayMasinivternro As Boolean
Dim HayMasinivternro1 As Boolean
Dim HayMasinivternro2 As Boolean
Dim HayMasinivternro3 As Boolean
Dim porcentaje As Double
Dim imputaTenro1 As Long
Dim imputaTenro2 As Long
Dim imputaTenro3 As Long
Dim imputaEstrnro1 As Long
Dim imputaEstrnro2 As Long
Dim imputaEstrnro3 As Long
Dim indice As Long
Dim Generar As Boolean
Dim cuenta As String
Dim MontoAImputar As Double
Dim generoAlguna As Boolean
Dim MontoRedondeo As Double
Dim valor_jornal As Double

Dim val_jor As Double
Dim pliqdesde As Date
Dim pliqhasta As Date
Dim tiene_tarja As Boolean
Dim cargar  As Boolean
Dim inx As Integer
Dim inxfin As Integer
Dim adctaasiento_ant As String
Dim adproyecto_ant As String
Dim adctaasiento As String
Dim adproyecto As String

On Error GoTo ME_Acumular
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 1) & "Acumulando para ternro = " & Ternro & " cliqnro = " & cliqnro & " pronro = " & pronro
    
    
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    'Busco que sea un empleado valido
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    StrSql = "SELECT * FROM empleado where empleado.ternro = " & Ternro
    OpenRecordset StrSql, rs_tercero
    If rs_tercero.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "ERROR. No se encontro el legajo"
        Exit Sub
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Empleado : " & rs_tercero!empleg & " - " & rs_tercero!terape & ", " & rs_tercero!ternom
    End If
    
    
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    'Verifico que el empleado pertenezca a los tipos de estructuras del modelo
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    HayMasinivternro = False
    HayMasinivternro1 = False
    HayMasinivternro2 = False
    HayMasinivternro3 = False
    
    If Masinivternro1 <> 0 Then
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & Ternro & " AND " & _
                 " tenro =" & Masinivternro1 & " AND " & _
                 " (htetdesde <= " & ConvFecha(vol_Fec_Asiento) & ") AND " & _
                 " ((" & ConvFecha(vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            estr1 = rs_Estructura!Estrnro
            HayMasinivternro1 = True
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El empleado NO pertenece al PRIMER nivel de estructura del modelo a la fecha " & vol_Fec_Asiento
            'errorCorte = True
            Exit Sub
        End If
        rs_Estructura.Close
    End If
    
    If Masinivternro2 <> 0 Then
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & Ternro & " AND " & _
                 " tenro =" & Masinivternro2 & " AND " & _
                 " (htetdesde <= " & ConvFecha(vol_Fec_Asiento) & ") AND " & _
                 " ((" & ConvFecha(vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            estr2 = rs_Estructura!Estrnro
            HayMasinivternro2 = True
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El empleado NO pertenece al SEGUNDO nivel de estructura del modelo a la fecha " & vol_Fec_Asiento
            'errorCorte = True
            Exit Sub
        End If
        rs_Estructura.Close
    End If
    
    If Masinivternro3 <> 0 Then
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & Ternro & " AND " & _
                 " tenro =" & Masinivternro3 & " AND " & _
                 " (htetdesde <= " & ConvFecha(vol_Fec_Asiento) & ") AND " & _
                 " ((" & ConvFecha(vol_Fec_Asiento) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            estr3 = rs_Estructura!Estrnro
            HayMasinivternro3 = True
        Else
            Flog.writeline Espacios(Tabulador * 1) & "ERROR. El empleado NO pertenece al TERCER nivel de estructura del modelo a la fecha " & vol_Fec_Asiento
            'errorCorte = True
            Exit Sub
        End If
        rs_Estructura.Close
    End If
    
    HayMasinivternro = HayMasinivternro1 Or HayMasinivternro2 Or HayMasinivternro3
        
        
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    'Tarja. Custom en principio para SMT
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    ' Busco la fecha de Inicio y fin del periodo de Liq en cuestion
    inx = 0
    inxfin = 0

    If HACE_TRAZA Then
        ind_con_act = 0
        ind_con_tot_act = 0
        ' Blanquear el vector de conceptos
        For indice = 0 To max_ind_con
            vec_con(indice).cuenta = ""
            vec_con(indice).proyecto = ""
            vec_con(indice).canthoras = 0
            
            vec_con_tot(indice).ConcNro = 0
            vec_con_tot(indice).canthoras = 0
        Next
    End If
    
    If ternro_ant <> Ternro Or asi_cod_ant <> asi_cod Then
        tot_jor = 0
        For indice = 0 To max_ind_con
            vec_jor(indice) = 0
            vec_cta(indice) = ""
            vec_pro(indice) = ""
        Next
        asi_cod_ant = asi_cod
        ternro_ant = Ternro
    End If
    
    StrSql = " SELECT pliqdesde, pliqhasta FROM periodo "
    StrSql = StrSql & " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro "
    StrSql = StrSql & " WHERE proceso.pronro = " & pronro
    OpenRecordset StrSql, rs_Periodo
    If rs_Periodo.EOF Then
        Flog.writeline Espacios(Tabulador * 3) & "No se encontro el periodo asociado al proceso. SQL --> " & StrSql
    Else
        pliqdesde = rs_Periodo!pliqdesde
        pliqhasta = rs_Periodo!pliqhasta
    End If

    ' Busco si tiene horas de Targa
    StrSql = "SELECT * FROM gti_acumdiario INNER JOIN empleado ON gti_acumdiario.ternro = empleado.ternro "
    StrSql = StrSql & " WHERE gti_acumdiario.adfecha >= " & ConvFecha(pliqdesde) & " AND gti_acumdiario.adfecha <= " & ConvFecha(pliqhasta)
    StrSql = StrSql & " AND gti_acumdiario.adctaasiento <> '' AND gti_acumdiario.adctaasiento IS NOT NULL"
    StrSql = StrSql & " AND gti_acumdiario.adctaasiento <> '0' AND gti_acumdiario.adcanthoras > 0"
    StrSql = StrSql & " AND gti_acumdiario.ternro = " & Ternro
    StrSql = StrSql & " ORDER BY gti_acumdiario.adctaasiento, gti_acumdiario.adproyecto"
    OpenRecordset StrSql, rs_gti_acumdiario
    tiene_tarja = False
'    If Not rs_gti_acumdiario.EOF Then
'        StrSql = "SELECT * FROM batch_empleado WHERE bpronro = " & NroProcesoBatch & " AND ternro = " & ternro
'        StrSql = StrSql & " AND estado = '" & asi_cod & "'"
'        OpenRecordset StrSql, rs_batch_empleado
'        If rs_batch_empleado.EOF Then
'            StrSql = "INSERT INTO batch_empleado (bpronro,ternro,estado) VALUES (" & NroProcesoBatch & "," & ternro & ",'" & asi_cod & "')"
'            objConn.Execute StrSql, , adExecuteNoRecords
'        Else
'            If rs_batch_empleado!estado <> CStr(asi_cod) Then
'                GoTo ya_procesado
'            End If
'        End If
'    End If
    
    Do While Not rs_gti_acumdiario.EOF
        If Trim(rs_gti_acumdiario!adctaasiento) <> "0" And Trim(rs_gti_acumdiario!adctaasiento) <> "00000000000" Then
            ' Buscar la Valorizacion
            StrSql = " SELECT * FROM tiph_con WHERE tiph_con.thnro = " & rs_gti_acumdiario!thnro
            OpenRecordset StrSql, rs_tiph_con
            If rs_tiph_con.EOF Then
                Flog.writeline Espacios(Tabulador * 3) & "El tipo de Hora con codigo " & rs_gti_acumdiario!thnro & " no se relaciona con un concepto de Liquidación."
            Else
                ' Verificar que no exista mas de un registro en la relacion, sino todo se cae!!!
                StrSql = " SELECT * FROM concepto INNER JOIN tiph_con ON concepto.concnro = tiph_con.concnro "
                StrSql = StrSql & " INNER JOIN detliq ON detliq.concnro = tiph_con.concnro "
                StrSql = StrSql & " WHERE tiph_con.concnro = " & rs_tiph_con!ConcNro & " AND detliq.cliqnro = " & cliqnro
                OpenRecordset StrSql, rs_detliq
                If rs_detliq.EOF Then
                    Flog.writeline Espacios(Tabulador * 3) & "El tipo de Hora con codigo " & rs_tiph_con!thnro & " en el Acum. Diario No Liquidado."
                    valor_jornal = 0
                Else
                    If rs_detliq!dlicant <> 0 Then
                        valor_jornal = rs_detliq!dlimonto / rs_detliq!dlicant
                    Else
                        valor_jornal = rs_detliq!dlimonto
                    End If
                End If
    
                tot_jor = tot_jor + rs_gti_acumdiario!adcanthoras * valor_jornal
                vec_jor(inx) = vec_jor(inx) + rs_gti_acumdiario!adcanthoras * valor_jornal
    
                ' Arma un vector de concepto y carga la cantidad de horas.
                If rs_gti_acumdiario!adcanthoras * valor_jornal > 0 Then
                    StrSql = " SELECT * FROM tiph_con WHERE tiph_con.thnro = " & rs_gti_acumdiario!thnro
                    OpenRecordset StrSql, rs_tiph_con2
                    Do While Not rs_tiph_con2.EOF
                        adctaasiento = Trim(IIf(IsNull(rs_gti_acumdiario!adctaasiento), "", rs_gti_acumdiario!adctaasiento))
                        adproyecto = Trim(IIf(IsNull(rs_gti_acumdiario!adproyecto), "", rs_gti_acumdiario!adproyecto))
                        Call InsertarVectorConceptoTarja(rs_tiph_con2!ConcNro, adctaasiento, adproyecto, rs_gti_acumdiario!adcanthoras, valor_jornal)
                        rs_tiph_con2.MoveNext
                    Loop
                    rs_tiph_con2.Close
                End If
    
    '               Arma un vector de concepto y carga la cantidad de horas.
    '               vec_con.canthoras = vec_con.canthoras + (gti_acumdiario.adcanthoras * valor_jornal)
    '               Aca y en la parte de los conceptos de Recibo es donde se carga este vector
    '               0 Debe(+), 1 Haber(-), 2 Variable(+/-), 3 Variable invertido(-/+)
    '            end if
    
            End If
            rs_tiph_con.Close
    
            adctaasiento_ant = IIf(IsNull(rs_gti_acumdiario!adctaasiento), "", rs_gti_acumdiario!adctaasiento)
            adproyecto_ant = IIf(IsNull(rs_gti_acumdiario!adproyecto), "", rs_gti_acumdiario!adproyecto)
        
            rs_gti_acumdiario.MoveNext
    
            cargar = False
            If rs_gti_acumdiario.EOF Then
                If adctaasiento_ant <> "" Then
                    cargar = True
                End If
            Else
                If adctaasiento_ant <> rs_gti_acumdiario!adctaasiento Or adproyecto_ant <> rs_gti_acumdiario!adproyecto Then
                    cargar = True
                End If
            End If
    
            If cargar Then
                ' datos necesarios para cuentas del haber
                vec_cta(inx) = adctaasiento_ant
                vec_pro(inx) = adproyecto_ant
                inx = inx + 1
            End If
        Else
            rs_gti_acumdiario.MoveNext
        End If
    Loop
    rs_gti_acumdiario.Close

    If tot_jor <> 0 Then
        tiene_tarja = True
        inxfin = inx - 1
    End If
    
    
    If Not tiene_tarja Then
        '--------------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------
        'Armar el vector de imputacion
        '--------------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------------
        If Not HayMasinivternro Then
            Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene apertura"
        Else
            'El modelo tiene apertura
            Flog.writeline Espacios(Tabulador * 1) & "El modelo tiene apertura"
            
            'Borro el vector de imputacion
            Call BorrarVectorImputacion
            ind_imp_act = 0
            
            Flog.writeline Espacios(Tabulador * 1) & "Busqueda de Distribuion contable"
            
            'Distribucion en % Fijos para cada empleado
            StrSql = "SELECT * FROM imputacion where imputacion.ternro = " & Ternro & _
                     " AND imputacion.masinro = " & masinro & _
                     " AND imputacion.porcentaje <> 0 " & _
                     " ORDER BY imputacion.impnro"
            OpenRecordset StrSql, rs_Imputacion
        
            If Not rs_Imputacion.EOF Then
                Flog.writeline Espacios(Tabulador * 1) & "El empleado tiene Distribucion Contable"
                
                'ARMO EL VECTOR DE IMPUTACIONES EN BASE A LO CARGADO DESDE ADP
                porcentaje = 0
                Do While Not rs_Imputacion.EOF
                    
                    ind_imp_act = ind_imp_act + 1
                    
                    'Controlo desbordamiento
                    If ind_imp_act > max_ind_imp Then
                        Flog.writeline Espacios(Tabulador * 1) & "Error. El indice del vector de imputaciones supero el max de " & max_ind_imp
                    End If
                    
                    imputaTenro1 = IIf(EsNulo(rs_Imputacion!Tenro), 0, rs_Imputacion!Tenro)
                    imputaTenro2 = IIf(EsNulo(rs_Imputacion!tenro2), 0, rs_Imputacion!tenro2)
                    imputaTenro3 = IIf(EsNulo(rs_Imputacion!tenro3), 0, rs_Imputacion!tenro3)
                    imputaEstrnro1 = IIf(EsNulo(rs_Imputacion!Estrnro), 0, rs_Imputacion!Estrnro)
                    imputaEstrnro2 = IIf(EsNulo(rs_Imputacion!estrnro2), 0, rs_Imputacion!estrnro2)
                    imputaEstrnro3 = IIf(EsNulo(rs_Imputacion!estrnro3), 0, rs_Imputacion!estrnro3)
                    
                    'Miro que componente tiene cargada
                    
                    'Si el modelo tiene apertura por tipo estructura 1
                    If (Masinivternro1 <> 0) Then
                       'cargo el tipo de estructura (debe coincidir con la del modelo)
                        vec_imputacion(ind_imp_act).te1 = Masinivternro1
                        If imputaEstrnro1 <> 0 Then
                            'cargo con la estructura de la imputacion
                            vec_imputacion(ind_imp_act).Estructura1 = imputaEstrnro1
                        Else
                            'cargo con la estructura del empleado
                            vec_imputacion(ind_imp_act).Estructura1 = estr1
                        End If
                    End If
                    
                    'Si el modelo tiene apertura por tipo estructura 2
                    If (Masinivternro2 <> 0) Then
                        'cargo el tipo de estructura (debe coincidir con la del modelo)
                        vec_imputacion(ind_imp_act).te2 = Masinivternro2
                        If imputaEstrnro2 <> 0 Then
                            'cargo con la estructura de la imputacion
                            vec_imputacion(ind_imp_act).Estructura2 = imputaEstrnro2
                        Else
                            'cargo con la estructura del empleado
                            vec_imputacion(ind_imp_act).Estructura2 = estr2
                        End If
                    End If
                    
                    'Si el modelo tiene apertura por tipo estructura 3
                    If (Masinivternro3 <> 0) Then
                        'cargo el tipo de estructura (debe coincidir con la del modelo)
                        vec_imputacion(ind_imp_act).te3 = Masinivternro3
                        If imputaEstrnro3 <> 0 Then
                            'cargo con la estructura de la imputacion
                            vec_imputacion(ind_imp_act).Estructura3 = imputaEstrnro3
                        Else
                            'cargo con la estructura del empleado
                            vec_imputacion(ind_imp_act).Estructura3 = estr3
                        End If
                    End If
                    
                    'Cargo el porcentaje
                    vec_imputacion(ind_imp_act).porcentaje = rs_Imputacion!porcentaje
                    
                    Flog.writeline Espacios(Tabulador * 1) & "Vector de Imputacion( " & ind_imp_act & ") TipoEst1 = " & vec_imputacion(ind_imp_act).te1 & " Estr1 = " & vec_imputacion(ind_imp_act).Estructura1 & " TipoEst2 = " & vec_imputacion(ind_imp_act).te2 & " Estr2 = " & vec_imputacion(ind_imp_act).Estructura2 & " TipoEst3 = " & vec_imputacion(ind_imp_act).te3 & " Estr3 = " & vec_imputacion(ind_imp_act).Estructura3 & " Porcentaje = " & vec_imputacion(ind_imp_act).porcentaje
                    
                    porcentaje = porcentaje + rs_Imputacion!porcentaje
                    
                    rs_Imputacion.MoveNext
                Loop
                rs_Imputacion.Close
                
                'Si el porcentaje es < 100 debo completar
                If porcentaje < 100 Then
                    'Si el porcentaje es menor o igual que 1 a la ultima imputacion la corrijo
                    If CDbl(100 - porcentaje) <= 1 Then
                        'A la ultima imputacion la completo con lo faltante
                        vec_imputacion(ind_imp_act).porcentaje = vec_imputacion(ind_imp_act).porcentaje + (100 - porcentaje)
                        Flog.writeline Espacios(Tabulador * 1) & "Correccion de la componente " & ind_imp_act & " por error de redondeo."
                        Flog.writeline Espacios(Tabulador * 1) & "Vector de Imputacion( " & ind_imp_act & ") TipoEst1 = " & vec_imputacion(ind_imp_act).te1 & " Estr1 = " & vec_imputacion(ind_imp_act).Estructura1 & " TipoEst2 = " & vec_imputacion(ind_imp_act).te2 & " Estr2 = " & vec_imputacion(ind_imp_act).Estructura2 & " TipoEst3 = " & vec_imputacion(ind_imp_act).te3 & " Estr3 = " & vec_imputacion(ind_imp_act).Estructura3 & " Porcentaje = " & vec_imputacion(ind_imp_act).porcentaje
                    Else
                        'sino inserto otra componente con el % faltante con la estructura del empleado
                        
                        ind_imp_act = ind_imp_act + 1
                        
                        If Masinivternro1 <> 0 Then
                            vec_imputacion(ind_imp_act).te1 = Masinivternro1
                            vec_imputacion(ind_imp_act).Estructura1 = estr1
                        End If
                        If Masinivternro2 <> 0 Then
                            vec_imputacion(ind_imp_act).te2 = Masinivternro2
                            vec_imputacion(ind_imp_act).Estructura2 = estr2
                        End If
                        If Masinivternro3 <> 0 Then
                            vec_imputacion(ind_imp_act).te3 = Masinivternro3
                            vec_imputacion(ind_imp_act).Estructura3 = estr3
                        End If
                        
                        vec_imputacion(ind_imp_act).porcentaje = (100 - porcentaje)
                        
                        Flog.writeline Espacios(Tabulador * 1) & "El % no es 100, completo con las estructuras del empleado."
                        Flog.writeline Espacios(Tabulador * 1) & "Vector de Imputacion( " & ind_imp_act & ") TipoEst1 = " & vec_imputacion(ind_imp_act).te1 & " Estr1 = " & vec_imputacion(ind_imp_act).Estructura1 & " TipoEst2 = " & vec_imputacion(ind_imp_act).te2 & " Estr2 = " & vec_imputacion(ind_imp_act).Estructura2 & " TipoEst3 = " & vec_imputacion(ind_imp_act).te3 & " Estr3 = " & vec_imputacion(ind_imp_act).Estructura3 & " Porcentaje = " & vec_imputacion(ind_imp_act).porcentaje
                    End If
                    
                End If
                
            Else
                rs_Imputacion.Close
                Flog.writeline Espacios(Tabulador * 1) & "El empleado NO posee Distribucion Contable"
                'Armo el vector de imputaciones con la distribucion del empleado al 100%
                
                ind_imp_act = ind_imp_act + 1
                
                If Masinivternro1 <> 0 Then
                    vec_imputacion(ind_imp_act).te1 = Masinivternro1
                    vec_imputacion(ind_imp_act).Estructura1 = estr1
                End If
                If Masinivternro2 <> 0 Then
                    vec_imputacion(ind_imp_act).te2 = Masinivternro2
                    vec_imputacion(ind_imp_act).Estructura2 = estr2
                End If
                If Masinivternro3 <> 0 Then
                    vec_imputacion(ind_imp_act).te3 = Masinivternro3
                    vec_imputacion(ind_imp_act).Estructura3 = estr3
                End If
                
                vec_imputacion(ind_imp_act).porcentaje = 100
                
                Flog.writeline Espacios(Tabulador * 1) & "Vector de Imputacion( " & ind_imp_act & ") TipoEst1 = " & vec_imputacion(ind_imp_act).te1 & " Estr1 = " & vec_imputacion(ind_imp_act).Estructura1 & " TipoEst2 = " & vec_imputacion(ind_imp_act).te2 & " Estr2 = " & vec_imputacion(ind_imp_act).Estructura2 & " TipoEst3 = " & vec_imputacion(ind_imp_act).te3 & " Estr3 = " & vec_imputacion(ind_imp_act).Estructura3 & " Porcentaje = " & vec_imputacion(ind_imp_act).porcentaje
                
            End If
    
        End If 'Si el modelo tiene distribucion contable
    
    End If 'Si el empleado no tiene Distribucion por Tarja
    
    
    'BORRO EL VECTOR QUE ACUMULA DETALLES DEL EMPLEADO
    If HACE_TRAZA Then
        Call BorrarDetalleAsiAuxEmp
    End If
           
            
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    'Calculo de las lineas del modelo
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    StrSql = "SELECT * FROM mod_linea WHERE masinro = " & masinro
    OpenRecordset StrSql, rs_Mod_Linea
    Do While Not rs_Mod_Linea.EOF
        
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 2) & "Procesando linea: " & rs_Mod_Linea!LinaOrden & " - " & rs_Mod_Linea!linadesc & " Cuenta: " & rs_Mod_Linea!linacuenta
        
        'Verifico que la cuenta no sea niveladora
        If UCase(rs_Mod_Linea!linadesc) = "NIVELADORA" Then
            'Cuenta Niveladora
            Flog.writeline Espacios(Tabulador * 3) & "Cuenta Niveladora. No se realiza acumulacion de la misma."
        Else
            'Analizo la cuenta
            
            'SI HACE TRAZA BORRO EL VECTOR QUE ACUMULA DETALLES DE EMPLEADO Y CUENTA
            If HACE_TRAZA Then
                Call BorrarDetalleAsiAux
            End If
            
            'Inicializo el monto de la linea
            monto_linea = 0
            
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            'BUSQUEDA DE ACUMULADORES QUE CONTRIBUYEN EN LA LINEA
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            StrSql = "SELECT * FROM asi_acu " & _
                     " WHERE asi_acu.masinro = " & rs_Mod_Linea!masinro & _
                     " AND asi_acu.linaorden =" & rs_Mod_Linea!LinaOrden
            OpenRecordset StrSql, rs_Asi_Acu_Con
    
            Do While Not rs_Asi_Acu_Con.EOF
            
                StrSql = "SELECT * FROM acu_liq " & _
                         " INNER JOIN acumulador ON acumulador.acunro = acu_liq.acunro " & _
                         " WHERE acu_liq.cliqnro = " & cliqnro & _
                         " AND acu_liq.acunro =" & rs_Asi_Acu_Con!acuNro
                OpenRecordset StrSql, rs_Asi_monto
                
                If Not rs_Asi_monto.EOF Then
                    
                    monto_aux = IIf(EsNulo(rs_Asi_monto!almonto), 0, rs_Asi_monto!almonto)
                    cant_aux = IIf(EsNulo(rs_Asi_monto!alcant), 0, rs_Asi_monto!alcant)
                    signo = "(+/-)"
                    'Si signo + o - entonces tomar valor absoluto
                    If rs_Asi_Acu_Con!signo <> 3 Then
                        monto_aux = Abs(monto_aux)
                        signo = "(+)"
                        'Si signo - entonces lo hago restar
                        If rs_Asi_Acu_Con!signo = 2 Then
                            monto_aux = -monto_aux
                            signo = "(-)"
                        End If
                    End If
                    
                    Flog.writeline Espacios(Tabulador * 3) & "ACUMULADOR " & rs_Asi_monto!acuNro & " " & rs_Asi_monto!acudesabr & " - MONTO = " & rs_Asi_monto!almonto & " - SIGNO = " & signo
                    monto_linea = monto_linea + monto_aux
                    
                    'GUARDO LOS DETALLES DEL ACUMULADOR QUE CONTRIBUYE EN EL VECTOR DE DETALLE POR LINEA EMPLEADO
                    If HACE_TRAZA Then
                        Call InsertarDetalleAsiAuxTarja(rs_Mod_Linea!masinro, NroVol, rs_Mod_Linea!linacuenta, rs_Mod_Linea!LinaOrden, rs_tercero!empleg, rs_tercero!terape & " " & rs_tercero!ternom, rs_Asi_monto!acuNro & "-" & rs_Asi_monto!acudesabr, cant_aux, monto_aux, 100, rs_tercero!Ternro, rs_Asi_monto!acuNro, 2, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H)
                    End If
                    
                End If 'rs_Asi_monto
                rs_Asi_monto.Close
                    
                rs_Asi_Acu_Con.MoveNext
            Loop
            
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            'BUSQUEDA DE CONCEPTOS QUE CONTRIBUYEN EN LA LINEA
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            StrSql = "SELECT * FROM asi_con " & _
                     " WHERE asi_con.masinro = " & rs_Mod_Linea!masinro & _
                     " AND asi_con.linaorden =" & rs_Mod_Linea!LinaOrden
            OpenRecordset StrSql, rs_Asi_Acu_Con
    
            Do While Not rs_Asi_Acu_Con.EOF
            
                StrSql = "SELECT * FROM detliq " & _
                         " INNER JOIN concepto ON concepto.concnro = detliq.concnro " & _
                         " WHERE detliq.cliqnro = " & cliqnro & _
                         " AND detliq.concnro =" & rs_Asi_Acu_Con!ConcNro
                OpenRecordset StrSql, rs_Asi_monto
                
                If Not rs_Asi_monto.EOF Then
                    
                    monto_aux = IIf(EsNulo(rs_Asi_monto!dlimonto), 0, rs_Asi_monto!dlimonto)
                    cant_aux = IIf(EsNulo(rs_Asi_monto!dlicant), 0, rs_Asi_monto!dlicant)
                    signo = "(+/-)"
                    'Si signo + o - entonces tomar valor absoluto
                    If rs_Asi_Acu_Con!signo <> 3 Then
                        monto_aux = Abs(monto_aux)
                        signo = "(+)"
                        'Si signo - entonces lo hago restar
                        If rs_Asi_Acu_Con!signo = 2 Then
                            monto_aux = -monto_aux
                            signo = "(-)"
                        End If
                    End If
                    
                    Flog.writeline Espacios(Tabulador * 3) & "CONCEPTO " & rs_Asi_monto!ConcCod & " " & rs_Asi_monto!concabr & " - MONTO = " & rs_Asi_monto!dlimonto & " - SIGNO = " & signo
                    monto_linea = monto_linea + monto_aux
                    
                    'GUARDO LOS DETALLES DEL CONCEPTO QUE CONTRIBUYE EN EL VECTOR DE DETALLE POR LINEA Y EMPLEADO
                    If HACE_TRAZA Then
                        Call calcularMontoCantTarja(rs_Mod_Linea!masinro, NroVol, rs_Mod_Linea!linacuenta, rs_Mod_Linea!LinaOrden, rs_tercero!empleg, rs_tercero!terape & " " & rs_tercero!ternom, rs_Asi_monto!ConcCod & "-" & rs_Asi_monto!concabr, cant_aux, monto_aux, rs_tercero!Ternro, rs_Asi_monto!ConcNro, 1, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H)
'                        Call InsertarDetalleAsiAux(rs_Mod_Linea!masinro, NroVol, rs_Mod_Linea!linacuenta, rs_Mod_Linea!LinaOrden, rs_tercero!empleg, rs_tercero!terape & " " & rs_tercero!ternom, rs_Asi_monto!Conccod & "-" & rs_Asi_monto!concabr, cant_aux, monto_aux, rs_tercero!ternro, rs_Asi_monto!concnro, 1, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H)
                    End If
                                    
                End If 'rs_Asi_monto
                rs_Asi_monto.Close
                    
                rs_Asi_Acu_Con.MoveNext
            Loop
            
            Flog.writeline Espacios(Tabulador * 2) & "MONTO LINEA: " & Round(monto_linea, 4)
            
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            'Insercion en la linea
            '--------------------------------------------------------------------------------------
            '--------------------------------------------------------------------------------------
            'monto_linea = Round(monto_linea, 4)
            
            If tiene_tarja Then
            
                'Debo distribuir de acuerdo al vector de distribucion de Tarja
                Flog.writeline Espacios(Tabulador * 2) & "Distribucion del monto de la linea por el vector de Tarja."

                'Para ver si la suma de los valores parciales de las lineas es igual al monto total de la linea
                'Sino corrijo por redondeo
                MontoRedondeo = 0
                generoAlguna = False
                
                For indice = 0 To inxfin
                
                    'Calculo el monto a imputar de acuerdo al vector
                    'MontoAImputar = Round((monto_linea * vec_imputacion(indice).porcentaje / 100), 4)
                    MontoAImputar = (monto_linea * vec_jor(indice) / tot_jor)
                    Flog.writeline Espacios(Tabulador * 3) & vec_jor(indice) / tot_jor & " % del monto de la linea = " & MontoAImputar
                    
'                    Flog.writeline Espacios(Tabulador * 3) & "Aplicando los Filtros de la linea de orden " & rs_Mod_Linea!LinaOrden & " Para la componente " & indice & " del vector de imputacion."
'                    Call FiltrosLinea(masinro, rs_Mod_Linea!LinaOrden, Masinivternro1, Masinivternro2, Masinivternro3, estr1, estr2, estr3, Generar)
'                    If Generar Then
                        
'                        Flog.writeline Espacios(Tabulador * 3) & "Filtro OK "
                        generoAlguna = True
                        cuenta = rs_Mod_Linea!linacuenta
                        Call ArmarCuentaTarja(cuenta, rs_tercero!Ternro, rs_tercero!empleg, Masinivternro1, Masinivternro2, Masinivternro3, estr1, estr2, estr3, True, vol_Fec_Asiento, rs_Mod_Linea!linaD_H, vec_cta(indice), vec_pro(indice))
                        Flog.writeline Espacios(Tabulador * 3) & "ARMADO DE CUENTA: " & rs_Mod_Linea!linacuenta & " ----------> " & cuenta
                                    
                        Call InsertarVectorLineaAsiAux(cuenta, rs_Mod_Linea!LinaOrden, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H, MontoAImputar)
                        If HACE_TRAZA Then
                            Call ResolverDetalleAsiTarja(NroVol, masinro, cuenta, vec_jor(indice) / tot_jor * 100, estr1, estr2, estr3)
'                            Call ResolverDetalleAsi(NroVol, masinro, cuenta, 100, vec_imputacion(indice).Estructura1, vec_imputacion(indice).Estructura2, vec_imputacion(indice).Estructura3)
                        End If

'                    End If
                    
                    'Acumulo en el redondeo
                    MontoRedondeo = MontoRedondeo + MontoAImputar
                    
                Next
                
                If generoAlguna Then
                    'REDONDEO
                    If Round(MontoRedondeo, 4) <> Round(monto_linea, 4) Then
                        Flog.writeline Espacios(Tabulador * 3) & "DIFERENCIA DE REDONDEO " & FormatNumber(MontoRedondeo - monto_linea, 100)
                    End If
                End If
                    
            Else
                If Not HayMasinivternro Then
                    'Si el modelo no tiene distribucion, no tengo vector de imputacion, el 100% de monto_linea va a la linea
                    cuenta = rs_Mod_Linea!linacuenta
                    Call ArmarCuentaTarja(cuenta, rs_tercero!Ternro, rs_tercero!empleg, 0, 0, 0, 0, 0, 0, False, vol_Fec_Asiento, rs_Mod_Linea!linaD_H, "", "")
                    Flog.writeline Espacios(Tabulador * 3) & "ARMADO DE CUENTA: " & rs_Mod_Linea!linacuenta & " ----------> " & cuenta
                    Call InsertarVectorLineaAsiAux(cuenta, rs_Mod_Linea!LinaOrden, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H, monto_linea)
                    
                    'SI HACE TRAZA ENTONCES RESOLVER DATOS FALTANTES
                    If HACE_TRAZA Then
                        Call ResolverDetalleAsiTarja(NroVol, masinro, cuenta, 100, 0, 0, 0)
                    End If
                    
                Else
                    'Debo distribuir de acuerdo al vector de distribucion
                    Flog.writeline Espacios(Tabulador * 2) & "Distribucion del monto de la linea por el vector de imputacion."
    
                    'Para ver si la suma de los valores parciales de las lineas es igual al monto total de la linea
                    'Sino corrijo por redondeo
                    MontoRedondeo = 0
                    generoAlguna = False
                    
                    For indice = 1 To ind_imp_act
                    
                        'Calculo el monto a imputar de acuerdo al vector
                        'MontoAImputar = Round((monto_linea * vec_imputacion(indice).porcentaje / 100), 4)
                        MontoAImputar = (monto_linea * vec_imputacion(indice).porcentaje / 100)
                        Flog.writeline Espacios(Tabulador * 3) & vec_imputacion(indice).porcentaje & " % del monto de la linea = " & MontoAImputar
                        
                        Flog.writeline Espacios(Tabulador * 3) & "Aplicando los Filtros de la linea de orden " & rs_Mod_Linea!LinaOrden & " Para la componente " & indice & " del vector de imputacion."
                        Call FiltrosLinea(masinro, rs_Mod_Linea!LinaOrden, Masinivternro1, Masinivternro2, Masinivternro3, vec_imputacion(indice).Estructura1, vec_imputacion(indice).Estructura2, vec_imputacion(indice).Estructura3, Generar)
                        If Generar Then
                            
                            Flog.writeline Espacios(Tabulador * 3) & "Filtro OK "
                            generoAlguna = True
                            cuenta = rs_Mod_Linea!linacuenta
                            Call ArmarCuentaTarja(cuenta, rs_tercero!Ternro, rs_tercero!empleg, Masinivternro1, Masinivternro2, Masinivternro3, vec_imputacion(indice).Estructura1, vec_imputacion(indice).Estructura2, vec_imputacion(indice).Estructura3, False, vol_Fec_Asiento, rs_Mod_Linea!linaD_H, "", "")
                            Flog.writeline Espacios(Tabulador * 3) & "ARMADO DE CUENTA: " & rs_Mod_Linea!linacuenta & " ----------> " & cuenta
                                        
                            Call InsertarVectorLineaAsiAux(cuenta, rs_Mod_Linea!LinaOrden, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H, MontoAImputar)
                            If HACE_TRAZA Then
                                Call ResolverDetalleAsiTarja(NroVol, masinro, cuenta, 100, vec_imputacion(indice).Estructura1, vec_imputacion(indice).Estructura2, vec_imputacion(indice).Estructura3)
                            End If
    
                        End If
                        
                        'Acumulo en el redondeo
                        MontoRedondeo = MontoRedondeo + MontoAImputar
                        
                    Next
                    
                    If generoAlguna Then
                        'REDONDEO
                        If Round(MontoRedondeo, 4) <> Round(monto_linea, 4) Then
                            Flog.writeline Espacios(Tabulador * 3) & "DIFERENCIA DE REDONDEO " & FormatNumber(MontoRedondeo - monto_linea, 100)
                        End If
                    End If
                    
                End If
                
            End If
            
        End If 'No es cuenta niveladora
            
        'Paso a la siguiente linea
        rs_Mod_Linea.MoveNext
        
    Loop
    
    
ya_procesado:
'    ternro_ant = ternro
    
If rs_tercero.State = adStateOpen Then rs_tercero.Close
If rs_Imputacion.State = adStateOpen Then rs_Imputacion.Close
If rs_Mod_Linea.State = adStateOpen Then rs_Mod_Linea.Close
If rs_Asi_monto.State = adStateOpen Then rs_Asi_monto.Close
If rs_Asi_Acu_Con.State = adStateOpen Then rs_Asi_Acu_Con.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_gti_acumdiario.State = adStateOpen Then rs_gti_acumdiario.Close
If rs_tiph_con.State = adStateOpen Then rs_tiph_con.Close
If rs_detliq.State = adStateOpen Then rs_detliq.Close

Set rs_tercero = Nothing
Set rs_Imputacion = Nothing
Set rs_Mod_Linea = Nothing
Set rs_Asi_monto = Nothing
Set rs_Asi_Acu_Con = Nothing
Set rs_Estructura = Nothing


Exit Sub
'Manejador de Errores del procedimiento
ME_Acumular:
    Flog.writeline "Error: " & Err.Description
'    Resume Next
    Flog.writeline "Procedimiento: Acumular"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
End Sub


Public Sub calcularMontoCantTarja(ByVal masinro As Integer, ByVal NroVol As Integer, ByVal linacuenta As String, ByVal LinaOrden As Integer, ByVal empleg As Long, ByVal Empleado As String, ByVal concepto As String, ByVal cant As Double, ByVal Monto As Double, ByVal Ternro As Long, ByVal ConcNro As Integer, ByVal tipoorigen As Integer, ByVal linadesc As String, ByVal linaD_H As Integer)
'ByRef Monto As Double, ByRef cant As Double, ByVal concepto As Integer)
' --------------------------------------------------------------------------------------------
' Descripcion: Imputa el porcentaje que le corresponde en el monto y la cantidad
' Autor      : Fernando Favre
' Fecha      : 05/06/2007
' --------------------------------------------------------------------------------------------
Dim I As Integer
Dim Seguir As Boolean
Dim totalHoras As Double
Dim monto_aux As Double
Dim cant_aux As Double
Dim porcentaje_aux As Double

On Error GoTo ME_calcularMontoCantTarja
    
    I = 0
    Seguir = True
    
    Do While I < max_ind_con And Seguir
        If vec_con_tot(I).ConcNro = ConcNro Then
            Seguir = False
        Else
            I = I + 1
        End If
    Loop
    
    If Seguir Then
        Call InsertarDetalleAsiAuxTarja(masinro, NroVol, linacuenta, LinaOrden, empleg, Empleado, concepto, cant, Monto, 100, Ternro, ConcNro, tipoorigen, linadesc, linaD_H) ' rs_Mod_Linea!masinro, NroVol, rs_Mod_Linea!linacuenta, rs_Mod_Linea!LinaOrden, rs_tercero!empleg, rs_tercero!terape & " " & rs_tercero!ternom, rs_Asi_monto!Conccod & "-" & rs_Asi_monto!concabr, cant_aux, monto_aux, porcentaje_aux, rs_tercero!ternro, rs_Asi_monto!concnro, 1, rs_Mod_Linea!linadesc, rs_Mod_Linea!linaD_H)
    Else
        totalHoras = vec_con_tot(I).canthoras
    
        I = 0
        
        Do While I < max_ind_con
            If vec_con(I).ConcNro = ConcNro Then
                monto_aux = Monto * vec_con(I).canthoras / totalHoras
                cant_aux = cant * vec_con(I).canthoras / totalHoras
                porcentaje_aux = vec_con(I).canthoras / totalHoras * 100
                Call InsertarDetalleAsiAuxTarja(masinro, NroVol, linacuenta, LinaOrden, empleg, Empleado, concepto, cant_aux, monto_aux, porcentaje_aux, Ternro, ConcNro, tipoorigen, linadesc, linaD_H)
            End If
            I = I + 1
        Loop
    End If
    
Exit Sub
'Manejador de Errores del procedimiento
ME_calcularMontoCantTarja:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: calcularMontoCantTarja"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
    
End Sub


Public Sub InsertarVectorConceptoTarja(ByVal ConcNro As Integer, ByVal adctaasiento As String, ByVal adproyecto As String, ByVal adcanthoras As Double, ByVal valorj As Double)
' --------------------------------------------------------------------------------------------
' Descripcion: Guarda en el vector de concepto la cantidad de horas. Se utiliza para TARJA
' Autor      : Fernando Favre
' Fecha      : 05/06/2007
' --------------------------------------------------------------------------------------------
Dim I As Integer
Dim Seguir As Boolean
    
On Error GoTo ME_InsertarVectorConceptoTarja

    ' Ingreso la proporcion de canthoras para el concepto, cuenta y proyecto
    I = 0
    Seguir = True
    
    Do While I < max_ind_con And Seguir
        If vec_con(I).ConcNro = ConcNro And vec_con(I).cuenta = adctaasiento And vec_con(I).proyecto = adproyecto Then
            Seguir = False
        Else
            I = I + 1
        End If
    Loop
    
    If Seguir Then
        I = ind_con_act
        ind_con_act = ind_con_act + 1
        vec_con(I).ConcNro = ConcNro
        vec_con(I).cuenta = adctaasiento
        vec_con(I).proyecto = adproyecto
    End If
'    vec_con(I).canthoras = vec_con(I).canthoras + (adcanthoras * valorj)
    vec_con(I).canthoras = vec_con(I).canthoras + adcanthoras '* valorj)
    
    ' Ingreso la proporcion de canthoras para el concepto
    I = 0
    Seguir = True
    
    Do While I < max_ind_con And Seguir
        If vec_con_tot(I).ConcNro = ConcNro Then
            Seguir = False
        Else
            I = I + 1
        End If
    Loop
    
    If Seguir Then
        I = ind_con_tot_act
        ind_con_tot_act = ind_con_tot_act + 1
        vec_con_tot(I).ConcNro = ConcNro
    End If
'    vec_con_tot(I).canthoras = vec_con_tot(I).canthoras + (adcanthoras * valorj)
    vec_con_tot(I).canthoras = vec_con_tot(I).canthoras + adcanthoras '* valorj)

Exit Sub
'Manejador de Errores del procedimiento
ME_InsertarVectorConceptoTarja:
    Flog.writeline "Error: " & Err.Description & " - Error: I = " & I
    Flog.writeline "Error: ConcNro:" & CStr(ConcNro) & " - adctaasiento: " & CStr(adctaasiento) & " - adproyecto: " & CStr(adproyecto) & " - adcanthoras: " & CStr(adcanthoras) & " - valorj: " & CStr(valorj)
'    Resume Next
    Flog.writeline "Procedimiento: InsertarVectorConceptoTarja"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
End Sub


Public Sub FiltrosLinea(ByVal masinro As Long, ByVal LinaOrden As Long, ByVal Masinivternro1 As Long, ByVal Masinivternro2 As Long, ByVal Masinivternro3 As Long, ByVal Estructura1 As Long, ByVal Estructura2 As Long, ByVal Estructura3 As Long, ByRef Generar As Boolean)
' --------------------------------------------------------------------------------------------
' Descripcion: verifica que se respeten las estructuras del filtro
' Autor      : Martin Ferraro
' Fecha      : 24/07/2006
' --------------------------------------------------------------------------------------------

Dim rs_Filtro As New ADODB.Recordset

On Error GoTo ME_FiltrosLinea

Generar = True
     
     
    'Si el modelo tiene nivel de apertura 1 miro si la linea tiene configurado algun filtro de primer nivel
    If Masinivternro1 <> 0 Then
        'reviso que tenga un filtro para el tipo de estructura
        StrSql = "SELECT * FROM mod_lin_filtro "
        StrSql = StrSql & " WHERE masinro = " & masinro
        StrSql = StrSql & " AND linaorden = " & LinaOrden
        StrSql = StrSql & " AND tenro = " & Masinivternro1
        OpenRecordset StrSql, rs_Filtro
        If Not rs_Filtro.EOF Then
            'si tiene filtro busco que exista para la estructura
            StrSql = "SELECT * FROM mod_lin_filtro "
            StrSql = StrSql & " WHERE masinro = " & masinro
            StrSql = StrSql & " AND linaorden = " & LinaOrden
            StrSql = StrSql & " AND tenro = " & Masinivternro1
            StrSql = StrSql & " AND estrnro = " & Estructura1
            If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
            OpenRecordset StrSql, rs_Filtro
            If rs_Filtro.EOF Then
                Generar = False
                Flog.writeline Espacios(Tabulador * 3) & "La linea no supera el filtro de PRIMER NIVEL."
            End If
        End If
    End If
    
    'Si el modelo tiene nivel de apertura 2 y supero el primer filtro, miro si la linea tiene configurado algun filtro de segundo nivel
    If ((Masinivternro2 <> 0) And Generar) Then
        'reviso que tenga un filtro para el tipo de estructura
        StrSql = "SELECT * FROM mod_lin_filtro "
        StrSql = StrSql & " WHERE masinro = " & masinro
        StrSql = StrSql & " AND linaorden = " & LinaOrden
        StrSql = StrSql & " AND tenro = " & Masinivternro2
        If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
        OpenRecordset StrSql, rs_Filtro
        If Not rs_Filtro.EOF Then
            'si tiene filtro busco que exista para la estructura
            StrSql = "SELECT * FROM mod_lin_filtro "
            StrSql = StrSql & " WHERE masinro = " & masinro
            StrSql = StrSql & " AND linaorden = " & LinaOrden
            StrSql = StrSql & " AND tenro = " & Masinivternro2
            StrSql = StrSql & " AND estrnro = " & Estructura2
            If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
            OpenRecordset StrSql, rs_Filtro
            If rs_Filtro.EOF Then
                Generar = False
                Flog.writeline Espacios(Tabulador * 3) & "La linea no supera el filtro de SEGUNDO NIVEL."
            End If
        End If
    End If

    'Si el modelo tiene nivel de apertura 3 y supero el segundo filtro, miro si la linea tiene configurado algun filtro de tercer nivel
    If ((Masinivternro3 <> 0) And Generar) Then
        'reviso que tenga un filtro para el tipo de estructura
        StrSql = "SELECT * FROM mod_lin_filtro "
        StrSql = StrSql & " WHERE masinro = " & masinro
        StrSql = StrSql & " AND linaorden = " & LinaOrden
        StrSql = StrSql & " AND tenro = " & Masinivternro3
        If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
        OpenRecordset StrSql, rs_Filtro
        If Not rs_Filtro.EOF Then
            'si tiene filtro busco que exista para la estructura
            StrSql = "SELECT * FROM mod_lin_filtro "
            StrSql = StrSql & " WHERE masinro = " & masinro
            StrSql = StrSql & " AND linaorden = " & LinaOrden
            StrSql = StrSql & " AND tenro = " & Masinivternro3
            StrSql = StrSql & " AND estrnro = " & Estructura3
            If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
            OpenRecordset StrSql, rs_Filtro
            If rs_Filtro.EOF Then
                Generar = False
                Flog.writeline Espacios(Tabulador * 3) & "La linea no supera el filtro de TERCER NIVEL."
            End If
        End If
    End If

If rs_Filtro.State = adStateOpen Then rs_Filtro.Close
Set rs_Filtro = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_FiltrosLinea:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: FiltrosLinea"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql


End Sub


Public Sub BorrarVectorImputacion()

On Error GoTo ME_BorrarVectorImputacion

    For ind_imp_act = 1 To max_ind_imp
        vec_imputacion(ind_imp_act).te1 = 0
        vec_imputacion(ind_imp_act).te2 = 0
        vec_imputacion(ind_imp_act).te3 = 0
        vec_imputacion(ind_imp_act).Estructura1 = 0
        vec_imputacion(ind_imp_act).Estructura2 = 0
        vec_imputacion(ind_imp_act).Estructura3 = 0
        vec_imputacion(ind_imp_act).porcentaje = 0
    Next

Exit Sub
'Manejador de Errores del procedimiento
ME_BorrarVectorImputacion:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: BorrarVectorImputacion"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub


Public Sub InicializarVectorLineaAsiAux()
    
On Error GoTo ME_InicializarVectorLineaAsiAux
    
    For ind_lineaAsiAux = 1 To max_ind_lineaAsiAux
        lineaAsiAux(ind_lineaAsiAux).cuenta = ""
        lineaAsiAux(ind_lineaAsiAux).Linea = 0
        lineaAsiAux(ind_lineaAsiAux).desclinea = ""
        lineaAsiAux(ind_lineaAsiAux).dh = 0
        lineaAsiAux(ind_lineaAsiAux).Monto = 0
    Next
    
    ind_lineaAsiAux = 0
    
Exit Sub
'Manejador de Errores del procedimiento
ME_InicializarVectorLineaAsiAux:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: InicializarVectorLineaAsiAux"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql


End Sub

Public Sub GuardarLineaAsi(ByVal vol_cod As Long, ByVal masinro As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Inserta las cuentas en la base de datos en linea_asi
' Autor      : Martin Ferraro
' Fecha      : 04/08/2006
' --------------------------------------------------------------------------------------------
Dim indice As Long
Dim rs_Linea_asi As New ADODB.Recordset
    
On Error GoTo ME_GuardarLineaAsi

    For indice = 1 To ind_lineaAsiAux
            
        'Miro si la linea ya esta en la base para el proceso y modelo
        StrSql = "SELECT * FROM linea_asi " & _
                 " WHERE linea_asi.vol_cod = " & vol_cod & _
                 " AND linea_asi.cuenta  = '" & Mid(lineaAsiAux(indice).cuenta, 1, 50) & "'" & _
                 " AND linea_asi.masinro = " & masinro
        OpenRecordset StrSql, rs_Linea_asi
        
        If rs_Linea_asi.EOF Then
        
            'No existe una linea con esa cuenta, entonces la inserto
            StrSql = "INSERT INTO linea_asi (cuenta,vol_cod,masinro,linea,desclinea,monto)" & _
                     " VALUES ('" & Mid(lineaAsiAux(indice).cuenta, 1, 50) & _
                     "'," & vol_cod & _
                     "," & masinro & _
                     "," & lineaAsiAux(indice).Linea & _
                     ",'" & Mid(lineaAsiAux(indice).desclinea, 1, 60) & _
                     "'," & Round(lineaAsiAux(indice).Monto, 4) & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
        
            'la linea existe, debo modificar el monto
            StrSql = "UPDATE linea_asi SET monto = monto + " & lineaAsiAux(indice).Monto & _
                     " WHERE linea_asi.vol_cod =" & vol_cod & _
                     " AND linea_asi.cuenta  ='" & Mid(lineaAsiAux(indice).cuenta, 1, 50) & "'" & _
                     " AND linea_asi.masinro =" & masinro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        rs_Linea_asi.Close
    
    Next
    
    'vuelvo a cero el indice de lineaAsiAux
    ind_lineaAsiAux = 0


    If HACE_TRAZA Then
        Call GuardarDetalleAsi
    End If

'cierro todo
If rs_Linea_asi.State = adStateOpen Then rs_Linea_asi.Close
Set rs_Linea_asi = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_GuardarLineaAsi:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: GuardarLineaAsi"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql


End Sub


Public Sub GuardarDetalleAsi()
Dim indice As Long

On Error GoTo ME_GuardarDetalleAsi

    For indice = 1 To ind_detalleAsiAuxEmp
    
            StrSql = "INSERT INTO detalle_asi (masinro, cuenta,dlcantidad,dlcosto1,dlcosto2,dlcosto3,dlcosto4,dldescripcion " & _
                     ",dlmonto,dlmontoacum,dlporcentaje,ternro,empleg,lin_orden,terape,vol_cod, origen, tipoorigen,linadesc,linaD_H)" & _
                     " VALUES (" & detalleAsiAuxEmp(indice).masinro & _
                     ",'" & Mid(detalleAsiAuxEmp(indice).cuenta, 1, 50) & _
                     "'," & detalleAsiAuxEmp(indice).dlcantidad & _
                     "," & detalleAsiAuxEmp(indice).dlcosto1 & _
                     "," & detalleAsiAuxEmp(indice).dlcosto2 & _
                     "," & detalleAsiAuxEmp(indice).dlcosto3 & _
                     ",0" & _
                     ",'" & Mid(detalleAsiAuxEmp(indice).dldescripcion, 1, 60) & _
                     "'," & Round(detalleAsiAuxEmp(indice).dlmonto, 4) & _
                     "," & Round(detalleAsiAuxEmp(indice).dlmontoacum, 4) & _
                     "," & detalleAsiAuxEmp(indice).dlporcentaje & _
                     "," & detalleAsiAuxEmp(indice).Ternro & _
                     "," & detalleAsiAuxEmp(indice).empleg & _
                     "," & detalleAsiAuxEmp(indice).lin_orden & _
                     ",'" & Mid(detalleAsiAuxEmp(indice).terape, 1, 50) & _
                     "'," & detalleAsiAuxEmp(indice).vol_cod & _
                     "," & detalleAsiAuxEmp(indice).Origen & _
                     "," & detalleAsiAuxEmp(indice).tipoorigen & _
                     ",'" & Mid(detalleAsiAuxEmp(indice).linadesc, 1, 40) & _
                     "'," & detalleAsiAuxEmp(indice).linaD_H & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        
    Next

    'Reseteo los indices - 11/05/2007 - Martin Ferraro
    ind_detalleAsiAux = 0
    ind_detalleAsiAuxEmp = 0

Exit Sub
'Manejador de Errores del procedimiento
ME_GuardarDetalleAsi:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: GuardarDetalleAsi"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub


Public Sub InsertarVectorLineaAsiAux(ByVal cuenta As String, ByVal Linea As Long, ByVal desclinea As String, ByVal dh As Integer, ByVal Monto As Double)
' --------------------------------------------------------------------------------------------
' Descripcion: Inserta o modifica la tabla temporal de linea asi
' Autor      : Martin Ferraro
' Fecha      : 04/08/2006
' --------------------------------------------------------------------------------------------
Dim indice As Long
Dim modifica As Boolean
    
On Error GoTo ME_InsertarVectorLineaAsiAux

    'Miro si existe alguna componente del vector con la misma cuenta
    modifica = False
    For indice = 1 To ind_lineaAsiAux
        If lineaAsiAux(indice).cuenta = cuenta Then
            modifica = True
            Exit For
        End If
    Next
    
    If modifica Then
        'Estoy parado en el indice de la cuenta a modif, debo sumar el monto
        lineaAsiAux(indice).Monto = lineaAsiAux(indice).Monto + Monto
    Else
        'Tengo que insertar la cuenta en el arreglo
        ind_lineaAsiAux = ind_lineaAsiAux + 1
        lineaAsiAux(ind_lineaAsiAux).cuenta = cuenta
        lineaAsiAux(ind_lineaAsiAux).Linea = Linea
        lineaAsiAux(ind_lineaAsiAux).desclinea = desclinea
        lineaAsiAux(ind_lineaAsiAux).dh = dh
        lineaAsiAux(ind_lineaAsiAux).Monto = Monto
    End If
    
Exit Sub
'Manejador de Errores del procedimiento
ME_InsertarVectorLineaAsiAux:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: InsertarVectorLineaAsiAux"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub



Public Sub ArmarCuenta(ByRef NroCuenta As String, ByVal Ternro As Long, ByVal Legajo As String, ByVal Masinivternro1 As Long, ByVal Masinivternro2 As Long, ByVal Masinivternro3 As Long, ByVal Estructura1 As Long, ByVal Estructura2 As Long, ByVal Estructura3 As Long)
' --------------------------------------------------------------------------------------------
' Descripcion: Arma la cuenta de acuerdo a la configuracion de la misma
' Autor      : Martin Ferraro
' Fecha      : 04/08/2006
' --------------------------------------------------------------------------------------------
Dim Aux_Cuenta As String
Dim Aux_Legajo As String

Dim DescEstructura1 As String
Dim DescEstructura2 As String
Dim DescEstructura3 As String

Dim I As Integer
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

Dim EsDocumento As Boolean
Dim CantD As Long
Dim TipoD As String
Dim TipoD_Actual As String
Dim DescDocumento As String

Dim rs_Estructura As New ADODB.Recordset
Dim rs_Documento As New ADODB.Recordset

On Error GoTo ME_ArmarCuenta

Aux_Cuenta = NroCuenta
Aux_Legajo = Legajo

'Descripcion de estructura 1
If Masinivternro1 = 0 Then
    'Modelo sin apertura
    DescEstructura1 = "00000000000000000000"
Else
    'Modelo con apertura, busco la descripcion de la estructura
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & Estructura1
    OpenRecordset StrSql, rs_Estructura
    
    If Not rs_Estructura.EOF Then
            DescEstructura1 = IIf(IsNull(rs_Estructura!estrcodext), "00000000000000000000", rs_Estructura!estrcodext & "00000000000000000000")
    Else
        DescEstructura1 = "00000000000000000000"
    End If
    rs_Estructura.Close
End If
DescEstructura1 = Left(DescEstructura1, 20)

'Descripcion de estructura 2
If Masinivternro2 = 0 Then
    'Modelo sin apertura
    DescEstructura2 = "00000000000000000000"
Else
    'Modelo con apertura, busco la descripcion de la estructura
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & Estructura2
    OpenRecordset StrSql, rs_Estructura
    
    If Not rs_Estructura.EOF Then
            DescEstructura2 = IIf(IsNull(rs_Estructura!estrcodext), "00000000000000000000", rs_Estructura!estrcodext & "00000000000000000000")
    Else
        DescEstructura2 = "00000000000000000000"
    End If
    rs_Estructura.Close
End If
DescEstructura2 = Left(DescEstructura2, 20)

'Descripcion de estructura 3
If Masinivternro3 = 0 Then
    'Modelo sin apertura
    DescEstructura3 = "00000000000000000000"
Else
    'Modelo con apertura, busco la descripcion de la estructura
    StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
             " WHERE estrnro = " & Estructura3
    OpenRecordset StrSql, rs_Estructura
    
    If Not rs_Estructura.EOF Then
            DescEstructura3 = IIf(IsNull(rs_Estructura!estrcodext), "00000000000000000000", rs_Estructura!estrcodext & "00000000000000000000")
    Else
        DescEstructura3 = "00000000000000000000"
    End If
    rs_Estructura.Close
End If
DescEstructura3 = Left(DescEstructura3, 20)


PosE1 = 1
PosE2 = 1
PosE3 = 1


'Voy recorriendo de Izquierda a Derecha el aux_cuenta y voy generando el NroCuenta
I = 1
NroCuenta = ""
CantL = 0
CantE = 0
CantD = 0
Termino = False

Do While Not (I > Len(Aux_Cuenta))
    ch = UCase(Mid(Aux_Cuenta, I, 1))

    Select Case ch
    Case "_", "-", ".", "*":
        NroCuenta = NroCuenta & ch
        I = I + 1
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
        NroCuenta = NroCuenta & ch
        I = I + 1
    Case "E": 'Estrcutura
        EsEstructura = True
        CantE = 1
        'leo el nro de la estructura
        I = I + 1
        ch = UCase(Mid(Aux_Cuenta, I, 1))
        TipoE = ch
        Termino = False
        
        Do While EsEstructura And Not Termino
            'leo el siguiente
            I = I + 1
            If Not (I > Len(Aux_Cuenta)) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            Else
                Termino = True
            End If
            
            If ch = "E" And Not Termino Then
                'leo lel nro de la estructura
                I = I + 1
                ch = UCase(Mid(Aux_Cuenta, I, 1))
                TipoE_Actual = ch
                
                Do While TipoE = TipoE_Actual And EsEstructura And Not Termino
                    CantE = CantE + 1
    
                    I = I + 1
                    If Not (I > Len(Aux_Cuenta)) Then
                        ch = UCase(Mid(Aux_Cuenta, I, 1))
                    Else
                        Termino = True
                    End If
                    
                    If ch = "E" Then
                        'leo el nro de la estructura
                        I = I + 1
                        ch = UCase(Mid(Aux_Cuenta, I, 1))
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
                NroCuenta = NroCuenta & Mid(DescEstructura1, PosE1, CantE)
                PosE1 = PosE1 + CantE
                If PosE1 >= 20 Then PosE1 = 1
            Case 2:
                'NroCuenta = NroCuenta & Right(Estructura2, CantE)
                NroCuenta = NroCuenta & Mid(DescEstructura2, PosE2, CantE)
                PosE2 = PosE2 + CantE
                If PosE2 >= 20 Then PosE2 = 1
            Case 3:
                'NroCuenta = NroCuenta & Right(Estructura3, CantE)
                NroCuenta = NroCuenta & Mid(DescEstructura3, PosE3, CantE)
                PosE3 = PosE3 + CantE
                If PosE3 >= 20 Then PosE3 = 1
            End Select
            
            TipoE = TipoE_Actual
            CantE = 1
        Loop
        
    Case "L": 'Legajo
        Termino = False
        CantL = 1
        I = I + 1
        If I <= Len(Aux_Cuenta) Then
            ch = UCase(Mid(Aux_Cuenta, I, 1))
        End If
        
        Do While ch = "L" And Not Termino
            CantL = CantL + 1
            I = I + 1
            If I <= Len(Aux_Cuenta) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            Else
                Termino = True
            End If
        Loop
        
        NroCuenta = NroCuenta & Right(Format(Aux_Legajo, "0000000000"), CantL)
     
  

    Case "D": 'DXX NUMERO DE TIPO DE DOC, DONDE XX ES EL TIPO DE DOC
        EsDocumento = True
        CantD = 1
        'leo el nro de documento
        I = I + 1
        ch = UCase(Mid(Aux_Cuenta, I, 2))
        'Avanzo otro porque lei 2
        I = I + 1
        TipoD = ch
        Termino = False

        Do While EsDocumento And Not Termino
            'leo el siguiente
            I = I + 1
            If Not (I > Len(Aux_Cuenta)) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            Else
                Termino = True
            End If

            If ch = "D" And Not Termino Then
                'leo lel nro de documento
                I = I + 1
                ch = UCase(Mid(Aux_Cuenta, I, 2))
                'Avanzo otro porque lei 2
                I = I + 1
                TipoD_Actual = ch

                Do While TipoD = TipoD_Actual And EsDocumento And Not Termino
                    CantD = CantD + 1

                    I = I + 1
                    If Not (I > Len(Aux_Cuenta)) Then
                        ch = UCase(Mid(Aux_Cuenta, I, 1))
                    Else
                        Termino = True
                    End If

                    If ch = "D" Then
                        'leo el nro de documento
                        I = I + 1
                        ch = UCase(Mid(Aux_Cuenta, I, 2))
                        'Avanzo otro porque lei 2
                        I = I + 1
                        TipoD_Actual = ch
                    Else
                        Termino = True
                    End If
                Loop

            Else
                EsDocumento = False
            End If
            
            

            'Busco el documento para reemplazar
            DescDocumento = "00000000000000000000"
            If IsNumeric(TipoD) Then
                StrSql = " SELECT nrodoc " & _
                         " From ter_doc " & _
                         " Where ter_doc.ternro = " & Ternro & " And ter_doc.tidnro = " & CLng(TipoD)
                OpenRecordset StrSql, rs_Documento
                If Not rs_Documento.EOF Then
                        DescDocumento = IIf(IsNull(rs_Documento!NroDoc), "00000000000000000000", rs_Documento!NroDoc & "00000000000000000000")
                Else
                    DescDocumento = "00000000000000000000"
                End If
                rs_Documento.Close
            End If
            DescDocumento = Left(DescDocumento, 20)
            
            'Reemplazo en la cuenta
            NroCuenta = NroCuenta & Mid(DescDocumento, 1, CantD)

            TipoD = TipoD_Actual
            CantD = 1
        
        Loop
    
     
    Case "a" To "z", "A" To "Z":
        NroCuenta = NroCuenta & ch
        I = I + 1
    Case Else:
        I = I + 1
    End Select
Loop


'cierro todo
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Documento.State = adStateOpen Then rs_Documento.Close
Set rs_Estructura = Nothing
Set rs_Documento = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_ArmarCuenta:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: ArmarCuenta"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub


Public Sub ArmarCuentaTarja(ByRef NroCuenta As String, ByVal Ternro As Long, ByVal Legajo As String, ByVal Masinivternro1 As Long, ByVal Masinivternro2 As Long, ByVal Masinivternro3 As Long, ByVal Estructura1 As Long, ByVal Estructura2 As Long, ByVal Estructura3 As Long, ByVal t_tarja As Boolean, ByVal FechaAsiento As Date, ByVal linea_DH As Integer, ByVal adctaasiento As String, ByVal adproyecto As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Arma la cuenta de acuerdo a la configuracion de la misma
' Autor      : Fernando Favre
' Fecha      : 25/04/2007
' --------------------------------------------------------------------------------------------
Dim Aux_Cuenta As String
Dim Aux_Legajo As String

Dim DescEstructura1 As String
Dim DescEstructura2 As String
Dim DescEstructura3 As String

Dim I As Integer
Dim ch As String
Dim CantL As Integer
Dim CantE As Integer
Dim TipoE As String
Dim TipoE_Actual As String
Dim EsEstructura As Boolean
Dim CantG As Integer
Dim TipoG As String
Dim TipoG_Actual As String
Dim EsEscala As Boolean
Dim Termino As Boolean

Dim PosE1 As Integer
Dim PosE2 As Integer
Dim PosE3 As Integer
Dim PosG1 As Integer

Dim EsDocumento As Boolean
Dim CantD As Long
Dim TipoD As String
Dim TipoD_Actual As String
Dim DescDocumento As String

Dim fec_desde_escala As Date
Dim fec_hasta_escala As Date
Dim valor_escala As String
Dim valor_escala_aux As String
Dim Seguir As Boolean

Dim rs_Estructura As New ADODB.Recordset
Dim rs_Documento As New ADODB.Recordset
Dim rs_escala As New ADODB.Recordset

On Error GoTo ME_ArmarCuenta

Aux_Cuenta = NroCuenta
Aux_Legajo = Legajo

'Si la cuenta es del DEBE y hay tarja, se toma la cuenta de la tarja, sino se arma
If t_tarja And linea_DH = 0 Then
    ' La tarja trae las cuentas YA armdas
    ' 0 Debe(+), 1 Haber(-), 2 Variable(+/-), 3 Variable invertido(-/+)
''    i = 0
''    usaVecConcepto = False
''    Do While i < max_ind_con And Not usaVecConcepto
''        If vec_con(i).concnro = 1 Then
''            usaVecConcepto = False
''        Else
''            i = i + 1
''        End If
''    Loop
    
''    If usaVecConcepto Then
        If Mid(adctaasiento, 4, 3) = "100" And Aux_Cuenta = "G1G1G1130E1E1E1E2E2" Then
            NroCuenta = Mid(adctaasiento, 1, 3) & "130" & Mid(adctaasiento, 7, 5)
        Else
            NroCuenta = adctaasiento
        End If
        
        If Mid(adctaasiento, 1, 3) = "141" Then
            NroCuenta = NroCuenta & adproyecto
        End If
        
''    Else
''        If Mid(adctaasiento, 4, 3) = "100" And Mid(Aux_Cuenta, 4, 3) = "E1E1E1130E1E1E1E1E1" Then
''            NroCuenta = Mid(vec_con(i).cuenta, 1, 3) & "130" & Mid(vec_con(i).cuenta, 7, 3)
''        Else
''            NroCuenta = vec_con(i).cuenta
''        End If
''    End If
    
    
'     if se utilizo el vector de concepto then
    '    /* TARJA con IMPUTACION DIRECTA (CONCEPTO DE HORAS) */
'        If Mid(adctaasiento, 4, 3) = "100" And Aux_Cuenta = "G1G1G1130E1E1E1E1E1" Then
'            NroCuenta = Mid(adctaasiento, 1, 3) & "130" & Mid(adctaasiento, 7, 3)
'        Else
'            NroCuenta = adctaasiento
'        End If
'     Else
'        If Mid(adctaasiento, 4, 3) = "100" And Mid(Aux_Cuenta, 4, 3) = "E1E1E1130E1E1E1E1E1" Then
'            NroCuenta = Mid(adctaasiento, 1, 3) & "130" & Mid(adctaasiento, 7, 3)
'        Else
'            NroCuenta = adctaasiento
'        End If
'     End If
'          IF NOT {4} THEN DO:
'              IF (SUBSTRING(vec_con.cuenta,4,3) = "100" AND mod_linea.lin_cuenta="***130*****")
'                 THEN ASSIGN nro_cuenta = SUBSTRING(vec_con.cuenta,1,3) + "130" + SUBSTRING(vec_con.cuenta,7,5).
'                 ELSE ASSIGN nro_cuenta = vec_con.cuenta.
'              ASSIGN vec_pro1 = vec_con.proyecto
'                     vec_cc1 = "0"
'                     vec_act1 = "0"
'                     vec_jor1 = vec_con.canthoras.
'          END.
'          Else: Do:
'              IF (SUBSTRING(vec_cta[inx],4,3) = "100" AND mod_linea.lin_cuenta="***130*****")
'                 THEN ASSIGN nro_cuenta = SUBSTRING(vec_cta[inx],1,3) + "130" + SUBSTRING(vec_cta[inx],7,5).
'                 ELSE ASSIGN nro_cuenta = vec_cta[inx].
'             ASSIGN vec_cc1 =  vec_cc [inx]
'                    vec_act1 = vec_act[inx]
'                    vec_pro1 = vec_pro[inx]
'                    vec_jor1 = vec_jor[inx].
'          END.
Else
    
    'Descripcion de estructura 1
    If Masinivternro1 = 0 Then
        'Modelo sin apertura
        DescEstructura1 = "00000000000000000000"
    Else
        'Modelo con apertura, busco la descripcion de la estructura
        StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
                 " WHERE estrnro = " & Estructura1
        OpenRecordset StrSql, rs_Estructura
        
        If Not rs_Estructura.EOF Then
                DescEstructura1 = IIf(IsNull(rs_Estructura!estrcodext), "00000000000000000000", rs_Estructura!estrcodext & "00000000000000000000")
        Else
            DescEstructura1 = "00000000000000000000"
        End If
        rs_Estructura.Close
    End If
    DescEstructura1 = Left(DescEstructura1, 20)
    
    'Descripcion de estructura 2
    If Masinivternro2 = 0 Then
        'Modelo sin apertura
        DescEstructura2 = "00000000000000000000"
    Else
        'Modelo con apertura, busco la descripcion de la estructura
        StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
                 " WHERE estrnro = " & Estructura2
        OpenRecordset StrSql, rs_Estructura
        
        If Not rs_Estructura.EOF Then
                DescEstructura2 = IIf(IsNull(rs_Estructura!estrcodext), "00000000000000000000", rs_Estructura!estrcodext & "00000000000000000000")
        Else
            DescEstructura2 = "00000000000000000000"
        End If
        rs_Estructura.Close
    End If
    DescEstructura2 = Left(DescEstructura2, 20)
    
    'Descripcion de estructura 3
    If Masinivternro3 = 0 Then
        'Modelo sin apertura
        DescEstructura3 = "00000000000000000000"
    Else
        'Modelo con apertura, busco la descripcion de la estructura
        StrSql = " SELECT estrcodext, estrnro, tenro FROM estructura " & _
                 " WHERE estrnro = " & Estructura3
        OpenRecordset StrSql, rs_Estructura
        
        If Not rs_Estructura.EOF Then
                DescEstructura3 = IIf(IsNull(rs_Estructura!estrcodext), "00000000000000000000", rs_Estructura!estrcodext & "00000000000000000000")
        Else
            DescEstructura3 = "00000000000000000000"
        End If
        rs_Estructura.Close
    End If
    DescEstructura3 = Left(DescEstructura3, 20)
    
    ' Descripcion de la cuenta en la escala
    If Masinivternro1 <> 0 And Masinivternro1 <> 0 Then
        StrSql = " SELECT vgrvalor, vgrorden FROM valgrilla WHERE cgrnro=60"
        StrSql = StrSql & " AND vgrcoor_1 = " & Estructura1
        StrSql = StrSql & " AND vgrcoor_2 = " & Estructura2
        StrSql = StrSql & " ORDER BY vgrcoor_4, vgrorden "
        OpenRecordset StrSql, rs_escala
        valor_escala = "00000000000000000000"
        Seguir = True
        Do While Not rs_escala.EOF And Seguir
            If rs_escala!vgrorden = 2 Then
                fec_desde_escala = CDate("01/" & Format(rs_escala!vgrvalor, "00") & "/" & Year(FechaAsiento))
            End If
            If rs_escala!vgrorden = 3 Then
                If CInt(rs_escala!vgrvalor) = 12 Then
                    fec_hasta_escala = CDate("31/12" & "/" & Year(FechaAsiento))
                Else
                    fec_hasta_escala = CDate("01/" & Format(CInt(rs_escala!vgrvalor) + 1, "00") & "/" & Year(FechaAsiento))
                    fec_hasta_escala = DateAdd("d", -1, fec_hasta_escala)
                End If
                If fec_desde_escala > fec_hasta_escala Then
                    fec_desde_escala = DateAdd("m", -12, fec_desde_escala)
                End If
            End If
            If rs_escala!vgrorden = 1 Then
                valor_escala_aux = rs_escala!vgrvalor
            End If
            
            If fec_desde_escala <= FechaAsiento And FechaAsiento <= fec_hasta_escala Then
                Seguir = False
                valor_escala = valor_escala_aux
            End If
            rs_escala.MoveNext
        Loop
        rs_escala.Close
    Else
        valor_escala = "00000000000000000000"
    End If
                
                
    PosE1 = 1
    PosE2 = 1
    PosE3 = 1
    PosG1 = 1
    
    
    'Voy recorriendo de Izquierda a Derecha el aux_cuenta y voy generando el NroCuenta
    I = 1
    NroCuenta = ""
    CantL = 0
    CantE = 0
    CantD = 0
    CantG = 0
    Termino = False
    
    Do While Not (I > Len(Aux_Cuenta))
        ch = UCase(Mid(Aux_Cuenta, I, 1))
    
        Select Case ch
        Case "_", "-", ".", "*":
            NroCuenta = NroCuenta & ch
            I = I + 1
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9":
            NroCuenta = NroCuenta & ch
            I = I + 1
        Case "E": 'Estrcutura
            EsEstructura = True
            CantE = 1
            'leo el nro de la estructura
            I = I + 1
            ch = UCase(Mid(Aux_Cuenta, I, 1))
            TipoE = ch
            Termino = False
            
            Do While EsEstructura And Not Termino
                'leo el siguiente
                I = I + 1
                If Not (I > Len(Aux_Cuenta)) Then
                    ch = UCase(Mid(Aux_Cuenta, I, 1))
                Else
                    Termino = True
                End If
                
                If ch = "E" And Not Termino Then
                    'leo lel nro de la estructura
                    I = I + 1
                    ch = UCase(Mid(Aux_Cuenta, I, 1))
                    TipoE_Actual = ch
                    
                    Do While TipoE = TipoE_Actual And EsEstructura And Not Termino
                        CantE = CantE + 1
        
                        I = I + 1
                        If Not (I > Len(Aux_Cuenta)) Then
                            ch = UCase(Mid(Aux_Cuenta, I, 1))
                        Else
                            Termino = True
                        End If
                        
                        If ch = "E" Then
                            'leo el nro de la estructura
                            I = I + 1
                            ch = UCase(Mid(Aux_Cuenta, I, 1))
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
                    NroCuenta = NroCuenta & Mid(DescEstructura1, PosE1, CantE)
                    PosE1 = PosE1 + CantE
                    If PosE1 >= 20 Then PosE1 = 1
                Case 2:
                    'NroCuenta = NroCuenta & Right(Estructura2, CantE)
                    NroCuenta = NroCuenta & Mid(DescEstructura2, PosE2, CantE)
                    PosE2 = PosE2 + CantE
                    If PosE2 >= 20 Then PosE2 = 1
                Case 3:
                    'NroCuenta = NroCuenta & Right(Estructura3, CantE)
                    NroCuenta = NroCuenta & Mid(DescEstructura3, PosE3, CantE)
                    PosE3 = PosE3 + CantE
                    If PosE3 >= 20 Then PosE3 = 1
                End Select
                
                TipoE = TipoE_Actual
                CantE = 1
            Loop
            
        Case "L": 'Legajo
            Termino = False
            CantL = 1
            I = I + 1
            If I <= Len(Aux_Cuenta) Then
                ch = UCase(Mid(Aux_Cuenta, I, 1))
            End If
            
            Do While ch = "L" And Not Termino
                CantL = CantL + 1
                I = I + 1
                If I <= Len(Aux_Cuenta) Then
                    ch = UCase(Mid(Aux_Cuenta, I, 1))
                Else
                    Termino = True
                End If
            Loop
            
            NroCuenta = NroCuenta & Right(Format(Aux_Legajo, "0000000000"), CantL)
         
      
    
        Case "D": 'DXX NUMERO DE TIPO DE DOC, DONDE XX ES EL TIPO DE DOC
            EsDocumento = True
            CantD = 1
            'leo el nro de documento
            I = I + 1
            ch = UCase(Mid(Aux_Cuenta, I, 2))
            'Avanzo otro porque lei 2
            I = I + 1
            TipoD = ch
            Termino = False
    
            Do While EsDocumento And Not Termino
                'leo el siguiente
                I = I + 1
                If Not (I > Len(Aux_Cuenta)) Then
                    ch = UCase(Mid(Aux_Cuenta, I, 1))
                Else
                    Termino = True
                End If
    
                If ch = "D" And Not Termino Then
                    'leo lel nro de documento
                    I = I + 1
                    ch = UCase(Mid(Aux_Cuenta, I, 2))
                    'Avanzo otro porque lei 2
                    I = I + 1
                    TipoD_Actual = ch
    
                    Do While TipoD = TipoD_Actual And EsDocumento And Not Termino
                        CantD = CantD + 1
    
                        I = I + 1
                        If Not (I > Len(Aux_Cuenta)) Then
                            ch = UCase(Mid(Aux_Cuenta, I, 1))
                        Else
                            Termino = True
                        End If
    
                        If ch = "D" Then
                            'leo el nro de documento
                            I = I + 1
                            ch = UCase(Mid(Aux_Cuenta, I, 2))
                            'Avanzo otro porque lei 2
                            I = I + 1
                            TipoD_Actual = ch
                        Else
                            Termino = True
                        End If
                    Loop
    
                Else
                    EsDocumento = False
                End If
                
                
    
                'Busco el documento para reemplazar
                DescDocumento = "00000000000000000000"
                If IsNumeric(TipoD) Then
                    StrSql = " SELECT nrodoc " & _
                             " From ter_doc " & _
                             " Where ter_doc.ternro = " & Ternro & " And ter_doc.tidnro = " & CLng(TipoD)
                    OpenRecordset StrSql, rs_Documento
                    If Not rs_Documento.EOF Then
                            DescDocumento = IIf(IsNull(rs_Documento!NroDoc), "00000000000000000000", rs_Documento!NroDoc & "00000000000000000000")
                    Else
                        DescDocumento = "00000000000000000000"
                    End If
                    rs_Documento.Close
                End If
                DescDocumento = Left(DescDocumento, 20)
                
                'Reemplazo en la cuenta
                NroCuenta = NroCuenta & Mid(DescDocumento, 1, CantD)
    
                TipoD = TipoD_Actual
                CantD = 1
            
            Loop
        
        Case "G": 'Grilla - Escala
            EsEscala = True
            CantG = 1
            'leo el nro de la escala
            I = I + 1
            ch = UCase(Mid(Aux_Cuenta, I, 1))
            TipoG = ch
            Termino = False
            
            Do While EsEscala And Not Termino
                'leo el siguiente
                I = I + 1
                If Not (I > Len(Aux_Cuenta)) Then
                    ch = UCase(Mid(Aux_Cuenta, I, 1))
                Else
                    Termino = True
                End If
                
                If ch = "G" And Not Termino Then
                    'leo lel nro de la escala
                    I = I + 1
                    ch = UCase(Mid(Aux_Cuenta, I, 1))
                    TipoG_Actual = ch
                    
                    Do While TipoG = TipoG_Actual And EsEscala And Not Termino
                        CantG = CantG + 1
        
                        I = I + 1
                        If Not (I > Len(Aux_Cuenta)) Then
                            ch = UCase(Mid(Aux_Cuenta, I, 1))
                        Else
                            Termino = True
                        End If
                        
                        If ch = "G" Then
                            'leo el nro de la escala
                            I = I + 1
                            ch = UCase(Mid(Aux_Cuenta, I, 1))
                            TipoG_Actual = ch
                        Else
                            Termino = True
                        End If
                    Loop
                    
                Else
                    EsEscala = False
                End If
                
                'reemplazo por el valor de la escala
                Select Case TipoG
                Case 1:
                    NroCuenta = NroCuenta & Mid(valor_escala, PosG1, CantG)
                    PosG1 = PosG1 + CantG
                    If PosG1 >= 20 Then PosG1 = 1
                End Select
                
                TipoG = TipoG_Actual
                CantG = 1
            Loop
                
        Case "a" To "z", "A" To "Z":
            NroCuenta = NroCuenta & ch
            I = I + 1
        Case Else:
            I = I + 1
        End Select
    Loop
    
    If Mid(NroCuenta, 1, 3) = "141" Then
        NroCuenta = NroCuenta & adproyecto
    End If

'    NroCuenta = NroCuenta & adproyecto
    
'    If Mid(Aux_Cuenta, 1, 6) = "E1E1E1" And substring(Aux_Cuenta, 10, 10) = "E2E2E2E2E2E2" Then
        ' armado de la escala de desicion + CCosto + Actividad
        
'    Else
'        NroCuenta = Mid(Aux_Cuenta, 1, 6) & cod_sucursal("999") + Mid(Aux_Cuenta, 10, 2)
'    End If
End If


'cierro todo
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Documento.State = adStateOpen Then rs_Documento.Close
Set rs_Estructura = Nothing
Set rs_Documento = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_ArmarCuenta:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: ArmarCuenta"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql
'Resume Next
End Sub

Public Sub BalanceEmpleado(ByRef balanceOK As Boolean)
Dim indice As Long
Dim montoDebe As Double
Dim montoHaber As Double
Dim DebeHaber As Integer
    
On Error GoTo ME_BalanceEmpleado

    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "BALANCE EMPLEADO"

    'Inicializo variables de balance
    balanceOK = True
    montoDebe = 0
    montoHaber = 0
    
    'Determino el debe y haber de la linea
    For indice = 1 To ind_lineaAsiAux
    
        Flog.writeline Espacios(Tabulador * 2) & "CUENTA " & lineaAsiAux(indice).cuenta
               
        Select Case lineaAsiAux(indice).dh
            Case 0: 'Debe
                DebeHaber = -1
                montoDebe = montoDebe + Abs(lineaAsiAux(indice).Monto)
                Flog.writeline Espacios(Tabulador * 3) & "MONTO = " & Round(lineaAsiAux(indice).Monto, 4) & " DEBE"
            Case 1: 'Haber
                DebeHaber = 0
                montoHaber = montoHaber + Abs(lineaAsiAux(indice).Monto)
                Flog.writeline Espacios(Tabulador * 3) & "MONTO = " & Round(lineaAsiAux(indice).Monto, 4) & " HABER"
            Case 2: 'Variable
                If lineaAsiAux(indice).Monto >= 0 Then
                    DebeHaber = -1
                    montoDebe = montoDebe + Abs(lineaAsiAux(indice).Monto)
                    Flog.writeline Espacios(Tabulador * 3) & "MONTO = " & Round(lineaAsiAux(indice).Monto, 4) & " VARIABLE, SE RESUELVE EN DEBE"
                Else
                    DebeHaber = 0
                    montoHaber = montoHaber + Abs(lineaAsiAux(indice).Monto)
                    Flog.writeline Espacios(Tabulador * 3) & "MONTO = " & Round(lineaAsiAux(indice).Monto, 4) & " VARIABLE, SE RESUELVE EN HABER"
                End If
            Case 3: 'Variable Invertida
                If lineaAsiAux(indice).Monto >= 0 Then
                    DebeHaber = 0
                    montoHaber = montoHaber + Abs(lineaAsiAux(indice).Monto)
                    Flog.writeline Espacios(Tabulador * 3) & "MONTO = " & Round(lineaAsiAux(indice).Monto, 4) & " VARIABLE INVERTIDA, SE RESUELVE EN HABER"
                Else
                    DebeHaber = -1
                    montoDebe = montoDebe + Abs(lineaAsiAux(indice).Monto)
                    Flog.writeline Espacios(Tabulador * 3) & "MONTO = " & Round(lineaAsiAux(indice).Monto, 4) & " VARIABLE INVERTIDA, SE RESUELVE EN DEBE"
                End If
            Case Else
                Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontro debe haber de la linea"
                Exit Sub
        End Select
        
    Next
    
    Flog.writeline Espacios(Tabulador * 2) & "MONTO DEBE = " & Round(montoDebe, 4) & " MONTO HABER = " & Round(montoHaber, 4)
    
    If Round(montoDebe, 4) <> Round(montoHaber, 4) Then
        Flog.writeline Espacios(Tabulador * 2) & "NO BALANCEA"
        balanceOK = False
    Else
        Flog.writeline Espacios(Tabulador * 2) & "BALANCEA"
        balanceOK = True
    End If
    

Exit Sub
'Manejador de Errores del procedimiento
ME_BalanceEmpleado:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: BalanceEmpleado"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub
    

Public Sub BalanceModelo(ByVal vol_cod As Long, ByVal masinro As Long)

Dim rs_lineaAsi As New ADODB.Recordset
Dim rs_modeloLinea As New ADODB.Recordset
Dim rs_asiento As New ADODB.Recordset

Dim DebeHaber As Integer
Dim montoDebe As Double
Dim montoHaber As Double
Dim montoLinea As Double
Dim cantLineas As Long
    
On Error GoTo ME_BalanceModelo
    
    Flog.writeline
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "BALANCEO MODELO " & masinro
    
    '-------------------------------------------------------------------------------
    'Busco todas las lineas insertadas en el modelo para asignar D/H
    '-------------------------------------------------------------------------------
    StrSql = " SELECT * FROM linea_asi " & _
             " WHERE linea_asi.masinro = " & masinro & _
             " AND linea_asi.vol_cod =" & vol_cod
    OpenRecordset StrSql, rs_lineaAsi
    
    montoDebe = 0
    montoHaber = 0
    cantLineas = 0
    Do While Not rs_lineaAsi.EOF
        
        'Busco la configuracion de la linea del modelo para ver si es debe o haber
        StrSql = "SELECT * FROM mod_linea " & _
                 " WHERE mod_linea.masinro = " & rs_lineaAsi!masinro & _
                 " AND mod_linea.linaorden =" & rs_lineaAsi!Linea & _
                 " ORDER BY masinro,linaorden"
        OpenRecordset StrSql, rs_modeloLinea
        If Not rs_modeloLinea.EOF Then
            
            Select Case rs_modeloLinea!linaD_H
                Case 0: 'Debe
                    DebeHaber = -1
                    montoLinea = Abs(rs_lineaAsi!Monto)
                    montoDebe = montoDebe + montoLinea
                    Flog.writeline Espacios(Tabulador * 1) & "CUENTA " & rs_lineaAsi!cuenta & " MONTO = " & rs_lineaAsi!Monto & " DEBE"
                Case 1: 'Haber
                    DebeHaber = 0
                    montoLinea = Abs(rs_lineaAsi!Monto)
                    montoHaber = montoHaber + montoLinea
                    Flog.writeline Espacios(Tabulador * 1) & "CUENTA " & rs_lineaAsi!cuenta & " MONTO = " & rs_lineaAsi!Monto & " HABER"
                Case 2: 'Variable
                    If rs_lineaAsi!Monto >= 0 Then
                        DebeHaber = -1
                        montoLinea = Abs(rs_lineaAsi!Monto)
                        montoDebe = montoDebe + montoLinea
                        Flog.writeline Espacios(Tabulador * 1) & "CUENTA " & rs_lineaAsi!cuenta & " MONTO = " & rs_lineaAsi!Monto & " VARIABLE, SE RESUELVE EN DEBE"
                    Else
                        DebeHaber = 0
                        montoLinea = Abs(rs_lineaAsi!Monto)
                        montoHaber = montoHaber + montoLinea
                        Flog.writeline Espacios(Tabulador * 1) & "CUENTA " & rs_lineaAsi!cuenta & " MONTO = " & rs_lineaAsi!Monto & " VARIABLE, SE RESUELVE EN HABER"
                    End If
                Case 3: 'Variable Invertida
                    If rs_lineaAsi!Monto >= 0 Then
                        DebeHaber = 0
                        montoLinea = Abs(rs_lineaAsi!Monto)
                        montoHaber = montoHaber + montoLinea
                        Flog.writeline Espacios(Tabulador * 1) & "CUENTA " & rs_lineaAsi!cuenta & " MONTO = " & rs_lineaAsi!Monto & " VARIABLE INVERTIDA, SE RESUELVE EN HABER"
                    Else
                        DebeHaber = -1
                        montoLinea = Abs(rs_lineaAsi!Monto)
                        montoDebe = montoDebe + montoLinea
                        Flog.writeline Espacios(Tabulador * 1) & "CUENTA " & rs_lineaAsi!cuenta & " MONTO = " & rs_lineaAsi!Monto & " VARIABLE INVERTIDA, SE RESUELVE EN DEBE"
                    End If
                Case Else
                    Flog.writeline Espacios(Tabulador * 0) & "ERROR. No se encontro debe haber de la linea"
                    Exit Sub
            End Select
        
            'Actualizo la linea con el valor debe/haber calculado y el valor absoluto del monto
            StrSql = "UPDATE linea_asi SET dh = " & DebeHaber & _
                     ",monto =" & Round(montoLinea, 4) & _
                     " WHERE masinro = " & rs_lineaAsi!masinro & _
                     " AND vol_cod =" & rs_lineaAsi!vol_cod & _
                     " AND cuenta ='" & rs_lineaAsi!cuenta & "'"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            cantLineas = cantLineas + 1
        
        End If
        rs_modeloLinea.Close
        
        rs_lineaAsi.MoveNext
        
    Loop
    
    
    Flog.writeline Espacios(Tabulador * 0) & "MONTO DEBE = " & Round(montoDebe, 4) & " MONTO HABER = " & Round(montoHaber, 4)
    
    '-------------------------------------------------------------------------------
    'Si esta desbalanceado miro si hay cuenta nivelado
    '-------------------------------------------------------------------------------
    If Round(montoDebe, 4) <> Round(montoHaber, 4) Then
            
        'Busco cuenta niveladora
        StrSql = "SELECT * FROM mod_linea " & _
                 " WHERE masinro = " & masinro & _
                 " AND upper(linadesc) = 'NIVELADORA'"
        OpenRecordset StrSql, rs_modeloLinea
        
        If Not rs_modeloLinea.EOF Then
            Flog.writeline Espacios(Tabulador * 0) & "Cuenta niveladora: " & rs_modeloLinea!LinaOrden & " - " & rs_modeloLinea!linadesc & " Cuenta: " & rs_modeloLinea!linacuenta
            
            'Calculo la diferencia a insertar en la cuenta niveladora
            If montoDebe > montoHaber Then
                DebeHaber = 0
                montoLinea = montoDebe - montoHaber
                Flog.writeline Espacios(Tabulador * 1) & "Cuenta niveladora HABER, Monto = " & montoLinea
                montoHaber = montoHaber + montoLinea
            Else
                DebeHaber = -1
                montoLinea = montoHaber - montoDebe
                Flog.writeline Espacios(Tabulador * 1) & "Cuenta niveladora DEBE, Monto = " & montoLinea
                montoDebe = montoDebe + montoLinea
            End If
            
            'inserto la niveladora
            StrSql = "INSERT INTO linea_asi (cuenta,vol_cod,masinro,linea,desclinea,monto,dh)" & _
                     " VALUES ('" & rs_modeloLinea!linacuenta & _
                     "'," & vol_cod & _
                     "," & masinro & _
                     "," & rs_modeloLinea!LinaOrden & _
                     ",'" & rs_modeloLinea!linadesc & _
                     "'," & Round(montoLinea, 4) & _
                     "," & DebeHaber & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            cantLineas = cantLineas + 1
        Else
            Flog.writeline Espacios(Tabulador * 0) & "No hay cuenta niveladora"
        End If
        rs_modeloLinea.Close
        
    End If
    
    
    '-------------------------------------------------------------------------------
    'Creo el asiento
    '-------------------------------------------------------------------------------
    StrSql = "SELECT * FROM asiento " & _
             " WHERE masinro = " & masinro & _
             " AND vol_cod = " & vol_cod
    OpenRecordset StrSql, rs_asiento
    
    If rs_asiento.EOF Then
        StrSql = "INSERT INTO asiento (masinro,asidebe,asihaber,vol_cod) " & _
                 " VALUES (" & masinro & _
                 "," & Round(montoDebe, 4) & _
                 "," & Round(montoHaber, 4) & _
                 "," & vol_cod & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        StrSql = "UPDATE asiento SET asidebe = " & Round(montoDebe, 4) & _
                 ",asihaber =" & Round(montoHaber, 4) & _
                 " WHERE masinro = " & masinro & _
                 " AND vol_cod =" & vol_cod
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    rs_asiento.Close
    
    Flog.writeline Espacios(Tabulador * 0) & "Cantidad de lineas = " & cantLineas
    Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------------------------"
    Flog.writeline Espacios(Tabulador * 0) & "----------------------------------------------------------------"

If rs_lineaAsi.State = adStateOpen Then rs_lineaAsi.Close
If rs_modeloLinea.State = adStateOpen Then rs_modeloLinea.Close
If rs_asiento.State = adStateOpen Then rs_asiento.Close
Set rs_lineaAsi = Nothing
Set rs_modeloLinea = Nothing
Set rs_asiento = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_BalanceModelo:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: BalanceModelo"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql


End Sub


Public Sub BorrarDetalleAsiAux()
    ind_detalleAsiAux = 0
End Sub


Public Sub InsertarDetalleAsiAux(ByVal masinro As Long, ByVal vol_cod As Long, ByRef cuenta As String, ByVal lin_orden As Long, ByVal empleg As Long, ByRef terape As String, ByRef dldescripcion As String, ByVal dlcantidad As Double, ByVal dlmonto As Double, ByVal Ternro As Long, ByVal Origen As Long, ByVal tipoorigen As Long, ByVal linadesc As String, ByVal linaD_H As Integer)
    
On Error GoTo ME_InsertarDetalleAsiAux

    ind_detalleAsiAux = ind_detalleAsiAux + 1
    
    detalleAsiAux(ind_detalleAsiAux).masinro = masinro
    detalleAsiAux(ind_detalleAsiAux).vol_cod = vol_cod
    detalleAsiAux(ind_detalleAsiAux).cuenta = cuenta
    detalleAsiAux(ind_detalleAsiAux).lin_orden = lin_orden
    detalleAsiAux(ind_detalleAsiAux).empleg = empleg
    detalleAsiAux(ind_detalleAsiAux).terape = terape
    detalleAsiAux(ind_detalleAsiAux).dldescripcion = dldescripcion
    detalleAsiAux(ind_detalleAsiAux).dlcantidad = dlcantidad
    detalleAsiAux(ind_detalleAsiAux).dlmonto = dlmonto
    detalleAsiAux(ind_detalleAsiAux).Ternro = Ternro
    detalleAsiAux(ind_detalleAsiAux).Origen = Origen
    detalleAsiAux(ind_detalleAsiAux).tipoorigen = tipoorigen
    detalleAsiAux(ind_detalleAsiAux).linadesc = linadesc
    detalleAsiAux(ind_detalleAsiAux).linaD_H = linaD_H

Exit Sub
'Manejador de Errores del procedimiento
ME_InsertarDetalleAsiAux:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: InsertarDetalleAsiAux"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub


Public Sub InsertarDetalleAsiAuxTarja(ByVal masinro As Long, ByVal vol_cod As Long, ByRef cuenta As String, ByVal lin_orden As Long, ByVal empleg As Long, ByRef terape As String, ByRef dldescripcion As String, ByVal dlcantidad As Double, ByVal dlmonto As Double, ByVal porcentaje As Double, ByVal Ternro As Long, ByVal Origen As Long, ByVal tipoorigen As Long, ByVal linadesc As String, ByVal linaD_H As Integer)
    
On Error GoTo ME_InsertarDetalleAsiAuxTarja

    ind_detalleAsiAux = ind_detalleAsiAux + 1
    
    detalleAsiAux(ind_detalleAsiAux).masinro = masinro
    detalleAsiAux(ind_detalleAsiAux).vol_cod = vol_cod
    detalleAsiAux(ind_detalleAsiAux).cuenta = cuenta
    detalleAsiAux(ind_detalleAsiAux).lin_orden = lin_orden
    detalleAsiAux(ind_detalleAsiAux).empleg = empleg
    detalleAsiAux(ind_detalleAsiAux).terape = terape
    detalleAsiAux(ind_detalleAsiAux).dldescripcion = dldescripcion
    detalleAsiAux(ind_detalleAsiAux).dlcantidad = dlcantidad
    detalleAsiAux(ind_detalleAsiAux).dlmonto = dlmonto
    detalleAsiAux(ind_detalleAsiAux).porcentaje = porcentaje
    detalleAsiAux(ind_detalleAsiAux).Ternro = Ternro
    detalleAsiAux(ind_detalleAsiAux).Origen = Origen
    detalleAsiAux(ind_detalleAsiAux).tipoorigen = tipoorigen
    detalleAsiAux(ind_detalleAsiAux).linadesc = linadesc
    detalleAsiAux(ind_detalleAsiAux).linaD_H = linaD_H

Exit Sub
'Manejador de Errores del procedimiento
ME_InsertarDetalleAsiAuxTarja:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: InsertarDetalleAsiAuxTarja"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub


Public Sub BorrarDetalleAsiAuxEmp()
    ind_detalleAsiAuxEmp = 0
End Sub

Public Sub InsertarDetalleAsiAuxEmp(ByVal masinro As Long, ByVal vol_cod As Long, ByRef cuenta As String, ByVal lin_orden As Long, ByVal empleg As Long, ByRef terape As String, ByRef dldescripcion As String, ByVal dlcantidad As Double, ByVal dlmonto As Double, ByVal dlmontoacum As Double, ByVal dlcosto1 As Long, ByVal dlcosto2 As Long, ByVal dlcosto3 As Long, ByVal dlcosto4 As Long, ByVal dlporcentaje As Double, ByVal Ternro As Long, ByVal Origen As Long, ByVal tipoorigen As Long, ByVal linadesc As String, ByVal linaD_H As Integer)

On Error GoTo ME_InsertarDetalleAsiAuxEmp

    ind_detalleAsiAuxEmp = ind_detalleAsiAuxEmp + 1

    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).masinro = masinro
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).vol_cod = vol_cod
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).cuenta = cuenta
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).lin_orden = lin_orden
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).empleg = empleg
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).terape = terape
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).dldescripcion = dldescripcion
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).dlcantidad = dlcantidad
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).dlmonto = dlmonto
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).dlmontoacum = dlmontoacum
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).dlcosto1 = dlcosto1
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).dlcosto2 = dlcosto2
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).dlcosto3 = dlcosto3
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).dlcosto4 = dlcosto4
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).dlporcentaje = dlporcentaje
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).Ternro = Ternro
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).Origen = Origen
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).tipoorigen = tipoorigen
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).linadesc = linadesc
    detalleAsiAuxEmp(ind_detalleAsiAuxEmp).linaD_H = linaD_H

Exit Sub
'Manejador de Errores del procedimiento
ME_InsertarDetalleAsiAuxEmp:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: InsertarDetalleAsiAuxEmp"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub

Public Sub ResolverDetalleAsi(ByVal vol_cod As Long, ByVal masinro As Long, ByRef cuenta As String, ByVal porcentaje As Double, ByVal estr1 As Long, ByVal estr2 As Long, ByVal estr3 As Long)

Dim rs_acumulado As New ADODB.Recordset
Dim indice As Long
Dim MontoAImputar As Double
Dim MontoAcumulado As Double

On Error GoTo ME_ResolverDetalleAsi

    'Recupero el monto ya acumulado
    StrSql = "SELECT sum(dlmonto) total FROM detalle_asi WHERE vol_cod = " & vol_cod & " AND masinro = " & masinro & " AND cuenta = '" & cuenta & "'"
    OpenRecordset StrSql, rs_acumulado
        MontoAcumulado = IIf(EsNulo(rs_acumulado!total), 0, rs_acumulado!total)
    rs_acumulado.Close
    
    
    For indice = 1 To ind_detalleAsiAux
        MontoAImputar = (detalleAsiAux(indice).dlmonto * porcentaje) / 100
        MontoAcumulado = MontoAcumulado + MontoAImputar
        Call InsertarDetalleAsiAuxEmp(detalleAsiAux(indice).masinro, detalleAsiAux(indice).vol_cod, cuenta, detalleAsiAux(indice).lin_orden, detalleAsiAux(indice).empleg, detalleAsiAux(indice).terape, detalleAsiAux(indice).dldescripcion, detalleAsiAux(indice).dlcantidad, MontoAImputar, MontoAcumulado, estr1, estr2, estr3, 0, porcentaje, detalleAsiAux(indice).Ternro, detalleAsiAux(indice).Origen, detalleAsiAux(indice).tipoorigen, detalleAsiAux(indice).linadesc, detalleAsiAux(indice).linaD_H)
    Next

If rs_acumulado.State = adStateOpen Then rs_acumulado.Close
Set rs_acumulado = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_ResolverDetalleAsi:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: ResolverDetalleAsi"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub

Public Sub ResolverDetalleAsiTarja(ByVal vol_cod As Long, ByVal masinro As Long, ByRef cuenta As String, ByVal porcentaje As Double, ByVal estr1 As Long, ByVal estr2 As Long, ByVal estr3 As Long)

Dim rs_acumulado As New ADODB.Recordset
Dim indice As Long
Dim MontoAImputar As Double
Dim MontoAcumulado As Double
Dim CantAImputar As Double

On Error GoTo ME_ResolverDetalleAsiTarja

    'Recupero el monto ya acumulado
    StrSql = "SELECT sum(dlmonto) total FROM detalle_asi WHERE vol_cod = " & vol_cod & " AND masinro = " & masinro & " AND cuenta = '" & cuenta & "'"
    OpenRecordset StrSql, rs_acumulado
        MontoAcumulado = IIf(EsNulo(rs_acumulado!total), 0, rs_acumulado!total)
    rs_acumulado.Close
    
    
    For indice = 1 To ind_detalleAsiAux
        MontoAImputar = (detalleAsiAux(indice).dlmonto * porcentaje) / 100
        CantAImputar = (detalleAsiAux(indice).dlcantidad * porcentaje) / 100
        MontoAcumulado = MontoAcumulado + MontoAImputar
        Call InsertarDetalleAsiAuxEmp(detalleAsiAux(indice).masinro, detalleAsiAux(indice).vol_cod, cuenta, detalleAsiAux(indice).lin_orden, detalleAsiAux(indice).empleg, detalleAsiAux(indice).terape, detalleAsiAux(indice).dldescripcion, CantAImputar, MontoAImputar, MontoAcumulado, estr1, estr2, estr3, 0, porcentaje, detalleAsiAux(indice).Ternro, detalleAsiAux(indice).Origen, detalleAsiAux(indice).tipoorigen, detalleAsiAux(indice).linadesc, detalleAsiAux(indice).linaD_H)
    Next

If rs_acumulado.State = adStateOpen Then rs_acumulado.Close
Set rs_acumulado = Nothing

Exit Sub
'Manejador de Errores del procedimiento
ME_ResolverDetalleAsiTarja:
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Procedimiento: ResolverDetalleAsiTarja"
    Flog.writeline "Ultimo SQL Ejecutado: " & StrSql

End Sub

