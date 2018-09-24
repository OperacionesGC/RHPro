Attribute VB_Name = "MdlNivGeren"
Option Explicit


'Const Version = "1.01"
'Const FechaVersion = "30/03/2009"
'Autor = Diego Nuñez
'   Version Inicial

Const Version = "1.02"
Const FechaVersion = "31/07/2009" 'Martin Ferraro - Se acomodo el log para Encriptacion de string connection


'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
Public Type TipoRestriccion
    Estrnro As Long
    Valor As Double
End Type

Global inx             As Integer
Global inxfin          As Integer

Global vec_testr1(50)  As Integer
Global vec_testr2(50)  As String
Global vec_testr3(50)  As String

Global vec_jor(50) As Double

Global Descripcion As String
Global Cantidad As Double

'Declaraciones de variables añadidas por Hernán D. Santonocito - 03/03/2006
Global nListaProc As String
Global nEmpresa As Long
'----------------------------------------------------------

Global rs_Proc_Vol As New ADODB.Recordset
Global rs_Mod_Linea As New ADODB.Recordset
Global rs_Empleado As New ADODB.Recordset
Global rs_Mod_Asiento As New ADODB.Recordset

Global BUF_mod_linea As New ADODB.Recordset
Global BUF_temp As New ADODB.Recordset
Global USA_DEBUG As Boolean

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte de Hs liquidadas.
' Autor      :
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
Dim iduser As String
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


    On Error GoTo ME_Main
    
    Nombre_Arch = PathFLog & "Reporte_NivGeren" & "-" & NroProcesoBatch & ".log"
    
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
    
'    On Error GoTo ME_Main
'
'    Nombre_Arch = PathFLog & "Reporte_NivGeren" & "-" & NroProcesoBatch & ".log"
'
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
'
'
'    ' Obtengo el Process ID
'    PID = GetCurrentProcessId
'    Flog.writeline "-------------------------------------------------"
'    Flog.writeline "Version                  : " & Version
'    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
'    Flog.writeline "PID                      : " & PID
'    Flog.writeline "-------------------------------------------------"
'    Flog.writeline
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 236 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        iduser = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call Calcular(NroProcesoBatch, bprcparam, iduser)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
        Progreso = 100
        'UpdateProgreso (Progreso)
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


Public Sub Calcular(ByVal bpronro As Long, ByVal Parametros As String, ByVal iduser As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte
' Autor      : FGZ
' Fecha      : 20/01/2004
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim Nroliq As Long
Dim Empresa As Long
Dim Empresa_Original As Long
Dim Lista_Pro As String
Dim Lista_Pro_F As String

Dim fechadesde As Date
Dim fechahasta As Date

Dim cant_dias_td         As Integer
Dim cant_dias_total      As Integer
Dim fecha_hasta_td       As Date
Dim fec_hasta_total      As Date
Dim tipdia_td            As Integer
Dim tipdia_total         As Integer
Dim nro_sitrev           As Integer
Dim nro_sitrev_total     As Integer

Dim asigno_sitrev    As Boolean
Dim Desde            As Date
Dim Hasta            As Date
Dim topeArreglo      As Integer   'USAR ESTA VARIABLE PARA EL TOPE
Dim arreglo(80)      As Double
Dim tope_ampo_max    As Double
Dim tope_ampo_min    As Double
Dim bruto            As Double
Dim imp_oso          As Double
Dim ent_oso          As Integer
Dim single_oso          As Integer
Dim imp_ss           As Double
Dim imp_ss_con       As Double
Dim ent_ss           As Integer
Dim single_ss           As Integer
Dim ent_ss_con       As Integer
Dim single_ss_con       As Integer
Dim ent_bruto        As Integer
Dim single_bruto        As Integer
Dim msr              As Double
Dim ent_msr          As Integer
Dim single_msr          As Integer
Dim ent_neto         As Integer
Dim single_neto         As Integer
Dim ent_afa          As Integer
Dim single_afa          As Integer
Dim cant_hijos       As Integer
Dim cant_cony        As Double
Dim adicional_os     As Double
Dim ent_aos          As Integer
Dim single_aos          As Integer
Dim ent_apor_vol     As Integer
Dim single_apor_vol     As Integer
Dim ent_exed_ss      As Integer
Dim single_exed_ss      As Integer
Dim ent_exed_os      As Integer
Dim single_exed_os      As Integer
Dim rebaja_promo     As Double
Dim ent_reb_pro      As Integer
Dim single_reb_pro      As Integer

Dim par_msr          As Long 'LIKE acumulador.acunro.
Dim par_asiflia      As Long 'LIKE acumulador.acunro.
Dim par_cuil         As Long 'LIKE tipodocu.tidnro.
Dim par_desvincula   As Long 'LIKE per.proceso.tprocnro.
Dim par_jubila       As Integer
Dim par_osocial      As String
Dim suma_osocial     As Boolean
Dim contador         As Integer
Dim desvinculado     As Boolean
Dim despedido        As Boolean
Dim X                As Boolean
Dim v_conyuge       As Integer
Dim v_canthijos      As Integer
Dim ent_apo_adi_os18 As Integer    'Version 18 r2 _ Marzo 2002
Dim single_apo_adi_os18 As Integer 'Version 18 r2 _ Marzo 2002
Dim ent_ap_adi_os    As Integer
Dim single_ap_adi_os    As Integer
Dim sitrev           As Integer

Dim escribir As Boolean
Dim cont_lic As Integer
Dim inimes As Date
Dim I      As Integer
Dim fecini As Date
Dim ult_lic  As Integer
Dim ult_fec  As Date
Dim Resto As Double

Dim pos1 As Integer
Dim pos2 As Integer

' auxiliares
Dim aux_fecha As Date
Dim Fecha_Inicio_periodo As Date
Dim Fecha_Fin_Periodo As Date
Dim UltimoEmpleado As Long

Dim Fecha_Inicio_Fase As Date
Dim Fecha_Fin_Fase As Date
Dim FechaAuxAsignada As Boolean 'Declarada por Hernán Santonocito - 03/03/2006

Dim rs_Empleados As New ADODB.Recordset
Dim rs_Acu_liq As New ADODB.Recordset
Dim rs_Proceso As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Confrep As New ADODB.Recordset
Dim rs_ConfrepAux As New ADODB.Recordset
Dim Aux_rs_Confrep As New ADODB.Recordset
Dim rs_Conceptos As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Impproarg As New ADODB.Recordset
Dim rs_Repsijp As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Cuil As New ADODB.Recordset
Dim rs_Familia As New ADODB.Recordset
Dim rs_Empresa As New ADODB.Recordset
Dim rs_TipoCont As New ADODB.Recordset
Dim rs_sucursal As New ADODB.Recordset
Dim rs_zona As New ADODB.Recordset
Dim rs_Estr_cod As New ADODB.Recordset
Dim rs_Convenios As New ADODB.Recordset
Dim rs_fases As New ADODB.Recordset
Dim rs_EmplFiltro As New ADODB.Recordset
Dim rs_Detalle As New ADODB.Recordset
Dim rs_ConfrepInd As New ADODB.Recordset
Dim rs_Consulta As New ADODB.Recordset
'Declaracion de objeto recordset agregada por Hernán D. Santonocito - 03/03/2006
Dim rs_HisEstructuras As ADODB.Recordset
'-------------------------------------------------------------------------------

Dim TipoEstr As Long
Dim Opcion As Integer


Dim Aux_ApeNom As String
Dim Aux_Cod_Cont As String
Dim Aux_Cod_sitr As String
Dim Aux_Cod_sitr1 As String
Dim Aux_Cod_sitr2 As String
Dim Aux_Cod_sitr3 As String
Dim Aux_diainisr1 As String
Dim Aux_diainisr2 As String
Dim Aux_diainisr3 As String
Dim Aux_Cod_Cond As String
Dim Aux_CUIL As String
Dim Aux_Reduccion As String
Dim Aux_Zona As String
Dim Aux_Localidad As String
Dim Aux_Cod_Obra_Social As String
Dim Aux_Contrato_Actual As String
Dim Aux_Actividad As String
Dim Aux_TipEmpNro As String
Dim Aux_Cod_Siniestro As String
Dim Aux_con_FNE As String
Dim Aux_Rem_Total As String
Dim Aux_Msr As String
Dim AUX_Imp_SS As String
Dim AUX_Imp_OS As String
Dim Aux_imp_ss_con As String
Dim Aux_Adi_OS As String
Dim Aux_Asig_Fliar As String
Dim Aux_Aporte_Voluntario As String
Dim Aux_Exc_SS As String
Dim Aux_Exc_OS As String
Dim Aux_RebajaPromovida As String
Dim Aux_Cant_Hijos As String
Dim Aux_Conyuges As String
Dim Aux_Adherentes As String
Dim Aux_Porc_Adi As String
Dim Aux_Apo_Adi_OS As String
Dim Aux_XConv As String
Dim Aux_imprem3 As String
Dim Aux_imprem4 As String
Dim Aux_Imprem5 As String
Dim Aux_DiasTrab As String
Dim Aux_Sue_Adic As String
Dim Aux_correspred As String
Dim Aux_caprecomlrt As String
Dim Aux_sac As String
Dim Aux_hrsextras As String
Dim Aux_zonadesf As String
Dim Aux_lar As String
Dim Aux_codsiniestro As String
Dim Aux_canthsext As String

Dim tcreduccion As Double
Dim Tiene_Convenio As Boolean

Dim Mensaje As String

Dim Restricciones(100) As TipoRestriccion
Dim LS_Res As Integer
Dim Aplica As Boolean
Dim Valor As Double

'FGZ - 23/08/2005
'Variable para el detallado del reporte
'se configuran en las columnas 60 en adelante
Dim Det_CantHDisc As Integer
Dim Det_Prenatal As Double
Dim Det_Escolar As Integer
Dim Det_PlusZonaDes As Double
Dim Det_Ret_Jubilacion As Double
Dim Det_Ret_Ley As Double
Dim Det_Ret_OS As Double
Dim Det_Total_Ret_OS As Double
Dim Det_Ret_Ansal As Double
Dim Det_Con_Jubilacion As Double
Dim Det_Con_Ley As Double
Dim Det_Con_FNE As Double
Dim Det_Con_AsigFliares As Double
Dim Det_Con_OS As Double
Dim Det_Con_Ansal As Double
Dim Det_Con_Total_OS As Double

Dim Edad As Integer             'Edad tope para considerar Escolar
Dim Estudia As Boolean          'Si se contempla si estudia o no
Dim Nivel_Estudio As Integer    'nivel de estudio considerado
Dim edad_F As Integer           'Edad del familiar
Dim Fam_Niv_Est As Long         'nivel de estuido del familiar
Dim Fam_Estudia As Boolean      'si el familiar estudia

Dim rs_Familiar As New ADODB.Recordset
Dim rs_Estudio_Actual As New ADODB.Recordset
Dim rs_Nivest As New ADODB.Recordset

Dim Aux_adicional As String
Dim Aux_premios As String
Dim Aux_remdec788 As String
Dim Aux_remimpo7 As String


'Inicializacion de Variables
cant_dias_td = 0
cant_dias_total = 0
fecha_hasta_td = CDate("01/01/1900")
fec_hasta_total = CDate("01/01/1900")
tipdia_td = 0
tipdia_total = 0
nro_sitrev = 0
nro_sitrev_total = -1
topeArreglo = 80
suma_osocial = False
Dim resto19 As Double
Dim sueldo As Double
Dim hrsextras As Double
Dim zonadesf As Double
Dim adicional As Double
Dim premios As Double
Dim sac As Double
Dim lar As Double

Dim legdesde As Long
Dim leghasta As Long
Dim estado As Integer
Dim anio As Integer
Dim perdesde As Integer
Dim perhasta As Integer
Dim tenro1 As String
Dim estrnro1 As String
Dim tenro2 As String
Dim estrnro2 As String
Dim tenro3 As String
Dim estrnro3 As String
Dim fecestr As String
Dim Orden As String
Dim ordenado As String
Dim ListadoASP As String
Dim Detalle() As Variant
Dim Evaluaciones(1 To 20, 1 To 20) As Variant
Dim Indicadores(1 To 20, 1 To 20) As Variant
Dim RegDetalle, ColDetalle As Integer
Dim MemoEmpleados As String
Dim MemoConfrep As String
Dim ColConfrep As Integer
Dim Salir, NoInd, Pase, NoEv As Boolean
Dim Memopliq() As Variant
Dim Memopliqmes() As Variant
Dim Memopliqanio() As Variant
Dim Memopliqdesde() As Variant
Dim Memopliqhasta() As Variant
Dim CantPeriodos As Integer
Dim Sumatoria, Prom, Min, Max As Double
Dim ColumnasConfrep() As Variant
Dim StrEncontrado As String
Dim LabelsRegistrosDetalle() As String
Dim LabelsRegistrosIndicadores() As String
Dim j, k As Integer
Dim MemoDotacion(20) As Integer
Dim NombrePeriodos() As String
Dim Memopliqdesc() As String
Dim NroFilasDetalle As Integer
Dim TipoEncontrado As String
Dim FormulaOriginal, StrEmpleado As String
Dim MesPeriodo, MesPeriodo2, anioAnt As Integer
Dim Concepto As Integer
Dim ConceptoFinal, AcumuladorFinal, NoEmp As Boolean
Dim fila, columna As Integer
Dim ValorF As Long
Dim ValoresF() As Long
Dim NombreF, Referencia As String
'' FGZ - 09/08/2004
'' Inicializo algunos valores para el arreglo
''Remuneracion 1
'arreglo(19) = 0.01
''Remuneracion 4
'arreglo(20) = 0.01
''Remuneracion 2 y Remuneracion 3
'arreglo(31) = 0.01
''Remuneracion 5
'arreglo(33) = 0.01
''Sueldo  + Adicional
'arreglo(34) = 0.01
Progreso = 0
IncPorc = 0
'FGZ - 08/11/2004 lo de arriba deja de tener vigencia
'Remuneracion 1
arreglo(19) = 0
'Remuneracion 4
arreglo(20) = 0
'Remuneracion 2 y Remuneracion 3
arreglo(31) = 0
'Remuneracion 5
arreglo(33) = 0
'Sueldo  + Adicional
arreglo(34) = 0


    ' Inicio codigo ejecutable
    On Error GoTo CE

' El formato del mismo es (nroliq, Todos_los_procesos, [pronro], proceso_aprbados)

' Levanto cada parametro por separado, el separador de parametros es "."
Flog.writeline "levantando parametros" & Parametros
If Not IsNull(Parametros) Then

    Flog.writeline "If Len(Parametros) >= 1 Then"
    
    If Len(Parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        
        Flog.writeline "pos1 =" & pos1 & " pos2=" & pos2
        
        legdesde = CLng(Mid(Parametros, pos1, pos2))
        'Nroliq = CLng(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro legdesde = " & legdesde
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        leghasta = CLng(Mid(Parametros, pos1, pos2))
        'Lista_Pro = Mid(Parametros, pos1, pos2 - pos1 + 1)
        'MAF - Estaba truncando en 100
        'Lista_Pro_F = Left(Lista_Pro, 1000)
        Flog.writeline "Parametro leghasta = " & leghasta
        ' esta lista tiene los nro de procesos separados por comas
        
        'Asigno el valor de lista de proceso a la variable global para poder usar en el SIJP
        'Agregado por Hernán D. Santonocito - 03/03/2006 -----------------------------------
        'nListaProc = Lista_Pro
        '-----------------------------------------------------------------------------------
        
'        pos1 = pos2 + 2
'        If Not Todos_Pro Then
'  '          pos1 = 1
'            pos2 = InStr(pos1, Parametros, ".") - 1
'            NroProc = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
'        Else
'            NroProc = 0
'        End If

'        pos1 = pos2 + 2
'        pos2 = InStr(pos1, Parametros, ".") - 1
'        Proc_Aprob = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        estado = CInt(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro estado = " & estado
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        anio = CInt(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro año = " & anio
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        perdesde = CInt(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro perdesde = " & perdesde
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        perhasta = CInt(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro perhasta = " & perhasta
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        tenro1 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro tenro1 = " & tenro1
       
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        estrnro1 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro estrnro1 = " & estrnro1
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        tenro2 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro tenro2 = " & tenro2
       
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        estrnro2 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro estrnro2 = " & estrnro2
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        tenro3 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro tenro3 = " & tenro3
       
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        estrnro3 = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro estrnro3 = " & estrnro3
        
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        fecestr = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro fecestr = " & fecestr
       
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        Orden = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro orden = " & Orden
        
        pos1 = pos1 + pos2 + 1
        pos2 = InStr(pos1, Parametros, ".") - 1
        pos2 = pos2 - pos1 + 1
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro ordenado =" & Mid(Parametros, pos1, pos2)
        
        ordenado = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro ordenado = " & ordenado
        
        pos1 = pos1 + pos2 + 1
        pos2 = Len(Parametros)
        ListadoASP = Mid(Parametros, pos1, pos2)
        Flog.writeline "Parametro de lista de empleados elegidos desde ASP = " & ListadoASP
        
    End If
Else
    Flog.writeline "parametros nulos"
End If
Flog.writeline "terminó de levantar los parametros"


'Me quedo con el mes del menor periodo de liquidación
RegDetalle = 0
StrSql = "SELECT pliqnro, pliqmes, pliqanio FROM periodo WHERE pliqanio = '" & anio & _
"' AND pliqnro = '" & perdesde & "'"
OpenRecordset StrSql, rs_Periodo
MesPeriodo = 0
If Not rs_Periodo.EOF Then MesPeriodo = rs_Periodo!pliqmes
rs_Periodo.Close

'Ahora busco el mes correspondiente al mayor de los períodos
StrSql = "SELECT pliqnro, pliqmes, pliqanio FROM periodo WHERE pliqanio = '" & anio & "' AND pliqnro = '" & perhasta & "'"
OpenRecordset StrSql, rs_Periodo
MesPeriodo2 = 0
If Not rs_Periodo.EOF Then MesPeriodo2 = rs_Periodo!pliqmes
rs_Periodo.Close
' Traigo todos los periodos de liquidación válidos para la selección del usuario
StrSql = "SELECT pliqnro, pliqmes, pliqanio, pliqdesde, pliqhasta, pliqdesc FROM periodo WHERE pliqmes >= '" & _
MesPeriodo & "' AND pliqmes <= '" & MesPeriodo2 & "' AND pliqanio = '" & anio & "' ORDER BY pliqmes DESC"
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If

rs_Periodo.MoveFirst
Do
    RegDetalle = RegDetalle + 1
    ReDim Preserve Memopliq(RegDetalle)
    ReDim Preserve Memopliqmes(RegDetalle)
    ReDim Preserve Memopliqanio(RegDetalle)
    ReDim Preserve Memopliqdesde(RegDetalle)
    ReDim Preserve Memopliqhasta(RegDetalle)
    ReDim Preserve Memopliqdesc(RegDetalle)
    Memopliq(RegDetalle) = rs_Periodo!pliqnro
    Memopliqmes(RegDetalle) = rs_Periodo!pliqmes
    Memopliqanio(RegDetalle) = rs_Periodo!pliqanio
    Memopliqdesde(RegDetalle) = rs_Periodo!pliqdesde
    Memopliqhasta(RegDetalle) = rs_Periodo!pliqhasta
    Memopliqdesc(RegDetalle) = rs_Periodo!pliqdesc
    rs_Periodo.MoveNext
Loop Until rs_Periodo.EOF
CantPeriodos = RegDetalle
rs_Periodo.Close




j = 0
'Armo el sector de detalle de un reporte
StrSql = "SELECT confnrocol, confval2, confetiq, conftipo FROM confrep where repnro = 254 order by confnrocol"
OpenRecordset StrSql, rs_Confrep

k = 1
If Not rs_Confrep.EOF Then
    rs_Confrep.MoveFirst
    Do
        If (rs_Confrep!conftipo) <> "FAC" And UCase(rs_Confrep!conftipo) <> "FCO" Then
            ReDim Preserve ColumnasConfrep(k)
            ColumnasConfrep(k) = rs_Confrep!Confetiq
        End If
        k = k + 1
        rs_Confrep.MoveNext
    Loop Until rs_Confrep.EOF
    rs_Confrep.MoveFirst
Else
    Flog.writeline "No se encontraron acumuladores configurados para el sector de detalles del reporte. Abortando"
    Exit Sub
End If

'Me guardo la consulta de empleados en StrEmpleado
If tenro3 <> "" And tenro3 <> "0" Then ' esto ocurre solo cuando se seleccionan los tres niveles
    StrSql = " SELECT DISTINCT empleado.ternro "
    StrSql = StrSql & " FROM empleado  "
    'StrSql = StrSql & " INNER JOIN fases ON altfec <=" & ConvFecha(Memopliqhasta(I)) & " AND (bajfec is NULL OR bajfec >" & ConvFecha(Memopliqdesde(I)) & ") AND empleado.ternro=fases.empleado"
    StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
    StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(Now) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(Now) & "))"
    If estrnro1 <> "" And estrnro1 <> "0" And estrnro1 <> "-1" Then 'cuando se le asigna un valor al nivel 1
        StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
    End If
    StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2
    StrSql = StrSql & " AND (estact2.htetdesde<=" & ConvFecha(Now) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(Now) & "))"
    If estrnro2 <> "" And estrnro2 <> "0" And estrnro2 <> "-1" Then 'cuando se le asigna un valor al nivel 2
        StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
    End If
    StrSql = StrSql & " INNER JOIN his_estructura estact3 ON empleado.ternro = estact3.ternro  AND estact3.tenro  = " & tenro3 & _
    " AND (estact3.htetdesde<=" & ConvFecha(Now) & " AND (estact3.htethasta is null or estact3.htethasta>=" & ConvFecha(Now) & "))"
    If estrnro3 <> "" And estrnro3 <> "0" And estrnro3 <> "-1" Then 'cuando se le asigna un valor al nivel 3
        StrSql = StrSql & " AND estact3.estrnro =" & estrnro3
    End If
    If estado = 1 Then
        StrSql = StrSql & " WHERE " & "(empleg >= " & legdesde & ") AND (empleg <= " & leghasta & ")"
    Else
        StrSql = StrSql & " WHERE " & "(empleg >= " & legdesde & ") AND (empleg <= " & leghasta & ") AND (empest = " & estado & ")"
    End If
'    StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,tenro3,estrnro3,"&orden
Else
    If tenro2 <> "" And tenro2 <> "0" Then ' ocurre cuando se selecciono hasta el segundo nivel
        StrSql = "SELECT DISTINCT empleado.ternro"
        StrSql = StrSql & " FROM empleado  "
        'StrSql = StrSql & " INNER JOIN fases ON altfec <=" & ConvFecha(Memopliqhasta(I)) & " AND (bajfec is NULL OR bajfec >" & ConvFecha(Memopliqdesde(I)) & ") AND empleado.ternro=fases.empleado"
        StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
        StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(Now) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(Now) & "))"
        If estrnro1 <> "" And estrnro1 <> "0" And estrnro1 <> "-1" Then
            StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
        End If
        StrSql = StrSql & " INNER JOIN his_estructura estact2 ON empleado.ternro = estact2.ternro  AND estact2.tenro  = " & tenro2 & _
        " AND (estact2.htetdesde<=" & ConvFecha(Now) & " AND (estact2.htethasta is null or estact2.htethasta>=" & ConvFecha(Now) & "))"
        If estrnro2 <> "" And estrnro2 <> "0" And estrnro2 <> "-1" Then
            StrSql = StrSql & " AND estact2.estrnro =" & estrnro2
        End If
        If estado = 1 Then
            StrSql = StrSql & " WHERE " & "(empleg >= " & legdesde & ") AND (empleg <= " & leghasta & ")"
        Else
            StrSql = StrSql & " WHERE " & "(empleg >= " & legdesde & ") AND (empleg <= " & leghasta & ") AND (empest = " & estado & ")"
        End If
'       StrSql = StrSql & " ORDER BY tenro1,estrnro1,tenro2,estrnro2,"&orden
    Else
        If tenro1 <> "" And tenro1 <> "0" Then ' Cuando solo selecionamos el primer nivel
            StrSql = "SELECT DISTINCT empleado.ternro "
            StrSql = StrSql & " FROM empleado  "
            'StrSql = StrSql & " INNER JOIN fases ON altfec <=" & ConvFecha(Memopliqhasta(I)) & " AND (bajfec is NULL OR bajfec >" & ConvFecha(Memopliqdesde(I)) & ") AND empleado.ternro=fases.empleado"
            StrSql = StrSql & " INNER JOIN his_estructura estact1 ON empleado.ternro = estact1.ternro  AND estact1.tenro  = " & tenro1
            StrSql = StrSql & " AND (estact1.htetdesde<=" & ConvFecha(Now) & " AND (estact1.htethasta is null or estact1.htethasta>=" & ConvFecha(Now) & "))"
            If estrnro1 <> "" And estrnro1 <> "0" And estrnro1 <> "-1" Then
                StrSql = StrSql & " AND estact1.estrnro =" & estrnro1
            End If
            StrSql = StrSql & " WHERE " & "(empleg >= " & legdesde & ") AND (empleg <= " & leghasta & ")"
            'StrSql = StrSql & " ORDER BY tenro1,estrnro1,"&orden

        Else ' cuando no hay nivel de estructura seleccionado
            StrSql = " SELECT DISTINCT empleado.ternro "
            StrSql = StrSql & " FROM empleado  "
            'StrSql = StrSql & " INNER JOIN fases ON altfec <=" & ConvFecha(Memopliqhasta(I)) & " AND (bajfec is NULL OR bajfec >" & ConvFecha(Memopliqdesde(I)) & ") AND empleado.ternro=fases.empleado"
            If estado = 1 Then
                StrSql = StrSql & " WHERE " & "(empleg >= " & legdesde & ") AND (empleg <= " & leghasta & ")"
            Else
                StrSql = StrSql & " WHERE " & "(empleg >= " & legdesde & ") AND (empleg <= " & leghasta & ") AND (empest = " & estado & ")"
            End If
'           StrSql = StrSql & " ORDER BY "&orden
        End If
    End If
End If
'Una vez q tengo los empleados filtro en el intervalo de período seleccionado por el usuario
StrEmpleado = StrSql




'Primero genero la memoria con los resultados sin detallar, es decir para el global de empleados elegidos
For I = 1 To CantPeriodos
    
    ReDim Preserve NombrePeriodos(I)
    NombrePeriodos(I) = Memopliqdesc(I)
    
    RegDetalle = I
    ColDetalle = 0
    If Not rs_Confrep.EOF Then
        rs_Confrep.MoveFirst
        Do
            ColDetalle = ColDetalle + 1
            'Pregunto si se trata de un acumulador o un concepto
            Select Case UCase(rs_Confrep!conftipo)
                Case "DAC" 'Es un acumulador
                    'Hago la consulta con la lista de empleados desde ASP o bien con la consulta SQL
                    If ListadoASP <> "" And ListadoASP <> "0" Then 'viene desde ASP
                        StrSql = "SELECT SUM(almonto) SUMA FROM acu_liq"
                        StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = acu_liq.cliqnro"
                        StrSql = StrSql & " AND cabliq.empleado IN(" & ListadoASP & "))"
                        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                        StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                        StrSql = StrSql & " And acu_liq.acuNro = " & rs_Confrep!Confval2
                        OpenRecordset StrSql, rs_EmplFiltro
                        ReDim Preserve Detalle(20, RegDetalle)
                        Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                    Else 'Hago la consulta de empleados
                        StrSql = "SELECT SUM(almonto) SUMA FROM acu_liq"
                        StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = acu_liq.cliqnro"
                        StrSql = StrSql & " AND cabliq.empleado IN(" & StrEmpleado & "))"
                        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                        StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                        StrSql = StrSql & " And acu_liq.acuNro = " & rs_Confrep!Confval2
                        OpenRecordset StrSql, rs_EmplFiltro
                        ReDim Preserve Detalle(20, RegDetalle)
                        Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                    End If
                Case "FAC" 'Es el acumulador final
                    'If Not AcumuladorFinal Then
                        AcumuladorFinal = True
                        'Detalle(RegDetalle, 1) = "F"
                        'RegDetalle = RegDetalle + 1
                        'Hago la consulta con la lista de empleados desde ASP o bien con la consulta SQL
                        If ListadoASP <> "" And ListadoASP <> "0" Then 'viene desde ASP
                            StrSql = "SELECT SUM(almonto) SUMA FROM acu_liq"
                            StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = acu_liq.cliqnro"
                            StrSql = StrSql & " AND cabliq.empleado IN(" & ListadoASP & "))"
                            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                            StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                            StrSql = StrSql & " And acu_liq.acuNro = " & rs_Confrep!Confval2
                            OpenRecordset StrSql, rs_EmplFiltro
                            'ReDim Preserve Detalle(20, RegDetalle)
                            'Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                            NombreF = rs_Confrep!Confetiq
                            If Not IsNull(rs_EmplFiltro(0)) Then
                                ValorF = ValorF + CLng(rs_EmplFiltro(0))
                            End If
                        Else 'Hago la consulta de empleados
                            StrSql = "SELECT SUM(almonto) SUMA FROM acu_liq"
                            StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = acu_liq.cliqnro"
                            StrSql = StrSql & " AND cabliq.empleado IN(" & StrEmpleado & "))"
                            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                            StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                            StrSql = StrSql & " And acu_liq.acuNro = " & rs_Confrep!Confval2
                            OpenRecordset StrSql, rs_EmplFiltro
                            'ReDim Preserve Detalle(20, RegDetalle)
                            'Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                            NombreF = rs_Confrep!Confetiq
                            If Not IsNull(rs_EmplFiltro(0)) Then
                                ValorF = ValorF + CLng(rs_EmplFiltro(0))
                            End If
                        End If
                    'End If
                Case "DCO" 'Es un concepto
                    'Hago la consulta con la lista de empleados desde ASP o bien con la consulta SQL
                    If ListadoASP <> "" And ListadoASP <> "0" Then 'viene desde ASP
                        'Primero me fijo de que concepto se trata
                        StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!Confval2
                        OpenRecordset StrSql, rs_Conceptos
                        If Not rs_Conceptos.EOF Then Concepto = rs_Conceptos!concnro
                        rs_Conceptos.Close
                        StrSql = "SELECT SUM(dlimonto) SUMA FROM detliq"
                        StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = detliq.cliqnro"
                        StrSql = StrSql & " AND cabliq.empleado IN(" & ListadoASP & "))"
                        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                        StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                        StrSql = StrSql & " And detliq.concnro = " & Concepto
                        OpenRecordset StrSql, rs_EmplFiltro
                        ReDim Preserve Detalle(20, RegDetalle)
                        Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                    Else
                        'Primero me fijo de que concepto se trata
                        StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!Confval2
                        OpenRecordset StrSql, rs_Conceptos
                        If Not rs_Conceptos.EOF Then Concepto = rs_Conceptos!concnro
                        rs_Conceptos.Close
                        StrSql = "SELECT SUM(dlimonto) SUMA FROM detliq"
                        StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = detliq.cliqnro"
                        StrSql = StrSql & " AND cabliq.empleado IN(" & StrEmpleado & "))"
                        StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                        StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                        StrSql = StrSql & " And detliq.concnro = " & Concepto
                        OpenRecordset StrSql, rs_EmplFiltro
                        ReDim Preserve Detalle(20, RegDetalle)
                        Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                    End If
                Case "FCO" 'Es el concepto final
                    'Hago la consulta con la lista de empleados desde ASP o bien con la consulta SQL
                    'If Not ConceptoFinal Then
                        ConceptoFinal = True
                        'Detalle(RegDetalle, 1) = "F"
                        'RegDetalle = RegDetalle + 1
                        If ListadoASP <> "" And ListadoASP <> "0" Then 'viene desde ASP
                            'Primero me fijo de que concepto se trata
                            StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!Confval2
                            OpenRecordset StrSql, rs_Conceptos
                            If Not rs_Conceptos.EOF Then Concepto = rs_Conceptos!concnro
                            rs_Conceptos.Close
                            StrSql = "SELECT SUM(dlimonto) SUMA FROM detliq"
                            StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = detliq.cliqnro"
                            StrSql = StrSql & " AND cabliq.empleado IN(" & ListadoASP & "))"
                            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                            StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                            StrSql = StrSql & " And detliq.concnro = " & Concepto
                            OpenRecordset StrSql, rs_EmplFiltro
                            NombreF = rs_Confrep!Confetiq
                            If Not IsNull(rs_EmplFiltro(0)) Then
                                ValorF = ValorF + CLng(rs_EmplFiltro(0))
                            End If
                            'ReDim Preserve Detalle(20, RegDetalle)
                            'Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                        Else
                            'Primero me fijo de que concepto se trata
                            StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!Confval2
                            OpenRecordset StrSql, rs_Conceptos
                            If Not rs_Conceptos.EOF Then Concepto = rs_Conceptos!concnro
                            rs_Conceptos.Close
                            StrSql = "SELECT SUM(dlimonto) SUMA FROM detliq"
                            StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = detliq.cliqnro"
                            StrSql = StrSql & " AND cabliq.empleado IN(" & StrEmpleado & "))"
                            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                            StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                            StrSql = StrSql & " And detliq.concnro = " & Concepto
                            OpenRecordset StrSql, rs_EmplFiltro
                            NombreF = rs_Confrep!Confetiq
                            If Not IsNull(rs_EmplFiltro(0)) Then
                                ValorF = ValorF + CLng(rs_EmplFiltro(0))
                            End If
                            'ReDim Preserve Detalle(20, RegDetalle)
                            'Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                        End If
                    'End If
            End Select
            rs_EmplFiltro.Close
            rs_Confrep.MoveNext
        Loop Until rs_Confrep.EOF
        If I < CantPeriodos Then rs_Confrep.MoveFirst
    End If
Next
rs_Confrep.Close

'Armo la memoria EmpleadosArray, según venga de ASP o bien lo calculo desde acá
'Levanto la lista de empleados que viene desde ASP, si es que viene
If ListadoASP <> "" And ListadoASP <> "0" Then
    Dim EmpleadosArray() As String
    EmpleadosArray = Split(ListadoASP, ",")
Else
    OpenRecordset StrEmpleado, rs_EmplFiltro
    If Not rs_EmplFiltro.EOF Then
        rs_EmplFiltro.MoveFirst
        j = 1
        Do
           ReDim Preserve EmpleadosArray(j) As String
           EmpleadosArray(j) = rs_EmplFiltro!ternro
           rs_EmplFiltro.MoveNext
           j = j + 1
        Loop Until rs_EmplFiltro.EOF
    Else
        NoEmp = True
    End If
End If

StrSql = "SELECT confnrocol, confval2, confetiq, conftipo FROM confrep where repnro = 254 order by confnrocol"
OpenRecordset StrSql, rs_Confrep

If Not rs_Confrep.EOF Then
    rs_Confrep.MoveFirst
End If

IncPorc = 60 / (UBound(EmpleadosArray) * CantPeriodos)
Progreso = 0
If Not NoEmp Then
    For j = 1 To UBound(EmpleadosArray)
        'Ahora genero la memoria con los resultados detallados, es decir para cada uno de los empleados elegidos
        'Guardo primero el ternro
        rs_Confrep.MoveFirst
        RegDetalle = RegDetalle + 1
        ReDim Preserve Detalle(20, RegDetalle)
        Detalle(1, RegDetalle) = "Empleado:"
        Detalle(2, RegDetalle) = EmpleadosArray(j)
        AcumuladorFinal = False
        ConceptoFinal = False
        
        For I = 1 To CantPeriodos
            RegDetalle = RegDetalle + 1
            ColDetalle = 0
            If Not rs_Confrep.EOF Then
                rs_Confrep.MoveFirst
                Do
                    ColDetalle = ColDetalle + 1
                    'Pregunto si se trata de un acumulador o un concepto
                    Select Case UCase(rs_Confrep!conftipo)
                        Case "DAC" 'Es un acumulador
                            StrSql = "SELECT SUM(almonto) SUMA FROM acu_liq"
                            StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = acu_liq.cliqnro"
                            StrSql = StrSql & " AND cabliq.empleado = " & EmpleadosArray(j) & ")"
                            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                            StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                            StrSql = StrSql & " And acu_liq.acuNro = " & rs_Confrep!Confval2
                            OpenRecordset StrSql, rs_EmplFiltro
                            ReDim Preserve Detalle(20, RegDetalle)
                            Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                        Case "FAC" 'Es el acumulador final
                            'If Not AcumuladorFinal Then
                                AcumuladorFinal = True
                                'ReDim Preserve Detalle(20, RegDetalle)
                                'Detalle(1, RegDetalle) = "F"
                                'RegDetalle = RegDetalle + 1
                                StrSql = "SELECT SUM(almonto) SUMA FROM acu_liq"
                                StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = acu_liq.cliqnro"
                                StrSql = StrSql & " AND cabliq.empleado = " & EmpleadosArray(j) & ")"
                                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                                StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                                StrSql = StrSql & " And acu_liq.acuNro = " & rs_Confrep!Confval2
                                OpenRecordset StrSql, rs_EmplFiltro
                                ReDim Preserve ValoresF(j)
                                If Not IsNull(rs_EmplFiltro(0)) Then
                                    ValoresF(j) = ValoresF(j) + CLng(rs_EmplFiltro(0))
                                End If
                                NombreF = rs_Confrep!Confetiq
                                'Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                            'End If
                        Case "DCO" 'Es un concepto
                            'Primero me fijo de que concepto se trata
                            StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!Confval2
                            OpenRecordset StrSql, rs_Conceptos
                            If Not rs_Conceptos.EOF Then Concepto = rs_Conceptos!concnro
                            rs_Conceptos.Close
                            StrSql = "SELECT SUM(dlimonto) SUMA FROM detliq"
                            StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = detliq.cliqnro"
                            StrSql = StrSql & " AND cabliq.empleado = " & EmpleadosArray(j) & ")"
                            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                            StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                            StrSql = StrSql & " And detliq.concnro = " & Concepto
                            OpenRecordset StrSql, rs_EmplFiltro
                            ReDim Preserve Detalle(20, RegDetalle)
                            Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                        Case "FCO" 'Es el concepto final
                            'Hago la consulta con la lista de empleados desde ASP o bien con la consulta SQL
                            'If Not ConceptoFinal Then
                                ConceptoFinal = True
                                'ReDim Preserve Detalle(20, RegDetalle)
                                'Detalle(1, RegDetalle) = "F"
                                'RegDetalle = RegDetalle + 1
                                'Primero me fijo de que concepto se trata
                                StrSql = "SELECT concnro FROM concepto WHERE conccod = " & rs_Confrep!Confval2
                                OpenRecordset StrSql, rs_Conceptos
                                If Not rs_Conceptos.EOF Then Concepto = rs_Conceptos!concnro
                                rs_Conceptos.Close
                                StrSql = "SELECT SUM(dlimonto) SUMA FROM detliq"
                                StrSql = StrSql & " INNER JOIN cabliq ON (cabliq.cliqnro = detliq.cliqnro"
                                StrSql = StrSql & " AND cabliq.empleado = " & EmpleadosArray(j) & ")"
                                StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
                                StrSql = StrSql & " Where Proceso.pliqnro = " & Memopliq(I)
                                StrSql = StrSql & " And detliq.concnro = " & Concepto
                                OpenRecordset StrSql, rs_EmplFiltro
                                ReDim Preserve ValoresF(j)
                                NombreF = rs_Confrep!Confetiq
                                If Not IsNull(rs_EmplFiltro(0)) Then
                                    ValoresF(j) = ValoresF(j) + CLng(rs_EmplFiltro(0))
                                End If
                                'ReDim Preserve Detalle(20, RegDetalle)
                                'Detalle(ColDetalle, RegDetalle) = rs_EmplFiltro(0)
                            'End If
                    End Select
                    rs_Confrep.MoveNext
                Loop Until rs_Confrep.EOF
                If I < CantPeriodos Then rs_Confrep.MoveFirst
            End If
            'Progreso = Round(Progreso + IncPorc, 2)
            Progreso = Progreso + IncPorc
            UpdateProgreso (Progreso)
        Next
    Next
Else
    Flog.writeline "No se encontraron empleados seleccionados para el reporte."
    Exit Sub
End If

Flog.writeline "Calculando Promedios, mínimos y máximos del sector de detalles"

If ListadoASP <> "" And ListadoASP <> "0" Then 'viene desde ASP
    Referencia = "Reporte generado para " & UBound(EmpleadosArray) & " empleados seleccionados manualmente"
Else
    Referencia = "Reporte generado para " & UBound(EmpleadosArray) & " empleados obtenidos por filtro."
    If estrnro1 <> "" Then
        Referencia = Referencia & " Estructuras involucradas: "
        StrSql = "select tedabr, estrdabr from estructura "
        StrSql = StrSql & "INNER JOIN tipoestructura ON tipoestructura.tenro=estructura.tenro "
        StrSql = StrSql & "WHERE estrnro = " & estrnro1
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            rs_Estructura.MoveFirst
            Referencia = Referencia & rs_Estructura!tedabr & " - " & rs_Estructura!estrdabr
        End If
    End If
    If estrnro2 <> "" Then
        StrSql = "select tedabr, estrdabr from estructura "
        StrSql = StrSql & "INNER JOIN tipoestructura ON tipoestructura.tenro=estructura.tenro "
        StrSql = StrSql & "WHERE estrnro = " & estrnro2
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            rs_Estructura.MoveFirst
            Referencia = Referencia & " - " & rs_Estructura!tedabr & " - " & rs_Estructura!estrdabr
        End If
    End If
    If estrnro3 <> "" Then
        StrSql = "select tedabr, estrdabr from estructura "
        StrSql = StrSql & "INNER JOIN tipoestructura ON tipoestructura.tenro=estructura.tenro "
        StrSql = StrSql & "WHERE estrnro = " & estrnro3
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            rs_Estructura.MoveFirst
            Referencia = Referencia & " - " & rs_Estructura!tedabr & " - " & rs_Estructura!estrdabr
        End If
    End If
End If



'Calculo los totales de cada periodo
For I = 1 To RegDetalle
    If IsNull(Detalle(20, I)) Then Detalle(20, I) = 0
    For j = 1 To 19
        If IsNull(Detalle(j, I)) Then Detalle(j, I) = 0
        Detalle(20, I) = Val(Detalle(20, I)) + Val(Detalle(j, I))
    Next
Next

'Const Constante = RegDetalle
'Armo la meoria con formato adecuado para insertar en la base de datos
ReDim detalle2(RegDetalle, 20)
fila = 0
For I = 1 To RegDetalle
    fila = fila + 1
    columna = 0
    For j = 1 To 20
        columna = columna + 1
        detalle2(fila, columna) = Detalle(j, I)
    Next
Next
'I = 1
'rs_Confrep.MoveFirst
'Do
'    ReDim Preserve ColumnasConfrep(I)
'    ColumnasConfrep(I) = StrEncontrado
'Loop Until rs_Confrep.EOF
'
'
'
'ReDim Preserve ColumnasConfrep(RegDetalle)
'ColumnasConfrep(RegDetalle) = StrEncontrado
'ReDim Preserve NombrePeriodos(I)
''NombrePeriodos(I) = Memopliqmes(I) & "/" & Memopliqanio(I)
'NombrePeriodos(I) = Memopliqdesc(I)


'Una vez cargadas todas las memorias hago los INSERTS correspondientes en la tabla



Flog.writeline "Insertando en la base:"


RegDetalle = 0: j = 1
'Primero inserto la memoria de nombres de periodos
If UBound(Detalle) > 0 Then
    StrSql = "INSERT INTO rep_nivgeren (bpronro,fecha,hora,iduser,filanro,valor,valor1"
    StrSql = StrSql & ",valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10"
    StrSql = StrSql & ",valor11,valor12"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & NroProcesoBatch & ","
    StrSql = StrSql & ConvFecha(Date) & ","
    StrSql = StrSql & "'" & Mid(Time, 1, 8) & "',"
    StrSql = StrSql & "'" & iduser & "',"
    StrSql = StrSql & j
    For ColDetalle = 1 To CantPeriodos
        StrSql = StrSql & ",'" & NombrePeriodos(ColDetalle) & "'"
    Next
    If CantPeriodos < 13 Then
        For ColDetalle = CantPeriodos + 1 To 13
            StrSql = StrSql & ",NULL"
        Next
    End If
    'StrSql = StrSql & "'Mínimo',"
    'StrSql = StrSql & "'Máximo'"
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
End If
j = j + 1
'Luego inserto los nombres de las columnas
If UBound(Detalle) > 0 Then
    StrSql = "INSERT INTO rep_nivgeren (bpronro,fecha,hora,iduser,filanro,valor,valor1"
    StrSql = StrSql & ",valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10"
    StrSql = StrSql & ",valor11,valor12"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & NroProcesoBatch & ","
    StrSql = StrSql & ConvFecha(Date) & ","
    StrSql = StrSql & "'" & Mid(Time, 1, 8) & "',"
    StrSql = StrSql & "'" & iduser & "',"
    StrSql = StrSql & j
    For ColDetalle = 1 To UBound(ColumnasConfrep)
        StrSql = StrSql & ",'" & ColumnasConfrep(ColDetalle) & "'"
    Next
    For ColDetalle = UBound(ColumnasConfrep) + 1 To 13
        StrSql = StrSql & ",NULL"
    Next
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
End If
j = j + 1
'Luego inserto la referencia
'If UBound(Detalle) > 0 Then
    StrSql = "INSERT INTO rep_nivgeren (bpronro,fecha,hora,iduser,filanro,valor,valor1"
    StrSql = StrSql & ",valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10"
    StrSql = StrSql & ",valor11,valor12"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & NroProcesoBatch & ","
    StrSql = StrSql & ConvFecha(Date) & ","
    StrSql = StrSql & "'" & Mid(Time, 1, 8) & "',"
    StrSql = StrSql & "'" & iduser & "',"
    StrSql = StrSql & j
    I = 1
    For k = 1 To Len(Referencia) / 40 + 1
        StrSql = StrSql & ",'" & Mid(Referencia, I, 40) & "'"
        I = k * 40 + 1
    Next
    For I = 1 To 14 - k
        StrSql = StrSql & ",NULL"
    Next
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
'End Ifç
j = j + 1
'En caso de haber sido configurado un acumulador final o un concepto final inserto un SI, caso contrario 0
'Si coloco un SI, luego coloco el número de registros (empleados+general)
If AcumuladorFinal Or ConceptoFinal Then
    StrSql = "INSERT INTO rep_nivgeren (bpronro,fecha,hora,iduser,filanro,valor,valor1"
    StrSql = StrSql & ",valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10"
    StrSql = StrSql & ",valor11,valor12"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & NroProcesoBatch & ","
    StrSql = StrSql & ConvFecha(Date) & ","
    StrSql = StrSql & "'" & Mid(Time, 1, 8) & "',"
    StrSql = StrSql & "'" & iduser & "',"
    StrSql = StrSql & j & ",'"
    StrSql = StrSql & "SI" & "',"
    StrSql = StrSql & UBound(ValoresF) + 1
    For k = 1 To 11
        StrSql = StrSql & ",NULL"
    Next
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    j = j + 1
    'Luego inserto los valores del concepto o acumulador final
    StrSql = "INSERT INTO rep_nivgeren (bpronro,fecha,hora,iduser,filanro,valor,valor1"
    StrSql = StrSql & ",valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10"
    StrSql = StrSql & ",valor11,valor12"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & NroProcesoBatch & ","
    StrSql = StrSql & ConvFecha(Date) & ","
    StrSql = StrSql & "'" & Mid(Time, 1, 8) & "',"
    StrSql = StrSql & "'" & iduser & "',"
    StrSql = StrSql & j & ",'"
    StrSql = StrSql & NombreF & "',"
    StrSql = StrSql & ValorF
    For k = 1 To 11
        StrSql = StrSql & ",NULL"
    Next
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    For I = 1 To UBound(ValoresF)
        j = j + 1
        StrSql = "INSERT INTO rep_nivgeren (bpronro,fecha,hora,iduser,filanro,valor,valor1"
        StrSql = StrSql & ",valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10"
        StrSql = StrSql & ",valor11,valor12"
        StrSql = StrSql & ") VALUES ("
        StrSql = StrSql & NroProcesoBatch & ","
        StrSql = StrSql & ConvFecha(Date) & ","
        StrSql = StrSql & "'" & Mid(Time, 1, 8) & "',"
        StrSql = StrSql & "'" & iduser & "',"
        StrSql = StrSql & j & ",'"
        StrSql = StrSql & NombreF & "',"
        StrSql = StrSql & ValoresF(I)
        For k = 1 To 11
            StrSql = StrSql & ",NULL"
        Next
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Next
Else
    StrSql = "INSERT INTO rep_nivgeren (bpronro,fecha,hora,iduser,filanro,valor,valor1"
    StrSql = StrSql & ",valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10"
    StrSql = StrSql & ",valor11,valor12"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & NroProcesoBatch & ","
    StrSql = StrSql & ConvFecha(Date) & ","
    StrSql = StrSql & "'" & Mid(Time, 1, 8) & "',"
    StrSql = StrSql & "'" & iduser & "',"
    StrSql = StrSql & j & ",'"
    StrSql = StrSql & "NO" & "'"
    For k = 1 To 12
        StrSql = StrSql & ",NULL"
    Next
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
End If

IncPorc = (99 - Progreso) / (UBound(detalle2))

'Comienzo a insertar las memorias
For I = 1 To UBound(detalle2)
    j = j + 1
    StrSql = "INSERT INTO rep_nivgeren (bpronro,fecha,hora,iduser,filanro,valor,valor1"
    StrSql = StrSql & ",valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10"
    StrSql = StrSql & ",valor11,valor12"
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & NroProcesoBatch & ","
    StrSql = StrSql & ConvFecha(Date) & ","
    StrSql = StrSql & "'" & Mid(Time, 1, 8) & "',"
    StrSql = StrSql & "'" & iduser & "',"
    StrSql = StrSql & j
    For ColDetalle = 1 To 12
        If IsNull(detalle2(I, ColDetalle)) Or detalle2(I, ColDetalle) = "" Then
            StrSql = StrSql & ",NULL"
        Else
            StrSql = StrSql & ",'" & detalle2(I, ColDetalle) & "'"
        End If
    Next
    'inserto los totales
    StrSql = StrSql & ",'" & detalle2(I, 20) & "'"
    StrSql = StrSql & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    Progreso = Round(Progreso + IncPorc, 2)
    UpdateProgreso (Progreso)
Next






'Fin de la transaccion
10 If Not HuboError Then
    MyCommitTrans
Else
    MyRollbackTrans
End If


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Acu_liq.State = adStateOpen Then rs_Acu_liq.Close
If rs_Proceso.State = adStateOpen Then rs_Proceso.Close
If rs_EmplFiltro.State = adStateOpen Then rs_EmplFiltro.Close
If rs_Detalle.State = adStateOpen Then rs_Detalle.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_ConfrepInd.State = adStateOpen Then rs_ConfrepInd.Close
If Aux_rs_Confrep.State = adStateOpen Then Aux_rs_Confrep.Close
If rs_Conceptos.State = adStateOpen Then rs_Conceptos.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Impproarg.State = adStateOpen Then rs_Impproarg.Close
If rs_Repsijp.State = adStateOpen Then rs_Repsijp.Close
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
If rs_Familia.State = adStateOpen Then rs_Familia.Close
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
If rs_TipoCont.State = adStateOpen Then rs_TipoCont.Close
If rs_sucursal.State = adStateOpen Then rs_sucursal.Close
If rs_zona.State = adStateOpen Then rs_zona.Close
If rs_Estr_cod.State = adStateOpen Then rs_Estr_cod.Close
If rs_Convenios.State = adStateOpen Then rs_Convenios.Close
If rs_Familiar.State = adStateOpen Then rs_Familiar.Close
If rs_Estudio_Actual.State = adStateOpen Then rs_Estudio_Actual.Close
If rs_Nivest.State = adStateOpen Then rs_Nivest.Close
If rs_Consulta.State = adStateOpen Then rs_Consulta.Close

Set rs_Empleados = Nothing
Set rs_ConfrepInd = Nothing
Set rs_Acu_liq = Nothing
Set rs_Proceso = Nothing
Set rs_EmplFiltro = Nothing
Set rs_Detalle = Nothing
Set rs_Periodo = Nothing
Set rs_Confrep = Nothing
Set Aux_rs_Confrep = Nothing
Set rs_Conceptos = Nothing
Set rs_Detliq = Nothing
Set rs_Impproarg = Nothing
Set rs_Repsijp = Nothing
Set rs_Tercero = Nothing
Set rs_Periodo = Nothing
Set rs_Estructura = Nothing
Set rs_Cuil = Nothing
Set rs_Familia = Nothing
Set rs_Empresa = Nothing
Set rs_TipoCont = Nothing
Set rs_sucursal = Nothing
Set rs_zona = Nothing
Set rs_Estr_cod = Nothing
Set rs_Convenios = Nothing
Set rs_Familiar = Nothing
Set rs_Estudio_Actual = Nothing
Set rs_Nivest = Nothing
Set rs_Consulta = Nothing

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyRollbackTrans
    MyBeginTrans
    Progreso = Round(Progreso + IncPorc, 2)
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

End Sub

Public Function UpdateProgreso(ByVal Progreso)
    MyBeginTrans
    Progreso = Round(Progreso + IncPorc, 2)
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Function

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


Public Function Calcular_Edad(ByVal Fecha As Date) As Integer
'...........................................................................
' Archivo       : edad.i                              fecha ini. : 20/01/92
' Autor         : FGZ
' Fecha         :
' Ultima Mod    : 25/07/2005 - Se calcula a la fecha fin del periodo.
'...........................................................................
Dim años  As Integer
Dim ALaFecha As Date

    ALaFecha = C_Date(Date)

    años = Year(ALaFecha) - Year(Fecha)
    If Month(ALaFecha) < Month(Fecha) Then
       años = años - 1
    Else
        If Month(ALaFecha) = Month(Fecha) Then
            If Day(ALaFecha) < Day(Fecha) Then
                años = años - 1
            End If
        End If
    End If
    Calcular_Edad = años
End Function


