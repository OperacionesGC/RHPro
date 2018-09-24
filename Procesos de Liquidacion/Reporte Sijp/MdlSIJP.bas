Attribute VB_Name = "MdlSIJP"
Option Explicit

'Version: 1.01
'   Nueva restriccion sobre columna Sueldo + Adic.

'Const Version = 1.02
'Const FechaVersion = "13/09/2005"

'Const Version = 1.03 ' Se cambio la descripcion al buscar la caja de jubilacion
'Const FechaVersion = "15/12/2005"

'Const Version = 1.04 ' Se agregó Aux_sac en las restricciones
'Const FechaVersion = "11/01/2006"

'Const Version = 1.05
' Se cambió la rutina de comparación de intervalos de fechas de fases y
' de estructuras, para los empleados con más de una empresa asociada

'Const Version = 1.06222
'Const FechaVersion = "03/03/2006"
'Autor Modificación = Hernán D. Santonocito
'Contempla los movimiento entre empresas dentro de la base de datos basandoce en las faces y la estructura de empresa

'Const Version = 1.07
'Const FechaVersion = "02/06/2006"
'Autor Modificación = Martin Ferraro
'Lista_Pro_F se truncaba en 100 y despues al borrar procesos no se encontraban todos
'Se encontro y corrigio error en sql fases: hacia movenext cuando era eof

'Const Version = 1.08
'Const FechaVersion = "10/07/2006"
'Autor Modificación = Martin Ferraro
'Se agregaron 4 columnas via confrep (42,43,44,45) segun SIJP Version 26

'Const Version = 1.09
'Const FechaVersion = "21/07/2006"
'Autor Modificación = Martin Ferraro
'Se cambio el orden en que se realiza el tope contra el imponible

'Const Version = 1.1
'Const FechaVersion = "13/10/2006"
'Autor Modificación = Breglia Maximiliano
'Se agregó un campo de cant de hs extra version 27

'Const Version = 1.2
'Const FechaVersion = "18/10/2006"
''Autor Modificación = Raul CHinestra
''Se agregó una condicion para la impresion de la cant de hs extra version 27

'Const Version = 1.13
'Const FechaVersion = "27/10/2006"
'Autor Modificación = FGZ
'Se agregó una un tipo de columna en el confrep porque hay varios conceptos de horas extras y debe acumular la cantidad y el tipo
' PCO no es cumulativo ==> se creo CCO que es cantidad del concepto y acumulativo.
' Aplica para la version 27
'OJO con la numeracion!!! la version anterior era 1.12 y no 1.2 esto puede confundir

'Const Version = 1.14
'Const FechaVersion = "10/01/2008"
'Autor Modificación = MBreglia
'agregue esta linea And hrsextras > 0 pq superaba el tope pero no tenia monto de hsestrs y si cantidad
                
'Const Version = 1.15
'Const FechaVersion = "14/01/2008"
'Autor Modificación = Martin Ferraro
'columna cod.Siniest. es 1 si existe licencia de tipo configurable por confrep con algun día
'dentro del rango de fechas del periodo del filtro.
                
'Const Version = 1.16
'Const FechaVersion = "15/02/2008"
'Autor Modificación = Martin Ferraro
'Rem imponible 5 puede ser tomada por restriccion
'
                
'Const Version = 1.17
'Const FechaVersion = "21/02/2008"
'Autor Modificación = Martin Ferraro
'Correccion Horas Extras cuando tenia menos de 30 dias trabajados


'Const Version = 1.18
'Const FechaVersion = "21/02/2008"
'Autor Modificación = Martin Ferraro
'Correccion a la version 1.15 cod siniestro
                
'Const Version = 1.19
'Const FechaVersion = "01/10/2008"
'Autor Modificación = Martin Ferraro
'En el campo localidad se trae provincia y localidad
                
'Const Version = 1.2
'Const FechaVersion = "06/10/2008"
'Autor Modificación = Martin Ferraro
'Se extendio el campo localidad y nuevo campo Aux_remimpo6 columna 79 confrep

'Const Version = 1.21
'Const FechaVersion = "01/12/2008"
'Adi_OS configurable en columna 78

'Const Version = 1.22
'Const FechaVersion = "03/04/2009" 'Martin Ferraro - SICOSS - Se agregaron tres nuevos campos configurables
'                                  Conceptos no remunerativos, Maternidad, Rectificacion de remuneracion
'                                  Se quito el TOPE de Rem1

'Const Version = 1.23
'Const FechaVersion = "03/06/2009" 'Modificado por MB porque no limpiaba esta variable aux_fecha
                                  'cuando cambiaba de empleado y si era una baja quedaba mal para el siguiente empleado

'Const Version = 1.24
'Const FechaVersion = "31/07/2009" 'Martin Ferraro - Encriptacion de string connection

'Const Version = 1.25
'Const FechaVersion = "07/09/2009" 'Martin Ferraro - Se inicializo en 0 los dias trabajados
                                  '                 Redondeo de la cant de horas si es > a 1. Sino 1

'Const Version = 1.26
'Const FechaVersion = "27/01/2010" 'Martin Ferraro - Version 33 SICOSS - Se Agregaron 3 Campos
                                  '                 Remuneración Imponible 9 X(9)
                                  '                 Contribución Tarea Diferencial(%) X(9)
                                  '                 Cantidad de Horas Trabajadas en el mes X(3)
                                  
'Const Version = 1.27
'Const FechaVersion = "07/05/2010" 'Martin Ferraro - Se aplicaron restricciones a los campos rem 6, 7, 8 y 9
                                  
'Const Version = 1.28
'Const FechaVersion = "22/07/2010" 'Martin Ferraro - SICOSS V.34 - Se agrego el campo de  Seguro colectivo de Vida Obligatorio
'                                                                 Cambio como aplica la restriccion a Aux_remImpo9

'Const Version = 1.29
'Const FechaVersion = "12/01/2011" 'MB - Se saco el topeo de 30 de dias trabajados para Teletech

'Const Version = 1.3
'Const FechaVersion = "27/05/2011" 'Si no tiene situación de revista, busca si tiene alguna estructura del tipo 'REV' en el confrep
                                'y le asigna el código a la situacion de revista
'Const Version = 1.31
'Const FechaVersion = "01/08/2011" 'Zamarbide Juan Alberto - CAS-13613 - CCU - Errores SICOSS - Sit. Revista.
                                  'Se corrigió el error que cuando existan más de 3 situaciones de revista, debería informar siempre las últimas 3 y no agruparlas,
                                  'salvo en casos que por procesos automáticos GIV se generen 2 del mismo tipo por períodos correlativos y agrupar esos períodos en 1.
                                  
'Const Version = 1.32
'Const FechaVersion = "19/08/2011" 'Zamarbide Juan Alberto - CAS-13613 - CCU - Errores SICOSS - Sit. Revista.
                                  'Correción errores varios
                                  
'Const Version = 1.33
'Const FechaVersion = "02/09/2011" 'Zamarbide Juan Alberto - CAS-13613 - CCU - Errores SICOSS - Sit. Revista.
'                                  'Se implemento la consulta para determinar si un empleado, posee en su estructura contrato un código 11 (Afip)
'                                  'Informe las situaciones de Revista por más de que tenga un cambio de fase en aluna fecha dentro del período informado.

'Const Version = "1.34"
'Const FechaVersion = "21/10/2011" 'FGZ - CAS-14114 - Citrusvil - Error en Version 1.33 de SICOSS.
'                                  'Se modificó la consulta para determinar si un empleado, posee en su estructura contrato un código 11 (Afip)
'                                  'el codigo es char.


'Const Version = "1.35"
'Const FechaVersion = "11/01/2012" 'FGZ - CAS-14792 - Coop. Seguros - Bug SICOSS.
                                  'Cuando hay + de una sit de revista y con el mismo codigo, estaba quedando mal.

'Const Version = "1.36"
'Const FechaVersion = "18/01/2012" 'JAZ - CAS-14826 - CCU - Error al generar Situacion de Revista
                                  'Se modificó la situación de revista para cuando tiene mas de 3 registros y cuando tiene 3 registros específicamente que devuelva la última sit. de revista actual
                                  
'Const Version = "1.37"
'Const FechaVersion = "08/02/2012" 'JAZ - CAS-15304 - Heidt & Asociados - Bug en generacion del SICOSS - R2
                                  'Se modificó la situación de revista para cuando tiene 3 registros para que devuelva la última sit. de revista actual netando las situaciones que sean iguales
                                  ' y dejando las situaciones en orden, por que salteaba la 2

'Const Version = "1.38"
'Const FechaVersion = "12/06/2012" 'Sebastian Stremel - Se inicializo variable - CAS-13764 - H&A - Bug en reporte SICOSS

'Const Version = "1.39"
'Const FechaVersion = "11/11/2013" 'Lisandro Moro - CAS-22321 - FUNDACION - ACTUALIZAR SICOSS - Correccion selects para Oracle

'Const Version = "1.40"
'Const FechaVersion = "07/08/2014" 'Dimatz Rafael - CAS-26666 - Se modifico para que ponga 1 si tiene TE Corresponde y 0 si tiene TE No Corresponde

'Const Version = "1.41"
'Const FechaVersion = "13/08/2014" 'Carmen Quintero - CAS-26384 - SANTANA - Error en Situacion de Revista Sicoss - Se modificó la logica cuando hay mas de 3 4iones de revistas.

'Const Version = "1.42"
'Const FechaVersion = "02/10/2014" 'Carmen Quintero - CAS-26384 - SANTANA - Error en Situacion de Revista Sicoss [Entrega 2] - Se agregó validacion para el caso cuando hay mas de 3 situaciones de revistas.


'Const Version = "1.43"
'Const FechaVersion = "14/10/2014" 'Fernandez, Matias - CAS-27548 - SANTANA TEXTILES - Error en situación de revista SICOSS multifase
                                  'cuando hay mas de una situacion de revista ACtivo en un mes se toma la fecha de inicio de la primera

'Const Version = "1.44"
'Const FechaVersion = "16/12/2014" 'Lisandro Moro - CAS-28427 - MONASTERIO BASE AMR - BUG EN SEGURO DE VIDA CUANDO TIENE DOS ESTRUCTURA

Const Version = "1.45"
Const FechaVersion = "28/01/2016" 'Gonzalez Nicolás - CAS-35372 - RHPro (Producto) - ARG - NOM - Bug reportes legales con sit de revista
                                  'Corrección de cód. a informar para cuando hay mas de 3 Situaciones de Revista
                                 

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
'Global nListaProc As Long
Global nListaProc As String
Global nEmpresa As Long
'----------------------------------------------------------

Global rs_Proc_Vol As New ADODB.Recordset
Global rs_Mod_Linea As New ADODB.Recordset
Global rs_Empleado As New ADODB.Recordset
Global rs_Mod_Asiento As New ADODB.Recordset

Global BUF_mod_linea As New ADODB.Recordset
Global BUF_temp As New ADODB.Recordset

Global Arrcod()
Global Arrdia()

Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Generador de Reporte SIJP.
' Autor      : FGZ
' Fecha      : 20/01/2003
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

    On Error GoTo ME_Main
    Nombre_Arch = PathFLog & "Generacion_Reporte_SIJP" & "-" & NroProcesoBatch & ".log"
    
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
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0 ,bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Pone el estado en procesando"
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 29 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        Call SIJP_06(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "no encontró el proceso"
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    
    If Not HuboError Then
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


Public Sub SIJP_06(ByVal bpronro As Long, ByVal Parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del reporte del SIJP
' Autor      : FGZ
' Fecha      : 20/01/2004
' Ult. Mod   : JAZ
' Fecha      : 01/08/2011
' --------------------------------------------------------------------------------------------
Dim Nroliq As Long
'Dim Todos_Pro As Boolean
'Dim Proc_Aprob As Integer
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
Dim arreglo(100)     As Double
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
Dim Aux_fecha As Date
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
Dim rs_aux As New ADODB.Recordset

'Declaracion de objeto recordset agregada por Hernán D. Santonocito - 03/03/2006
Dim rs_HisEstructuras As ADODB.Recordset
'-------------------------------------------------------------------------------

Dim TipoEstr As Long
Dim Opcion As Integer


Dim Aux_ApeNom As String
Dim Aux_Cod_Cont As String
Dim Aux_Cod As String
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

Dim Aux_ConcNoRem As String
Dim Aux_Maternidad As String
Dim Aux_RectifDeRem As String

Dim tcreduccion As Double
Dim Tiene_Convenio As Boolean

Dim Mensaje As String

Dim Restricciones(100) As TipoRestriccion
Dim LS_Res As Integer
Dim Aplica As Boolean
Dim Valor As Double
Dim indice As String
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
Dim Aux_remimpo6 As String

'Inicializacion de Variables
cant_dias_td = 0
cant_dias_total = 0
fecha_hasta_td = CDate("01/01/1900")
fec_hasta_total = CDate("01/01/1900")
tipdia_td = 0
tipdia_total = 0
nro_sitrev = 0
nro_sitrev_total = -1
topeArreglo = 100
suma_osocial = False
Dim resto19 As Double
Dim sueldo As Double
Dim hrsextras As Double
Dim zonadesf As Double
Dim adicional As Double
Dim premios As Double
Dim sac As Double
Dim lar As Double

Dim listaTipoLic As String

'Version SICOSS 33 -----------------------
Dim Aux_remImpo9 As String
Dim Aux_tareaDiff As String
Dim Aux_horTrabMes As String
'-----------------------------------------

'Version SICOSS 34 -----------------------
Dim Aux_SegVida As String

'Version 1.31 ----------------------------
Dim sr_cod(3) As String
Dim sr_dia(3) As String
Dim es_ultimo As Boolean
Dim p As Integer

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
        
        Nroliq = CLng(Mid(Parametros, pos1, pos2))
        Flog.writeline "Parametro NroLiq = " & Nroliq
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, Parametros, ".") - 1
        Lista_Pro = Mid(Parametros, pos1, pos2 - pos1 + 1)
        'MAF - Estaba truncando en 100
        Lista_Pro_F = Left(Lista_Pro, 1000)
        Flog.writeline "Parametro Lista_Pro = " & Lista_Pro
        ' esta lista tiene los nro de procesos separados por comas
        
        'Asigno el valor de lista de proceso a la variable global para poder usar en el SIJP
        'Agregado por Hernán D. Santonocito - 03/03/2006 -----------------------------------
        nListaProc = Lista_Pro
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
        
        pos1 = pos2 + 2
        pos2 = Len(Parametros)
        
        Flog.writeline "Posicion 1 = " & pos1
        Flog.writeline "Pos 2 = " & pos2
        Flog.writeline "Parametro Empresa " & Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Empresa = Mid(Parametros, pos1, pos2 - pos1 + 1)
        
        Flog.writeline "Parametro Empresa = " & Empresa
        ' GdeCos - Si el estrnro de la empresa es 0, entonces se buscan todas las empresas
        If Empresa = 0 Then
            Flog.writeline "Parametro Empresa = Todas"
        Else
            Flog.writeline "Parametro Empresa = " & Empresa
        End If
        
    End If
Else
    Flog.writeline "parametros nulos"
End If
Flog.writeline "terminó de levantar los parametros"

Empresa_Original = Empresa

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(Nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline "No se encontró el Periodo"
    Exit Sub
End If
Aux_fecha = rs_Periodo!pliqhasta
Fecha_Inicio_periodo = rs_Periodo!pliqdesde
Fecha_Fin_Periodo = rs_Periodo!pliqhasta

'Configuracion del Reporte
StrSql = "SELECT * FROM confrep WHERE repnro = 8 ORDER BY confnrocol"
OpenRecordset StrSql, rs_Confrep
If rs_Confrep.EOF Then
    Flog.writeline "No se encontró la configuración del Reporte"
    Exit Sub
End If
  

'14/01/2008 - Martin Ferraro - Cargo los tipos de lic para buscar la cond. siniest
Flog.writeline
listaTipoLic = ""
StrSql = "SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 8 "
StrSql = StrSql & " AND upper(conftipo) = 'LIC'"
StrSql = StrSql & " AND confnrocol = 73"
OpenRecordset StrSql, rs_ConfrepAux
If rs_ConfrepAux.EOF Then
    Flog.writeline "No se encuentra tipos de lic para el cod siniestro en la columna 73 tipo Lic del confrep"
Else
    Do While Not rs_ConfrepAux.EOF
        If listaTipoLic = "" Then
            listaTipoLic = rs_ConfrepAux!confval
        Else
            listaTipoLic = listaTipoLic & "," & rs_ConfrepAux!confval
        End If
        rs_ConfrepAux.MoveNext
    Loop
    
    Flog.writeline "Tipos de lic para cod siniestro = " & listaTipoLic
End If
rs_ConfrepAux.Close
Flog.writeline

'levanto las restricciones
StrSql = "SELECT * FROM confrep "
StrSql = StrSql & " WHERE repnro = 8 "
StrSql = StrSql & " AND upper(conftipo) = 'RES'"
OpenRecordset StrSql, rs_ConfrepAux
If rs_ConfrepAux.EOF Then
    Flog.writeline "No hay restricciones configuradas"
    LS_Res = -1
End If
I = 0
Do While Not rs_ConfrepAux.EOF
    Restricciones(I).Estrnro = IIf(Not EsNulo(rs_ConfrepAux!confval), rs_ConfrepAux!confval, 0)
    Restricciones(I).Valor = IIf(Not EsNulo(rs_ConfrepAux!confval2), CDbl(rs_ConfrepAux!confval2), 0)
    
    rs_ConfrepAux.MoveNext
    'If Not rs_Confrep.EOF Then
    If Not rs_ConfrepAux.EOF Then
        I = I + 1
    End If
Loop
LS_Res = I
  

'Depuracion del Temporario
'Borro todos los registros de la clave ingresada
StrSql = "DELETE FROM repsijp " & _
         " WHERE pliqnro = " & Nroliq & _
         " AND lista_pronro = '" & Lista_Pro_F & "'"
' GdeCos - 26/5/2005
If Empresa <> 0 Then
    StrSql = StrSql & " AND empresa = " & Empresa
End If

objConn.Execute StrSql, , adExecuteNoRecords
Flog.writeline "Borro todos los registros de la clave ingresada"

'StrSql = "SELECT pliqdesde, pliqhasta, proceso.pronro FROM periodo "
'StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
'StrSql = StrSql & " WHERE periodo.pliqnro =" & Nroliq
'StrSql = StrSql & " AND proceso.pronro IN (" & Lista_Pro & ")"
StrSql = "SELECT * FROM periodo "
StrSql = StrSql & " WHERE periodo.pliqnro =" & Nroliq
OpenRecordset StrSql, rs_Periodo

If Not rs_Periodo.EOF Then
    fechadesde = rs_Periodo!pliqdesde
    fechahasta = rs_Periodo!pliqhasta
End If

'-----------------------------------------------------------------------------------------------
' RCH - 18/10/2006
' Calculo del Mopre

Dim Valor_Ampo As Double
Dim Cant_Diaria_Ampos_1 As Double
Dim Ampo_Max_1 As Double

Dim rs_Ampo As New ADODB.Recordset

StrSql = "SELECT * FROM ampo WHERE ampofecha <= " & ConvFecha(fechahasta)
StrSql = StrSql & "  AND ampotconnro = 1 "
StrSql = StrSql & "  ORDER BY ampofecha desc "
OpenRecordset StrSql, rs_Ampo

Valor_Ampo = 0
Cant_Diaria_Ampos_1 = 0
Ampo_Max_1 = 0

If Not rs_Ampo.EOF Then
    Valor_Ampo = rs_Ampo!Valor
    Cant_Diaria_Ampos_1 = rs_Ampo!ampodiario
    Ampo_Max_1 = rs_Ampo!ampomax
End If
                
rs_Ampo.Close
Set rs_Ampo = Nothing

'-----------------------------------------------------------------------------------------------

Flog.writeline "MyBeginTrans"

'Comienzo la transaccion
MyBeginTrans

cont_lic = 0
UltimoEmpleado = -1


Do While Not rs_Periodo.EOF
    
    '11/01/2005 Maxi
    StrSql = "SELECT * FROM proceso "
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
    StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
    'StrSql = StrSql & " INNER JOIN his_estructura ON cabliq.empleado = his_estructura.ternro AND his_estructura.estrnro= " & Empresa & " AND htetdesde <= " & ConvFecha(fechahasta) & " AND (htethasta IS NULL OR htethasta >= " & ConvFecha(fechahasta) & ") AND his_estructura.tenro=10 "
    'StrSql = StrSql & " INNER JOIN fases ON fases.empleado = empleado.ternro"
    'StrSql = StrSql & " INNER JOIN his_estructura empresa ON empresa.ternro = empleado.ternro AND empresa.htetdesde >= fases.altfec"
    'StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = empresa.estrnro AND emp.estrnro =" & Empresa
    StrSql = StrSql & " WHERE proceso.pliqnro =" & rs_Periodo!pliqnro
    StrSql = StrSql & " AND proceso.pronro IN (" & Lista_Pro & ")"
    'StrSql = StrSql & " AND empresa.estrnro = " & Empresa
    'StrSql = StrSql & " AND empresa.tenro = 10 "
    'StrSql = StrSql & " AND empresa.htetdesde <= " & ConvFecha(fechahasta)
    'StrSql = StrSql & " AND (empresa.htethasta IS NULL OR empresa.htethasta >= " & ConvFecha(fechadesde) & ")"
    'StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(fechahasta)
    'StrSql = StrSql & " AND (fases.bajfec IS NULL OR fases.bajfec >= " & ConvFecha(fechadesde) & ")"
    StrSql = StrSql & " ORDER BY empleado.ternro, proceso.pronro"
    OpenRecordset StrSql, rs_Empleados
    
    If rs_Empleados.State = adStateOpen Then
        Flog.writeline "busco los empleados"
    Else
        Flog.writeline "se supero el tiempo de espera de la canculta de empleados"
        HuboError = True
    End If
    
    If Not HuboError Then
        'seteo de las variables de progreso
        Progreso = 0
        CConceptosAProc = rs_Periodo.RecordCount
        CEmpleadosAProc = rs_Empleados.RecordCount
        If CEmpleadosAProc = 0 Then
           Flog.writeline "no hay empleados"
           CEmpleadosAProc = 1
        End If
        IncPorc = ((100 / CEmpleadosAProc) * (100 / CConceptosAProc)) / 100
        'IncPorcEmpleado = (100 / CConceptosAProc)
        
        Flog.writeline
        Flog.writeline
        Flog.writeline
        Flog.writeline "------------------------------------------------------------------"
        
        Do While Not rs_Empleados.EOF
        
            Flog.writeline "Legajo: ------------------->" & rs_Empleados!empleg
            
            ' FGZ - 28/12/2004
            ' Si el empleado no esta activo ==> seteo la fecha de baja
            'If Not CBool(rs_Empleados!empest) Then
                             
               StrSql = " SELECT * FROM fases WHERE real = -1 AND empleado = " & rs_Empleados!Ternro & _
                        " AND (altfec <= " & ConvFecha(fechahasta) & ") " & _
                        " AND ((bajfec >= " & ConvFecha(fechadesde) & ")" & _
                        " OR (bajfec is null ))" & _
                        " ORDER BY altfec"
                
               OpenRecordset StrSql, rs_fases
               '---------inicio ver 1.34
               ' Creo el Select para verificar si el empleado tiene un Contrato de tipo 11 (Afip)
               
               StrSql = "SELECT * FROM empleado e,his_estructura he,estructura es,estr_cod ec, tipocod tc"
               StrSql = StrSql & " WHERE e.empleg =" & rs_Empleados!empleg
               'FGZ - 18/10/2011 -----------------------------------------
               'StrSql = StrSql & " AND nrocod = 11 AND es.tenro = 18"
               StrSql = StrSql & " AND nrocod = '11' AND es.tenro = 18"
               'FGZ - 18/10/2011 -----------------------------------------
               StrSql = StrSql & " AND he.ternro = e.ternro"
               StrSql = StrSql & " AND es.estrnro = he.estrnro"
               StrSql = StrSql & " AND es.estrnro = ec.estrnro"
               StrSql = StrSql & " AND ec.tcodnro = tc.tcodnro"
               ' Abo el recordset
               OpenRecordset StrSql, rs_aux
               'ver 1.34
               'Si dicho empleado tiene dicha estructura, asigno como fecha de inicio de Fase, el inicio del periodo, más allá de que el empleado tenga una fase abierta en dicho mes
               'Según resolución AFIP para el SICORE
               If Not rs_aux.EOF Then
                    Fecha_Inicio_Fase = Fecha_Inicio_periodo
                    Fecha_Fin_Fase = Fecha_Fin_Periodo
               Else
                    If rs_fases.RecordCount > 1 Then rs_fases.MoveFirst
                    If rs_fases.RecordCount > 0 Then
                            Flog.writeline "Comienza proceso de comparación de fechas de fases con las del período del SIJP"
                            Do While Not rs_fases.EOF
                                'Asigno la fecha de alta de la fase si es mayor a la del periodo
                                Fecha_Inicio_Fase = IIf(rs_fases!altfec > Fecha_Inicio_periodo, rs_fases!altfec, Fecha_Inicio_periodo)
                                Flog.writeline "Asigno a fecha de inicio de fase el valor " & Fecha_Inicio_Fase
                                If Not EsNulo(rs_fases!bajfec) Then
                                    Flog.writeline "El valor de fecha de baja no es nulo"
                                    'Asigno la fecha de baja de la fase si es menor a la del periodo
                                    Fecha_Fin_Fase = IIf(rs_fases!bajfec < Fecha_Fin_Periodo, rs_fases!bajfec, Fecha_Fin_Periodo)
                                    'If Fecha_Fin_Fase < Fecha_Inicio_Fase Then
                                    '    Fecha_Inicio_Fase = CDate("01/" & Month(Fecha_Fin_Fase) & "/" & Year(Fecha_Fin_Fase))
                                    'End If
                                    Flog.writeline "Asigno a fecha de fin de fase el valor " & Fecha_Fin_Fase
                                Else
                                    Fecha_Fin_Fase = Fecha_Fin_Periodo
                                    Flog.writeline "El valor de fecha de baja es nulo"
                                    Flog.writeline "El valor asignado a la Fecha de Fin de Fase es " & Fecha_Fin_Fase
                                End If
                                ' 04/06/2009 Modificado por MB porque no limpiaba esta variable aux_fecha
                                Aux_fecha = Fecha_Fin_Fase
                                Flog.writeline "Valor de Aux_Fecha: " & Aux_fecha
                                StrSql = " SELECT * FROM his_estructura "
                                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                                StrSql = StrSql & " his_estructura.estrnro = " & Empresa
                                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                                Flog.writeline "Hago la consulta sobre los históricos de estructura"
                                Flog.writeline "Consulta: " & StrSql
                                Set rs_HisEstructuras = New ADODB.Recordset
                                OpenRecordset StrSql, rs_HisEstructuras
                                If Err.Number <> 0 Then
                                    Flog.writeline "Error: " & Err.Number & " - Desc: " & Err.Description
                                    Err.Clear
                                End If
                                If rs_HisEstructuras.RecordCount > 1 Then rs_HisEstructuras.MoveFirst
                                Do While Not rs_HisEstructuras.EOF
                                    If rs_HisEstructuras!htetdesde <= Fecha_Inicio_Fase Then
                                        If rs_HisEstructuras!htethasta = Fecha_Fin_Fase Then
                                            Fecha_Inicio_Fase = Fecha_Inicio_periodo
                                            'Fecha_Fin_Fase = Fecha_Fin_Periodo
                                            Aux_fecha = Fecha_Fin_Fase
                                            Flog.writeline
                                            Flog.writeline "Fecha de alta de fase definitiva: " & Fecha_Inicio_Fase
                                            Flog.writeline "Fecha de baja de fase definitiva: " & Fecha_Fin_Fase
                                            Flog.writeline "Valor de Aux_Fecha: " & Aux_fecha
                                            Flog.writeline
                                            FechaAuxAsignada = True
                                            Exit Do
                                        End If
                                    End If
                                    rs_HisEstructuras.MoveNext
                                Loop
                                If rs_HisEstructuras.State = adStateOpen Then rs_HisEstructuras.Close
                                Set rs_HisEstructuras = Nothing
                                If Not FechaAuxAsignada Then
                                    rs_fases.MoveNext
                                Else
                                    Exit Do
                                End If
                        Loop
                    Else 'Si no encuentra fases para el empleado sigue con el proximo registro
                                'Martin Ferraro - Comente esta linea porque hacia movenext cuando
                                'era EOF
                                'rs_fases.MoveNext
                    End If
            End If '------ fin ver 1.34
            
            ' FGZ - 18/03/2004
            If rs_fases.State = adStateOpen Then rs_fases.Close
            Set rs_fases = Nothing
            
            ' ver 1.34
            If rs_aux.State = adStateOpen Then rs_aux.Close
            Set rs_aux = Nothing
            
            rs_Confrep.MoveFirst
            If rs_Empleados!Ternro <> UltimoEmpleado Then  'Es el primero
                
                UltimoEmpleado = rs_Empleados!Ternro
    
            
                'rs_Confrep.MoveFirst
                
                'Inicializar totales
                For contador = 1 To topeArreglo
                    arreglo(contador) = 0
                Next contador
                
                imp_oso = 0
                imp_ss = 0
                imp_ss_con = 0
                despedido = False
                desvinculado = False
                v_canthijos = 0
                v_conyuge = 0
            End If
            
            suma_osocial = False
            
            Do While Not rs_Confrep.EOF
            
                Select Case UCase(Trim(rs_Confrep!conftipo))
                Case "AC":
                    StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & rs_Empleados!cliqnro & _
                             " AND acunro =" & rs_Confrep!confval
                    OpenRecordset StrSql, rs_Acu_liq
                    If Not rs_Acu_liq.EOF Then
                        If rs_Confrep!confnrocol = 30 Or rs_Confrep!confnrocol = 40 Then
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Acu_liq!alcant
                        Else
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Acu_liq!almonto
                        End If
                    End If
                    
                Case "CO":
                    StrSql = "SELECT * FROM concepto "
                    StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                    StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                    OpenRecordset StrSql, rs_Conceptos
                    If Not rs_Conceptos.EOF Then
                        StrSql = "SELECT * FROM detliq WHERE concnro = " & rs_Conceptos!ConcNro & _
                                 " AND cliqnro =" & rs_Empleados!cliqnro
                        OpenRecordset StrSql, rs_Detliq
                        If Not rs_Detliq.EOF Then
                            If rs_Detliq!dlimonto <> 0 Then
                                arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlimonto
                            End If
                        End If
                    End If
                    
                Case "PCO":
                    StrSql = "SELECT * FROM concepto "
                    StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                    StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                    OpenRecordset StrSql, rs_Conceptos
                    If Not rs_Conceptos.EOF Then
                        StrSql = "SELECT * FROM detliq WHERE concnro = " & rs_Conceptos!ConcNro & _
                                 " AND cliqnro =" & rs_Empleados!cliqnro
                        OpenRecordset StrSql, rs_Detliq
                        If Not rs_Detliq.EOF Then
                            If rs_Detliq!dlicant <> 0 Then
                                'arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlicant
                                'FGZ - 08/04/2005 - los conceptos marcados como PCO no son acumulativos
                                arreglo(rs_Confrep!confnrocol) = rs_Detliq!dlicant
                            End If
                        End If
                    End If
                
                Case "CCO":
                    StrSql = "SELECT * FROM concepto "
                    StrSql = StrSql & " WHERE (concepto.conccod = " & rs_Confrep!confval
                    StrSql = StrSql & " OR concepto.conccod = '" & rs_Confrep!confval2 & "')"
                    OpenRecordset StrSql, rs_Conceptos
                    If Not rs_Conceptos.EOF Then
                        StrSql = "SELECT * FROM detliq WHERE concnro = " & rs_Conceptos!ConcNro & _
                                 " AND cliqnro =" & rs_Empleados!cliqnro
                        OpenRecordset StrSql, rs_Detliq
                        If Not rs_Detliq.EOF Then
                            If rs_Detliq!dlicant <> 0 Then
                                arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Detliq!dlicant
                            End If
                        End If
                    End If
                
                
                Case "IM": 'Imponibles
                    StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & rs_Empleados!cliqnro & _
                             " AND acunro =" & rs_Confrep!confval
                    OpenRecordset StrSql, rs_Impproarg
                    If Not rs_Impproarg.EOF Then
                        If rs_Confrep!confnrocol = 30 Or rs_Confrep!confnrocol = 40 Then
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Impproarg!ipacant
                        Else
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Impproarg!ipamonto
                        End If
                    End If
                
                Case "I1": 'Imponible Sueldo
                    StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & rs_Empleados!cliqnro & _
                             " AND acunro =" & rs_Confrep!confval & _
                             " AND tconnro = 1"
                    OpenRecordset StrSql, rs_Impproarg
                    If Not rs_Impproarg.EOF Then
                        If rs_Confrep!confnrocol = 30 Or rs_Confrep!confnrocol = 40 Then
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Impproarg!ipacant
                        Else
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Impproarg!ipamonto
                        End If
                    End If
                
                Case "I2": 'Imponible LAR
                    StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & rs_Empleados!cliqnro & _
                             " AND acunro =" & rs_Confrep!confval & _
                             " AND tconnro = 2"
                    OpenRecordset StrSql, rs_Impproarg
                    If Not rs_Impproarg.EOF Then
                        If rs_Confrep!confnrocol = 30 Or rs_Confrep!confnrocol = 40 Then
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Impproarg!ipacant
                        Else
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Impproarg!ipamonto
                        End If
                    End If
                
                Case "I3": 'Imponible SAC
                    StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & rs_Empleados!cliqnro & _
                             " AND acunro =" & rs_Confrep!confval & _
                             " AND tconnro = 3"
                    OpenRecordset StrSql, rs_Impproarg
                    If Not rs_Impproarg.EOF Then
                        If rs_Confrep!confnrocol = 30 Or rs_Confrep!confnrocol = 40 Then
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Impproarg!ipacant
                        Else
                            arreglo(rs_Confrep!confnrocol) = arreglo(rs_Confrep!confnrocol) + rs_Impproarg!ipamonto
                        End If
                    End If
                    
                Case "TE": 'Tipo Estructura
                    arreglo(rs_Confrep!confnrocol) = rs_Confrep!confval
                Case Else
                
                End Select
            
            
                rs_Confrep.MoveNext
            Loop
            
            
            'Reviso si es el ultimo empleado
            If EsElUltimoEmpleado(rs_Empleados, UltimoEmpleado) Then
    
                ' ----------------------------------------------------------------
                'Buscar el apellido y nombre
                StrSql = "SELECT * FROM tercero WHERE ternro = " & rs_Empleados!Ternro
                OpenRecordset StrSql, rs_Tercero
                If Not rs_Tercero.EOF Then
                    Aux_ApeNom = Left(rs_Tercero!terape & " " & rs_Tercero!ternom, 50)
                Else
                    Flog.writeline "No se encontró el tercero"
                    Exit Sub
                End If
                
                Flog.writeline "Busco las Estructuras con Aux_Fecha: " & Aux_fecha
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                ' Buscar el CONTRATO ACTUAL
                StrSql = " SELECT estrnro FROM his_estructura " & _
                         " WHERE ternro = " & rs_Empleados!Ternro & " AND " & _
                         " tenro = 18 AND " & _
                         " (htetdesde <= " & ConvFecha(Aux_fecha) & ") AND " & _
                         " ((" & ConvFecha(Aux_fecha) & " <= htethasta) or (htethasta is null))"
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    'Aux_Cod_Cont = Left(CStr(rs_Estructura!estrnro), 3)
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_Cont = Left(CStr(rs_Estr_cod!nrocod), 3)
                    Else
                        Flog.writeline "No se encontró el codigo interno para el Tipo de Contrato"
                        Aux_Cod_Cont = "0"
                    End If
                Else
                    Flog.writeline "No se encontró el Tipo de Contrato"
                    Aux_Cod_Cont = "0"
    '                Exit Sub
                End If
                                
                Flog.writeline "Buscar Situacion de Revista Actual"
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                'Buscar Situacion de Revista Actual
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 30 AND "
                'StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
                'StrSql = StrSql & " ((" & ConvFecha(Fecha_Inicio_periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                
                '-----------mdf ini
                'StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Fase) & ") AND "
                'StrSql = StrSql & " ((" & ConvFecha(Fecha_Inicio_Fase) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(Fecha_Inicio_periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                '-----------mdf fin


'                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(aux_fecha) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                         
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql, rs_Estructura
    
                Flog.writeline "inicializo"
                'FGZ - 29/06/2004
                Aux_Cod_sitr1 = ""
                Aux_diainisr1 = ""
                Aux_Cod_sitr2 = ""
                Aux_diainisr2 = ""
                Aux_Cod_sitr3 = ""
                Aux_diainisr3 = ""
                
                'FGZ - 29/06/2004
                
                Select Case rs_Estructura.RecordCount
                Case 0:
                        'EAM- Si no tiene situación de revista, busca si tiene alguna estructura 'REV' en el confrep
                        Aux_Cod_sitr1 = Buscar_SituacionRevistaConfig(rs_Empleados!Ternro, fechadesde, fechahasta)
                        Aux_diainisr1 = 1
                        If Aux_Cod_sitr1 <> 0 Then
                            Flog.writeline "Se asignó la situación de revista del confrep: " & Aux_Cod_sitr1
                            Aux_Cod_sitr = Aux_Cod_sitr1
                        Else
                            Flog.writeline "no hay situaciones de revista"
                            Aux_Cod_sitr = Aux_Cod_sitr1
                        End If
                Case 1:
                    'Aux_Cod_sitr1 = rs_Estructura!estrcodext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_sitr1 = Left(CStr(rs_Estr_cod!nrocod), 2)
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr1 = 1
                    End If
                    
                    If rs_Estructura!htetdesde < Fecha_Inicio_periodo Then
                        Aux_diainisr1 = 1
                    Else
                        Aux_diainisr1 = Day(rs_Estructura!htetdesde)
                    End If
                    'FGZ - 08/07/2005
                    If CInt(Aux_diainisr1) > Day(Fecha_Fin_Periodo) Then
                        Aux_diainisr1 = CStr(Day(Fecha_Fin_Periodo))
                    End If
                    
                    Aux_Cod_sitr = Aux_Cod_sitr1
                    Flog.writeline "hay 1 situaciones de revista"
                Case 2:
                    'Primer situacion
                    'Aux_Cod_sitr1 = rs_Estructura!estrcodext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_sitr1 = Left(CStr(rs_Estr_cod!nrocod), 2)
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr1 = 1
                    End If
                    
                    If rs_Estructura!htetdesde < Fecha_Inicio_periodo Then
                        Aux_diainisr1 = 1
                    Else
                        Aux_diainisr1 = Day(rs_Estructura!htetdesde)
                    End If
                    'FGZ - 08/07/2005
                    If CInt(Aux_diainisr1) > Day(Fecha_Fin_Periodo) Then
                        Aux_diainisr1 = CStr(Day(Fecha_Fin_Periodo))
                    End If
                    
                    'siguiente situacion
                    rs_Estructura.MoveNext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Aux_Cod_sitr1 <> Left(CStr(rs_Estr_cod!nrocod), 2) Then ' Agregado ver. 1.31
                            If Not rs_Estr_cod.EOF Then
                                Aux_Cod_sitr2 = Left(CStr(rs_Estr_cod!nrocod), 2)
                            Else
                                Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                                Aux_Cod_sitr2 = 1
                            End If
                            'Aux_Cod_sitr2 = rs_Estructura!estrcodext
                            Aux_diainisr2 = Day(rs_Estructura!htetdesde)
                            'FGZ - 08/07/2005
                            If CInt(Aux_diainisr2) > Day(Fecha_Fin_Periodo) Then
                                Aux_diainisr2 = CStr(Day(Fecha_Fin_Periodo))
                            End If
                            Aux_Cod_sitr = Aux_Cod_sitr2
                            Flog.writeline "hay 2 situaciones de revista"
                    Else 'FGZ - 11/01/2012 --------------------------------------------
                        'Si es la misma sit de revista ==> le asigno la anterior
                        Aux_Cod_sitr = Aux_Cod_sitr1
                    End If
                    
                Case 3:
                    'Primer situacion (1)
                    'Aux_Cod_sitr1 = rs_Estructura!estrcodext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_sitr1 = Left(CStr(rs_Estr_cod!nrocod), 2)
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                        Aux_Cod_sitr1 = 1
                    End If
                    
                    If rs_Estructura!htetdesde < Fecha_Inicio_periodo Then
                        Aux_diainisr1 = 1
                    Else
                        Aux_diainisr1 = Day(rs_Estructura!htetdesde)
                    End If
                    'FGZ - 08/07/2005
                    If CInt(Aux_diainisr1) > Day(Fecha_Fin_Periodo) Then
                        Aux_diainisr1 = CStr(Day(Fecha_Fin_Periodo))
                    End If
                    
                    'siguiente situacion (2)
                    rs_Estructura.MoveNext
                    'Aux_Cod_sitr2 = rs_Estructura!estrcodext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Aux_Cod_sitr1 <> Left(CStr(rs_Estr_cod!nrocod), 2) Then 'Agregado ver 1.31
                        If Not rs_Estr_cod.EOF Then
                            Aux_Cod_sitr2 = Left(CStr(rs_Estr_cod!nrocod), 2)
                        Else
                            Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                            Aux_Cod_sitr2 = 1
                        End If
                        Aux_diainisr2 = Day(rs_Estructura!htetdesde)
                        'FGZ - 08/07/2005
                        If CInt(Aux_diainisr2) > Day(Fecha_Fin_Periodo) Then
                            Aux_diainisr2 = CStr(Day(Fecha_Fin_Periodo))
                        End If
                    Else 'FGZ - 11/01/2012 --------------------------------------------
                        'Si es la misma sit de revista ==> le asigno la anterior
                        Aux_Cod_sitr = Aux_Cod_sitr1
                    End If
                    
                    'siguiente situacion (3)
                    rs_Estructura.MoveNext
                    'Aux_Cod_sitr3 = rs_Estructura!estrcodext
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Aux_Cod_sitr2 <> Left(CStr(rs_Estr_cod!nrocod), 2) Then 'Agregado ver 1.31
                        If Not rs_Estr_cod.EOF Then
                            If Aux_Cod_sitr2 <> "" Then 'Agregado ver 1.37 - JAZ
                                Aux_Cod_sitr3 = Left(CStr(rs_Estr_cod!nrocod), 2)
                            Else
                                Aux_Cod_sitr2 = Left(CStr(rs_Estr_cod!nrocod), 2)
                            End If
                        Else
                            Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                            Aux_Cod_sitr3 = 1
                        End If
                        If Aux_Cod_sitr3 <> "" Then 'Agregado ver 1.37 - JAZ
                            Aux_diainisr3 = Day(rs_Estructura!htetdesde)
                        Else
                            Aux_diainisr2 = Day(rs_Estructura!htetdesde)
                        End If
                        'FGZ - 08/07/2005
                        If Aux_diainisr3 <> "" Then 'Agregado ver 1.37 - JAZ
                            If CInt(Aux_diainisr3) > Day(Fecha_Fin_Periodo) Then
                                Aux_diainisr3 = CStr(Day(Fecha_Fin_Periodo))
                            End If
                            Aux_Cod_sitr = Aux_Cod_sitr3
                        Else
                            If CInt(Aux_diainisr2) > Day(Fecha_Fin_Periodo) Then
                                Aux_diainisr2 = CStr(Day(Fecha_Fin_Periodo))
                            End If
                            Aux_Cod_sitr = Aux_Cod_sitr2
                        End If
                        Flog.writeline "hay 3 situaciones de revista"
                    Else 'FGZ - 11/01/2012 --------------------------------------------
                        'Si es la misma sit de revista ==> le asigno la anterior
                        If Aux_Cod_sitr2 <> "" Then 'Modificado ver 1.36 - JAZ
                            Aux_Cod_sitr = Aux_Cod_sitr2
                        Else
                            Aux_Cod_sitr = Aux_Cod_sitr1
                        End If
                    End If
                    
                Case Else 'mas de tres situaciones ==> toma las ultimas tres pero verifica que no haya situaciones iguales en dif periodos
'                    Agregado 13/08/2014
                     If Not rs_Estructura.EOF Then
                        Dim k
                        k = 0
                        Do While Not rs_Estructura.EOF
                           If (k = 0) Then
                                ReDim Preserve Arrcod(k)
                                ReDim Preserve Arrdia(k)
                                
                                'Arrcod(k) = rs_Estructura!estrcodext 'NG - v 1.45
                                
                                StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                                StrSql = StrSql & " AND tcodnro = 1"
                                OpenRecordset StrSql, rs_Estr_cod
                                If Not rs_Estr_cod.EOF Then
                                    Arrcod(k) = Left(CStr(rs_Estr_cod!nrocod), 2)
                                Else
                                    Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                                    Arrcod(k) = 1
                                End If
                                
                                
                                'Arrdia(k) = Day(rs_Estructura!htetdesde)
                                'Agregado 02/10/2014
                                If rs_Estructura!htetdesde < Fecha_Inicio_periodo Then
                                   Arrdia(k) = 1
                                Else
                                   Arrdia(k) = Day(rs_Estructura!htetdesde)
                                End If
                                'fin
                                k = k + 1
                           Else
                           
                                'NG - v 1.45------------------------------------------------------------------------
                                Aux_Cod = ""
                                StrSql = "SELECT nrocod FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                                StrSql = StrSql & " AND tcodnro = 1"
                                OpenRecordset StrSql, rs_Estr_cod
                                If Not rs_Estr_cod.EOF Then
                                    Aux_Cod = Left(CStr(rs_Estr_cod!nrocod), 2)
                                Else
                                    Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
                                    Aux_Cod = 1
                                End If
                                '----------------------------------------------------------------------------------

                           
                               If Arrcod(k - 1) <> rs_Estructura!estrcodext Then
                                    ReDim Preserve Arrcod(k)
                                    ReDim Preserve Arrdia(k)
                                    'Arrcod(k) = rs_Estructura!estrcodext 'NG - v 1.45
                                    Arrcod(k) = Aux_Cod
                                    
                                    If rs_Estructura!htetdesde < Fecha_Inicio_periodo Then
                                        Arrdia(k) = 1
                                    Else
                                        Arrdia(k) = Day(rs_Estructura!htetdesde)
                                    End If
                                    k = k + 1
                               End If
                           End If
          
                           rs_Estructura.MoveNext
                        Loop
                     End If
                     
                     If UBound(Arrcod) >= 2 Then
                        Aux_Cod_sitr3 = Arrcod(UBound(Arrcod))
                        Aux_Cod_sitr2 = Arrcod(UBound(Arrcod) - 1)
                        Aux_Cod_sitr1 = Arrcod(UBound(Arrcod) - 2)
                        Aux_diainisr3 = Arrdia(UBound(Arrdia))
                        Aux_diainisr2 = Arrdia(UBound(Arrdia) - 1)
                        Aux_diainisr1 = Arrdia(UBound(Arrdia) - 2)
                     End If
                    
                     If UBound(Arrcod) = 1 Then
                        Aux_Cod_sitr3 = ""
                        Aux_Cod_sitr2 = Arrcod(UBound(Arrcod))
                        Aux_Cod_sitr1 = Arrcod(UBound(Arrcod) - 1)
                        Aux_diainisr3 = ""
                        Aux_diainisr2 = Arrdia(UBound(Arrdia))
                        Aux_diainisr1 = Arrdia(UBound(Arrdia) - 1)
                     End If
                    
                     If UBound(Arrcod) = 0 Then
                        Aux_Cod_sitr3 = ""
                        Aux_Cod_sitr2 = ""
                        Aux_Cod_sitr1 = Arrcod(UBound(Arrcod))
                        Aux_diainisr3 = ""
                        Aux_diainisr2 = ""
                        Aux_diainisr1 = Arrdia(UBound(Arrdia))
                     End If
                     
                     Aux_Cod_sitr = Arrcod(UBound(Arrcod))
                    'fin
                    

                     
                     Flog.writeline "hay + de 3 situaciones de revista"

   '--------------JAZ Comentado 01-08-2011 Ver 1.31 Caso CAS - 13613 ----------------------------
  '                  rs_Estructura.MoveLast
  '                  rs_Estructura.MovePrevious
  '                  rs_Estructura.MovePrevious
                    
                    'Primer situacion (1)
                    'Aux_Cod_sitr1 = rs_Estructura!estrcodext
  '                  StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
  '                  StrSql = StrSql & " AND tcodnro = 1"
  '                  OpenRecordset StrSql, rs_Estr_cod
  '                  If Not rs_Estr_cod.EOF Then
  '                      Aux_Cod_sitr1 = Left(CStr(rs_Estr_cod!nrocod), 2)
  '                  Else
  '                      Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
  '                      Aux_Cod_sitr1 = 1
  '                  End If
  '
  '                  If rs_Estructura!htetdesde < Fecha_Inicio_periodo Then
  '                      Aux_diainisr1 = 1
  '                  Else
  '                      Aux_diainisr1 = Day(rs_Estructura!htetdesde)
  '                  End If
  '                  'FGZ - 08/07/2005
  '                  If CInt(Aux_diainisr1) > Day(Fecha_Fin_Periodo) Then
  '                      Aux_diainisr1 = CStr(Day(Fecha_Fin_Periodo))
  '                  End If
                    
                    'siguiente situacion (2)
  '                  rs_Estructura.MoveNext
                    'Aux_Cod_sitr2 = rs_Estructura!estrcodext
  '                  StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
  '                  StrSql = StrSql & " AND tcodnro = 1"
  '                  OpenRecordset StrSql, rs_Estr_cod
  '                  If Not rs_Estr_cod.EOF Then
  '                      Aux_Cod_sitr2 = Left(CStr(rs_Estr_cod!nrocod), 2)
  '                  Else
  '                      Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
  '                      Aux_Cod_sitr2 = 1
  '                  End If
                    
  '                  Aux_diainisr2 = Day(rs_Estructura!htetdesde)
  '                  'FGZ - 08/07/2005
  '                  If CInt(Aux_diainisr2) > Day(Fecha_Fin_Periodo) Then
  '                      Aux_diainisr2 = CStr(Day(Fecha_Fin_Periodo))
  '                  End If
                
                    'siguiente situacion (3)
  '                  rs_Estructura.MoveNext
                    'Aux_Cod_sitr3 = rs_Estructura!estrcodext
  '                  StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
  '                  StrSql = StrSql & " AND tcodnro = 1"
  '                  OpenRecordset StrSql, rs_Estr_cod
  '                  If Not rs_Estr_cod.EOF Then
  '                      Aux_Cod_sitr3 = Left(CStr(rs_Estr_cod!nrocod), 2)
  '                  Else
  '                      Flog.writeline "No se encontró el codigo interno para la Situacion de Revista"
  '                      Aux_Cod_sitr3 = 1
  '                  End If
                    
  '                  Aux_diainisr3 = Day(rs_Estructura!htetdesde)
                    'FGZ - 08/07/2005
  '                  If CInt(Aux_diainisr3) > Day(Fecha_Fin_Periodo) Then
  '                      Aux_diainisr3 = CStr(Day(Fecha_Fin_Periodo))
  '                  End If
                
  '                  Aux_Cod_sitr = Aux_Cod_sitr3
                    
  '                  Flog.writeline "hay + de 3 situaciones de revista"
  '------------------------- FIN COMENTARIO ---------------------------------
                End Select
                
                'FGZ - 28/12/2004
                'No puede haber situaciones de revista iguales consecutivas.
                'Antes ese caso, me quedo con la primera de las iguales y consecutivas
                If Aux_Cod_sitr3 = Aux_Cod_sitr2 Then
                    'Elimino la situacion de revista 3
                    Aux_Cod_sitr3 = ""
                    Aux_diainisr3 = ""
                End If
                If Aux_Cod_sitr2 = Aux_Cod_sitr1 Then
                    'Elimino la situacion de revista 2 y la 3 la pongo en la 2
                    Aux_Cod_sitr2 = Aux_Cod_sitr3
                    Aux_diainisr2 = Aux_diainisr3
                    
                    Aux_Cod_sitr3 = ""
                    Aux_diainisr3 = ""
                End If
                
                
    'per.repsijp.cod-sitr = substring(string(tt-sitrev.codigo,">9") + fill(" ",2),1,2).
    'Aux_Cod_sitr = Aux_Cod_sitr1
    
    '            cant_sr = 3
    '            Do While Not rs_Estructura.EOF And cant_sr > 0
    '                If rs_Estructura!htetdesde < Fecha_Inicio_periodo Then
    '
    '                End If
    '
    '                rs_Estructura.MoveNext
    '            Loop
            
            
            
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                ' Buscar la condicion
                Flog.writeline "Buscar la condicion"
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 31 AND "
'                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_fecha) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(Aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_Cond = Left(CStr(rs_Estr_cod!nrocod), 2)
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Condicion de SIJP"
                        Aux_Cod_Cond = "0"
                    End If
                    'Aux_Cod_Cond = Format(rs_Estructura!estrcodext, "#0")
                Else
                    Flog.writeline "No se encontro la Condicion de SIJP"
                    Aux_Cod_Cond = "0"
                End If
                
                ' ----------------------------------------------------------------
                ' Buscar el CUIL
                Flog.writeline "Buscar el CUIL"
                StrSql = " SELECT cuil.nrodoc FROM tercero " & _
                         " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) " & _
                         " WHERE tercero.ternro= " & rs_Empleados!Ternro
                OpenRecordset StrSql, rs_Cuil
                If Not rs_Cuil.EOF Then
                    Aux_CUIL = Left(CStr(rs_Cuil!NroDoc), 13)
                    If Cuil_Valido(Aux_CUIL, Mensaje) Then
                        Aux_CUIL = Replace(CStr(Aux_CUIL), "-", "")
                    Else
                        Flog.writeline "Error al obtener los datos del cuil. " & Mensaje & Aux_CUIL
                        Aux_CUIL = ""
                    End If
                Else
                    Flog.writeline "Error al obtener los datos del cuil"
                    Aux_CUIL = ""
                End If
    
                ' ----------------------------------------------------------------
                ' Buscar la ZONA de la sucursal
                Flog.writeline "Buscar la Zona de la Sucursal"
                Aux_Zona = "00"
                
                'Busco el tipo en la configuracion del reporte
                StrSql = "SELECT * FROM confrep WHERE repnro = 8 " & _
                         " AND confnrocol = 41"
                If Aux_rs_Confrep.State = adStateOpen Then Aux_rs_Confrep.Close
                OpenRecordset StrSql, Aux_rs_Confrep
                
                If Aux_rs_Confrep.EOF Then
                    Flog.writeline "No se encontró la configuración del Reporte para la columna 41 (Zona)"
                    'Por defaul busco la del empleado
                Else
                    Opcion = Aux_rs_Confrep!confval
                End If
    
                ' De acuerdo a la opcion busco la zona
                Select Case Opcion
                Case 1: 'Sucursal
                    'Cargo el tipo de estructura segun sea sucursal
                    TipoEstr = 10
                
                    StrSql = " SELECT estrnro FROM his_estructura " & _
                             " WHERE ternro = " & rs_Empleados!Ternro & " AND " & _
                             " tenro = 1 AND " & _
                             " (htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND " & _
                             " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= htethasta) or (htethasta is null))"
                    OpenRecordset StrSql, rs_Estructura
                
                    If Not rs_Estructura.EOF Then
                        StrSql = " SELECT * FROM sucursal " & _
                                 " WHERE estrnro =" & rs_Estructura!Estrnro
                        OpenRecordset StrSql, rs_sucursal
                        
                        If Not rs_sucursal.EOF Then
                            
                            'Aux_Reduccion = Format(CStr(rs_sucursal!sucporred), 0)
                            Aux_Reduccion = 0
                            
                            StrSql = " SELECT zona.zonanro,zona.zonacod,locdesc,provdesc FROM detdom " & _
                                     " INNER JOIN zona ON zona.zonanro = detdom.zonanro " & _
                                     " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                                     " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
                                     " LEFT JOIN provincia ON provincia.provnro = detdom.provnro " & _
                                     " WHERE cabdom.ternro = " & rs_sucursal!Ternro '& " AND " & _
                                     '" cabdom.tipnro =" & TipoEstr
                            OpenRecordset StrSql, rs_zona
                            If Not rs_zona.EOF Then
                                Aux_Zona = Left(CStr(IIf(Not IsNull(rs_zona!zonacod), rs_zona!zonacod, "00")), 2)
                                Aux_Localidad = IIf(EsNulo(rs_zona!provdesc), " - ", rs_zona!provdesc & " - ")
                                Aux_Localidad = Aux_Localidad & IIf(EsNulo(rs_zona!locdesc), "", rs_zona!locdesc)
                            Else
                                Flog.writeline "No se encuentra zona laboral. SQL : " & StrSql
                            End If
                        End If  ' If Not rs_Sucursal.EOF Then
                    End If ' If Not rs_Estructura.EOF Then
                Case 2: 'Empresa
                    'Cargo el tipo de estructura segun sea sucursal o Empresa
                    TipoEstr = 12
                    StrSql = " SELECT estrnro FROM his_estructura " & _
                             " WHERE ternro = " & rs_Empleados!Ternro & " AND " & _
                             " tenro = 10 AND " & _
                             " (htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND " & _
                             " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= htethasta) or (htethasta is null))"
                    OpenRecordset StrSql, rs_Estructura
                
                    If Not rs_Estructura.EOF Then
                        StrSql = " SELECT * FROM empresa " & _
                                 " WHERE estrnro =" & rs_Estructura!Estrnro
                        OpenRecordset StrSql, rs_sucursal
                        
                        If Not rs_sucursal.EOF Then
                            
                            'Aux_Reduccion = Format(CStr(rs_sucursal!sucporred), 0)
                            Aux_Reduccion = 0
                            
                            StrSql = " SELECT zona.zonanro,zona.zonacod,locdesc,provdesc FROM detdom " & _
                                     " INNER JOIN zona ON zona.zonanro = detdom.zonanro " & _
                                     " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                                     " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
                                     " LEFT JOIN provincia ON provincia.provnro = detdom.provnro " & _
                                     " WHERE cabdom.ternro = " & rs_sucursal!Ternro '& " AND " & _
                                     '" cabdom.tipnro =" & TipoEstr
                            OpenRecordset StrSql, rs_zona
                            If Not rs_zona.EOF Then
                                'Aux_Zona = Left(CStr(rs_zona!zonanro), 2)
                                Aux_Zona = Left(CStr(IIf(Not IsNull(rs_zona!zonacod), rs_zona!zonacod, "00")), 2)
                                Aux_Localidad = IIf(EsNulo(rs_zona!provdesc), " - ", rs_zona!provdesc & " - ")
                                Aux_Localidad = Aux_Localidad & IIf(EsNulo(rs_zona!locdesc), "", rs_zona!locdesc)
                            Else
                                Flog.writeline "No se encuentra zona laboral. SQL : " & StrSql
                            End If
                        End If  ' If Not rs_Sucursal.EOF Then
                    End If ' If Not rs_Estructura.EOF Then
                Case 3: 'Empleado
                    TipoEstr = 1
                    StrSql = " SELECT zona.zonanro,zona.zonacod,locdesc,provdesc FROM detdom " & _
                             " INNER JOIN zona ON zona.zonanro = detdom.zonanro " & _
                             " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                             " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
                             " LEFT JOIN provincia ON provincia.provnro = detdom.provnro " & _
                             " WHERE cabdom.ternro = " & rs_Empleados!Ternro '& " AND " & _
                             '" cabdom.tidonro =" & TipoEstr
                    OpenRecordset StrSql, rs_zona
                    If Not rs_zona.EOF Then
                        'Aux_Zona = Left(CStr(rs_zona!zonanro), 2)
                        Aux_Zona = Left(CStr(IIf(Not IsNull(rs_zona!zonacod), rs_zona!zonacod, "00")), 2)
                                Aux_Localidad = IIf(EsNulo(rs_zona!provdesc), " - ", rs_zona!provdesc & " - ")
                                Aux_Localidad = Aux_Localidad & IIf(EsNulo(rs_zona!locdesc), "", rs_zona!locdesc)
                    Else
                        Flog.writeline "No se encuentra zona laboral. SQL : " & StrSql
                    End If
                Case 4: 'Laboral
                    TipoEstr = 3
                    StrSql = " SELECT zona.zonanro,zona.zonacod,locdesc,provdesc FROM detdom " & _
                             " INNER JOIN zona ON zona.zonanro = detdom.zonanro " & _
                             " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                             " INNER JOIN localidad ON localidad.locnro = detdom.locnro " & _
                             " LEFT JOIN provincia ON provincia.provnro = detdom.provnro " & _
                             " WHERE cabdom.ternro = " & rs_Empleados!Ternro '& " AND " & _
                             '" cabdom.tidonro = " & TipoEstr
                    OpenRecordset StrSql, rs_zona
                    If Not rs_zona.EOF Then
                        'Aux_Zona = Left(CStr(rs_zona!zonanro), 2)
                        Aux_Zona = Left(CStr(IIf(Not IsNull(rs_zona!zonacod), rs_zona!zonacod, "00")), 2)
                        Aux_Localidad = IIf(EsNulo(rs_zona!provdesc), " - ", rs_zona!provdesc & " - ")
                        Aux_Localidad = Aux_Localidad & IIf(EsNulo(rs_zona!locdesc), "", rs_zona!locdesc)
                    Else
                        Flog.writeline "No se encuentra zona laboral. SQL : " & StrSql
                    End If
                End Select
                Aux_Zona = Format(Aux_Zona, "00")
    
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                ' Buscar la Obra Social del empleado
                Flog.writeline "Buscar la Obra Social del empleado"
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 17 AND "
'                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_fecha) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(Aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Cod_Obra_Social = Left(rs_Estr_cod!nrocod, 6)
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Obra Social"
                        Aux_Cod_Obra_Social = "0"
                    End If
                    'Aux_Cod_Obra_Social = Format(rs_Estructura!estrcodext, "#####0")
                Else
                    Flog.writeline "No se encontro la Obra Social"
                    Aux_Cod_Obra_Social = "0"
                End If
                
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                ' Verificar si es un Contrato para Jubilados, de ser asi tomar ese codigo
                Flog.writeline "Verificar si es un Contrato para Jubilados, de ser asi tomar ese codigo"
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 18 AND "
'                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_fecha) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(Aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Contrato_Actual = Left(rs_Estr_cod!nrocod, 6)
                    Else
                        Flog.writeline "No se encontró el codigo interno para el contrato actual"
                        Aux_Contrato_Actual = "0"
                    End If
                    
                    'Aux_Contrato_Actual = Format(rs_Estructura!estrcodext, "#####0")
                    
                    'tipo de contrato (tcreduccion)
                    StrSql = " SELECT * FROM tipocont " & _
                             " WHERE tipocont.estrnro = " & rs_Estructura!Estrnro
                    OpenRecordset StrSql, rs_TipoCont
                    If Not rs_TipoCont.EOF Then
                        If Not IsNull(rs_TipoCont!tcreduccion) Then
                            tcreduccion = Left(CStr(rs_TipoCont!tcreduccion), 3)
                        Else
                            tcreduccion = 0
                        End If
                    Else
                        Flog.writeline "No se encontró el tipo de contrato de la empresa"
                        tcreduccion = "0"
                    End If
                Else
                    Flog.writeline "No se encontro el tipo de Contrato Actual"
                    Aux_Contrato_Actual = "0"
                    tcreduccion = 0
    '                Exit Sub
                End If
                
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                ' Buscar la actividad del empleado
                Flog.writeline "Buscar la actividad del empleado"
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 29 AND "
'                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_fecha) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(Aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_Actividad = Left(rs_Estr_cod!nrocod, 2)
                    Else
                        Flog.writeline "No se encontró el codigo interno para la Actividad del empleado"
                        Aux_Actividad = "0"
                    End If
                
                    'Aux_Actividad = Format(rs_Estructura!estrcodext, "#0")
                Else
                    Flog.writeline "No se encontro la Actividad"
                    Aux_Actividad = "00"
                End If
                           
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                ' Buscar el Tipo de Empleador
                Flog.writeline "Buscar el Tipo de Empleador"
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 10 AND "
'                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_fecha) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(Aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    ' GdeCos - Seteo el nombre de la empresa a la que pertenece el empleado
                    If Empresa = 0 Then
                        Empresa = rs_Estructura!Estrnro
                    End If
                    
                    StrSql = " SELECT * FROM empresa " & _
                             " INNER JOIN tipempdor ON empresa.tipempnro = tipempdor.tipempnro " & _
                             " WHERE empresa.estrnro = " & rs_Estructura!Estrnro
                    OpenRecordset StrSql, rs_Empresa
                    If Not rs_Empresa.EOF Then
                        Aux_TipEmpNro = Format(rs_Empresa!tipempcoddgi, "0")
                    Else
                        Flog.writeline "No se encontró la empresa en tipempdor"
                        Aux_TipEmpNro = "0"
                    End If
    '                StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
    '                StrSql = StrSql & " AND tcodnro = 1"
    '                OpenRecordset StrSql, rs_Estr_cod
    '                If Not rs_Estr_cod.EOF Then
    '                    Aux_TipEmpNro = Left(CStr(rs_Estr_cod!nrocod), 1)
    '                Else
    '                    Flog.writeline "No se encontró el codigo interno para la Empresa"
    '                    Aux_TipEmpNro = "0"
    '                End If
                    
                Else
                    Flog.writeline "No tiene empresa asociada en ese periodo"
                    Aux_TipEmpNro = "0"
                End If
                
                ' ----------------------------------------------------------------
                ' Buscar Cant. de Hijos
                Flog.writeline "Buscar Cant. de Hijos"
                StrSql = " SELECT familiar.ternro FROM familiar " & _
                         " WHERE familiar.empleado = " & rs_Empleados!Ternro & " AND " & _
                         " familiar.parenro =2"
                OpenRecordset StrSql, rs_Familia
                Do While Not rs_Familia.EOF
                    v_canthijos = v_canthijos + 1
        
                    rs_Familia.MoveNext
                Loop
                
                ' ----------------------------------------------------------------
                ' Buscar Conyugue
                Flog.writeline "Buscar Conyugue"
                StrSql = " SELECT familiar.ternro FROM familiar " & _
                         " WHERE familiar.empleado = " & rs_Empleados!Ternro & " AND " & _
                         " familiar.parenro =3"
                If rs_Familia.State = adStateOpen Then rs_Familia.Close
                OpenRecordset StrSql, rs_Familia
    
                If Not rs_Familia.EOF Then
                    v_conyuge = 1
                End If
    
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                ' Buscar el codigo de Siniestrado del empleado
'                Flog.writeline "Buscar el codigo de Siniestrado del empleado"
'                StrSql = " SELECT * FROM his_estructura "
'                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
'                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!ternro & " AND "
'                StrSql = StrSql & " his_estructura.tenro = 42 AND "
''                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
''                StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
'                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(aux_fecha) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
'                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
'                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
'                OpenRecordset StrSql, rs_Estructura
'                If Not rs_Estructura.EOF Then
'                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
'                    StrSql = StrSql & " AND tcodnro = 1"
'                    OpenRecordset StrSql, rs_Estr_cod
'                    If Not rs_Estr_cod.EOF Then
'                        Aux_codsiniestro = Left(CStr(rs_Estr_cod!nrocod), 2)
'                    Else
'                        Flog.writeline "No se encontró el codigo de Siniestro SIJP"
'                        Aux_codsiniestro = "0"
'                    End If
'
'                    'Aux_codsiniestro = Format(rs_Estructura!estrcodext, "#0")
'                Else
'                    Flog.writeline "No se encontro el codigo del siniestrado del empleado"
'                    Aux_codsiniestro = "00"
'                End If
                
                '14/01/2008 - Martin Ferraro - Codigo de Siniestro = 01 si existe licencia de tipo confrep
                'en el rango de fechas del periodo. Sino Codigo de Siniestro = 01
                Aux_codsiniestro = "00"
                If Len(listaTipoLic) > 0 Then
                    StrSql = " SELECT * FROM emp_lic"
                    StrSql = StrSql & " WHERE emp_lic.tdnro in (" & listaTipoLic & ")"
                    StrSql = StrSql & " AND ("
                    StrSql = StrSql & " ((emp_lic.elfechadesde <= " & ConvFecha(Fecha_Inicio_periodo) & ") AND"
                    StrSql = StrSql & " (emp_lic.elfechahasta >= " & ConvFecha(Fecha_Inicio_periodo) & "))"
                    StrSql = StrSql & " OR"
                    StrSql = StrSql & " ((emp_lic.elfechadesde >= " & ConvFecha(Fecha_Inicio_periodo) & ") AND"
                    StrSql = StrSql & " (emp_lic.elfechadesde <= " & ConvFecha(Fecha_Fin_Periodo) & "))"
                    StrSql = StrSql & " )"
                    StrSql = StrSql & " AND emp_lic.licestnro = 2"
                    StrSql = StrSql & " AND emp_lic.empleado = " & rs_Empleados!Ternro
                    OpenRecordset StrSql, rs_Estructura
                    If Not rs_Estructura.EOF Then
                        Aux_codsiniestro = "01"
                        Flog.writeline "Codigo siniestro: Licencias encontradas para el empleado"
                    Else
                        Aux_codsiniestro = "00"
                        Flog.writeline "Codigo siniestro: No se encontraron los tipos de licencia " & listaTipoLic & " para el empleado"
                    End If
                    rs_Estructura.Close
                End If
                
                
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                ' Version 19  Buscar regimen
                Flog.writeline "Version 19  Buscar regimen"
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 15 AND "
'                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_fecha) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(Aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    Flog.writeline "Se encontro la estructura de Caja de jubilacion " & rs_Estructura!Estrnro
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!Estrnro
                    StrSql = StrSql & " AND tcodnro = 1"
                    OpenRecordset StrSql, rs_Estr_cod
                    If Not rs_Estr_cod.EOF Then
                        Aux_con_FNE = Format(rs_Estr_cod!nrocod, "0")
                    Else
                        Flog.writeline "No se encontró el codigo de Caja de Jubilacion SIJP"
                        Aux_con_FNE = "0"
                    End If
                
                    'Aux_con_FNE = Format(rs_Estructura!estrcodext, "0")
                Else
                    Flog.writeline "No se encontro el Regimen"
                    Aux_con_FNE = "0"
                End If
               
                ' ----------------------------------------------------------------
                ' FGZ - 28/04/2004 - Codigos
                ' Buscar si es convencionado
                Flog.writeline "Buscar si es convencionado"
                StrSql = " SELECT * FROM his_estructura "
                StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                StrSql = StrSql & " his_estructura.tenro = 19 AND "
'                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Fecha_Fin_Periodo) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(Fecha_Fin_Periodo) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_fecha) & ") AND "
                StrSql = StrSql & " ((" & ConvFecha(Aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                OpenRecordset StrSql, rs_Estructura
                If Not rs_Estructura.EOF Then
                    StrSql = "SELECT * FROM convenios WHERE estrnro =" & rs_Estructura!Estrnro
                    OpenRecordset StrSql, rs_Convenios
                    'Version 20 - Octubre 2002 - Modificar con toda la estructura de situacion de revista
                    If Not rs_Convenios.EOF Then
                        Flog.writeline "Antes de preguntar por conveniado"
                       Aux_XConv = IIf(CBool(IIf(Not IsNull(rs_Convenios!conveniado), rs_Convenios!conveniado, 0)), "1", "0")
                       Flog.writeline "despues de preguntar por conveniado"
                    Else
                        Aux_XConv = "0"
                    End If
                    Tiene_Convenio = True
                Else
                    Tiene_Convenio = False
                    Aux_XConv = "0"
                    Flog.writeline "No se encontro el Convenio"
                End If
                    
        '            If Tiene_Convenio Then
        '               Aux_XConv = "1"
        '            Else
        '                Aux_XConv = "0"
        '            End If
                
                ' ----------------------------------------------------------------
                ' MAURICIO 2/11/99   TOMAR IMPONIBLES DE LOS ACUMULADORES CORRESPONDIENTES
                Flog.writeline "MAURICIO 2/11/99   TOMAR IMPONIBLES DE LOS ACUMULADORES CORRESPONDIENTES"
                bruto = IIf(arreglo(18) = 0 And (CInt(Aux_Cod_sitr) <> 5 And CInt(Aux_Cod_sitr) <> 6), 0, arreglo(18))
                msr = IIf(arreglo(2) = 0 And (CInt(Aux_Cod_sitr) <> 5 And CInt(Aux_Cod_sitr) <> 6), 0, arreglo(2))
                
                'FGZ - 09/08/2004
                'If arreglo(19) = 0 Then arreglo(19) = 0.01
                'FGZ - 08/11/2004 - deja de tener vigencia
                
                imp_ss = IIf(arreglo(19) = 0 And CInt(Aux_Cod_sitr) <> 5 And CInt(Aux_Cod_sitr) <> 6, 0, arreglo(19))
                Resto = imp_ss         'Para administrar el tope en la suma de todo contra el imponible 1
                imp_oso = IIf(arreglo(20) = 0 And CInt(Aux_Cod_sitr) <> 5 And CInt(Aux_Cod_sitr) <> 6, 0, arreglo(20))
                imp_ss_con = IIf(arreglo(31) = 0 And CInt(Aux_Cod_sitr) <> 5 And CInt(Aux_Cod_sitr) <> 6, 0, arreglo(31))
                
                Aux_Rem_Total = Format(bruto, "#####0.00")
                Aux_Msr = Format(msr, "#####0.00")
                AUX_Imp_SS = Format(imp_ss, "#####0.00")
                Aux_imp_ss_con = Format(imp_ss_con, "#####0.00")
                AUX_Imp_OS = Format(imp_oso, "#####0.00")
                
                'adicional_os = 0      '(imp-oso * ( 8 / 100))  + (- arreglo[15])      8 % fijo
                Aux_Adi_OS = Format(arreglo(78), "#####0.00")
                Aux_Asig_Fliar = Format(arreglo(3), "#####00.00")
                Aux_Aporte_Voluntario = Format(-arreglo(11), "#####0.00")
                Aux_Exc_SS = Format(arreglo(12), "#####0.00")
                Aux_Exc_OS = Format(arreglo(13), "#####0.00")
                
                
                Aux_ConcNoRem = Format(arreglo(80), "#####0.00")
                Aux_Maternidad = Format(arreglo(81), "#####0.00")
                Aux_RectifDeRem = Format(arreglo(82), "#####0.00")
                
                'Version 33
                '07/05/2010 - Martin - Tomar de restricciones
                'If Not Aplica Then
                    Aux_remImpo9 = Format(arreglo(83), "#####0.00")
                'Else
                '    Aux_remImpo9 = Valor
                'End If
                Aux_tareaDiff = Format(arreglo(84), "#####0.00")
                Aux_horTrabMes = Format(arreglo(85), "#####0.00")
                
                
                rebaja_promo = 0
                Aux_RebajaPromovida = IIf(tcreduccion <> 100, Format(tcreduccion, "##0"), Format(0, "##0"))
                Aux_Cant_Hijos = Format(v_canthijos, "#0")
                Aux_Conyuges = Format(v_conyuge, "0")
                Aux_Adherentes = Format(Abs(arreglo(9)), "#0")
                
                ' Version 19
                Flog.writeline "Version 19"
                If IsNull(arreglo(14)) Or arreglo(14) = 0 Then
                    Aux_Porc_Adi = Format(0, "#0")
                Else
                    Aux_Porc_Adi = Format(arreglo(14), "#0")
                End If
           
                'AGREGADO POR VERSION 14
                Flog.writeline "AGREGADO POR VERSION 14"
                Aux_imprem3 = Format(imp_ss_con, "#####0.00")
                Aux_imprem4 = Format(imp_oso, "#####0.00")
                'Aux_correspred = "T"  'SIEMPRE VERDADERO
                Aux_correspred = " "
                Aux_caprecomlrt = Format(0, "##0.00") 'SIEMPRE 0
                
                Flog.writeline "AGREGADO POR VERSION 17"
                'AGREGADO POR VERSION 17
                Aux_Apo_Adi_OS = Format(arreglo(32), "#####0.00")
                
                'Version 18 r2 - Marzo 2002 modif por Version 19
                Flog.writeline "Version 18 r2 - Marzo 2002 modif por Version 19"
                'Aux_Adi_OS = Format(0, "#####0.00")
    
                'Version 19 - Agosto 2002
                Flog.writeline "Version 19 - Agosto 2002"
                Aux_Cant_Hijos = Format(v_canthijos, "#0")
                Aux_Imprem5 = Format(arreglo(33), "#####0.00")
                
                ' Incializo defaults x si no esta configurado el complemento.
                Flog.writeline "Incializo defaults x si no esta configurado el complemento."
                '07/09/2009 - Martin Ferraro - Se inicializo en 0 los dias trabajados
                'Aux_DiasTrab = "       30"
                Aux_DiasTrab = "        0"
                Aux_Sue_Adic = arreglo(34)
                
                ' Adrián.  Se inicializan las varibles por si llega a ser una liq de SAC o LAR
                Flog.writeline "Adrián.  Se inicializan las varibles por si llega a ser una liq de SAC o LAR"
                Aux_sac = arreglo(35)
                Aux_lar = arreglo(38)
                
                
                ' COMPLEMENTO : A partir de DIC. 2002 se usan estos campos que ya estan creados
                ' USO
    '            ' Topeo el SAC a manopla
    '            If arreglo(35) > 2400 Then
    '                arreglo(35) = 2400
    '            End If
    '
    '            resto = resto - arreglo(35)  'SAC
    '            resto = resto - arreglo(36)  'Hrs. Extras
    '            resto = resto - arreglo(37)  'Zona Desfavorable
    '
    '            ' Esto es por si solo se descuenta vacaciones - MAB - 06/02/03
    '            resto = resto - IIf(arreglo(38) < 0, 0, arreglo(38)) 'Vacaciones
    '
    '
    '            ' Lo que queda se lo mando al sueldo
    '            arreglo(34) = Max(resto, 0)
    '            arreglo(38) = Max(arreglo(38), 0) 'Esto es por si solo se descuenta vacaciones - MAB - 06/02/03
    
                ' Si da negativo ajusto las vacaciones , en lugar de ‚sto exceptuar los acumuladores del contrato con problemas
                If Resto < 0 And arreglo(38) > 0 Then
                    If arreglo(38) > Abs(Resto) Then 'Para que no de menor a cero - MAB - 06/02/2003
                        arreglo(38) = arreglo(38) + Resto
                    Else
                        arreglo(38) = 0
                    End If
                End If
    
                '03/08/2005 - MB en Deloitte Personal
                'comentarie la parte donde asignaba 0 a sueldo + adicionales si si este era 0.01 o 1
'                ' Para Autonomos Maxi 20/01/03
'                ' Adrián.  Se toco la condicion para que tenga en cuenta si es una liquidacion por SAC o por LAR antes de clarear el arreglo.
'                If (CSng(Aux_Sue_Adic) = 0 Or CSng(Aux_Sue_Adic) = 1 Or CSng(Aux_Sue_Adic) = 0.01) And _
'                   (CSng(Aux_sac) = 0 Or CSng(Aux_sac) = 1 Or CSng(Aux_sac) = 0.01) And _
'                   (CSng(Aux_lar) = 0 Or CSng(Aux_lar) = 1 Or CSng(Aux_lar) = 0.01) Then
'
'                    arreglo(35) = 0  'SAC
'                    arreglo(36) = 0 'Hrs. Extras
'                    arreglo(37) = 0 'Zona Desfavorable
'                    arreglo(38) = 0 'Vacaciones
'                    'arreglo(34) = 0.01 'Sueldo + adic para que de la suma igual al imponible
'                    'FGZ - 08/11/2004 - la linea anterioro deja de tener vigencia
'                    arreglo(34) = 0 'Sueldo + adic para que de la suma igual al imponible
'                End If
                '03/08/2005 - MB en Deloitte Personal
                
                'Martin Ferraro - 11/07/2006 - Se cambio por lo de abajo porque esto no topeaba
                'col19 debe ser igual o mayor a (col34+col42+col36+col37+col43) sino topeo
'                Aux_Sue_Adic = Format((arreglo(34) - arreglo(36) - arreglo(37)), "#####0.00")
                'Aux_sac = Format(arreglo(35), "#####0.00")
'                Aux_hrsextras = Format(arreglo(36), "#####0.00")
'                Aux_zonadesf = Format(arreglo(37), "#####0.00")
                'Aux_lar = Format(arreglo(38), "#####0.00")
                'Aux_sac = Format(arreglo(35), "#####0.00")

              
              
                'Inicializo el resto como el total del imp_ss (col 19)
                resto19 = imp_ss
                
                'Estos valores no pueden ser 0
                If arreglo(34) < 0 Then
                    arreglo(34) = 0
                End If
                If arreglo(36) < 0 Then
                    arreglo(36) = 0
                End If
                If arreglo(37) < 0 Then
                    arreglo(37) = 0
                End If
                If arreglo(42) < 0 Then
                    arreglo(42) = 0
                End If
                If arreglo(43) < 0 Then
                    arreglo(43) = 0
                End If
                If arreglo(35) < 0 Then
                    arreglo(35) = 0
                End If
                If arreglo(38) < 0 Then
                    arreglo(38) = 0
                End If
                
'21/07/2006 - Martin Ferraro - Se cambio el orden en que se realizan los topes
'                'Topeo Sueldo
'                If resto19 > arreglo(34) Then
'                    resto19 = resto19 - arreglo(34)
'                    sueldo = arreglo(34)
'                Else
'                    sueldo = resto19
'                    resto19 = 0
'                End If
'                'Topeo SAC
'                If resto19 > arreglo(35) Then
'                    resto19 = resto19 - arreglo(35)
'                    sac = arreglo(35)
'                Else
'                    sac = resto19
'                    resto19 = 0
'                End If
'                'Topeo lar
'                If resto19 > arreglo(38) Then
'                    resto19 = resto19 - arreglo(38)
'                    lar = arreglo(38)
'                Else
'                    lar = resto19
'                    resto19 = 0
'                End If
'                'topeo Horas Extras
'                If resto19 > arreglo(36) Then
'                    resto19 = resto19 - arreglo(36)
'                    hrsextras = arreglo(36)
'                Else
'                    hrsextras = resto19
'                    resto19 = 0
'                End If
'                'Topeo Adicionales
'                If resto19 > arreglo(42) Then
'                    resto19 = resto19 - arreglo(42)
'                    adicional = arreglo(42)
'                Else
'                    adicional = resto19
'                    resto19 = 0
'                End If
'                'topeo zona desf.
'                If resto19 > arreglo(37) Then
'                    resto19 = resto19 - arreglo(37)
'                    zonadesf = arreglo(37)
'                Else
'                    zonadesf = resto19
'                    resto19 = 0
'                End If
'                'Tope Premio
'                If resto19 > arreglo(43) Then
'                    resto19 = resto19 - arreglo(43)
'                    premios = arreglo(43)
'                Else
'                    premios = resto19
'                    resto19 = 0
'                End If

'               ------------------------------------------------------------------------------------------------------------
'               ------------------------------------------------------------------------------------------------------------
'               SICOSS - 06/04/2009 - Martin Ferraro - Se quito el topeo----------------------------------------------------
'               ------------------------------------------------------------------------------------------------------------
'               ------------------------------------------------------------------------------------------------------------
'                Flog.writeline " Se inicia el tope de columnas Vs Imponible"
'                Flog.writeline "--------------------------------------------"
'                Flog.writeline "--------------------------------------------"
'                Flog.writeline " Imponible = Resto = " & resto19
'                Flog.writeline " Topea contra SAC. SAC = " & arreglo(35)
'
'                'Topeo SAC
'                If resto19 > arreglo(35) Then
'                    resto19 = resto19 - arreglo(35)
'                    sac = arreglo(35)
'                Else
'                    sac = resto19
'                    resto19 = 0
'                End If
'                Flog.writeline " Resultado SAC = " & sac
'                Flog.writeline " Resto = " & resto19
'                Flog.writeline "--------------------------------------------"
'                Flog.writeline " Topea contra LAR. LAR = " & arreglo(38)
'
'                'Topeo lar
'                If resto19 > arreglo(38) Then
'                    resto19 = resto19 - arreglo(38)
'                    lar = arreglo(38)
'                Else
'                    lar = resto19
'                    resto19 = 0
'                End If
'                Flog.writeline " Resultado LAR = " & lar
'                Flog.writeline " Resto = " & resto19
'                Flog.writeline "--------------------------------------------"
'                Flog.writeline " Topea contra Sueldo. Sueldo = " & arreglo(34)
'
'                'Topeo Sueldo
'                If resto19 > arreglo(34) Then
'                    resto19 = resto19 - arreglo(34)
'                    sueldo = arreglo(34)
'                Else
'                    sueldo = resto19
'                    resto19 = 0
'                End If
'                Flog.writeline " Resultado Sueldo = " & sueldo
'                Flog.writeline " Resto = " & resto19
'                Flog.writeline "--------------------------------------------"
'                Flog.writeline " Topea contra Horas Extras. Horas Extras = " & arreglo(36)
'
'                'topeo Horas Extras
'                If resto19 > arreglo(36) Then
'                    resto19 = resto19 - arreglo(36)
'                    hrsextras = arreglo(36)
'                Else
'                    hrsextras = resto19
'                    resto19 = 0
'                End If
'                Flog.writeline " Resultado Horas Extras = " & hrsextras
'                Flog.writeline " Resto = " & resto19
'                Flog.writeline "--------------------------------------------"
'                Flog.writeline " Topea contra Premio. Premio = " & arreglo(43)
'
'                'Tope Premio
'                If resto19 > arreglo(43) Then
'                    resto19 = resto19 - arreglo(43)
'                    premios = arreglo(43)
'                Else
'                    premios = resto19
'                    resto19 = 0
'                End If
'                Flog.writeline " Resultado Premio = " & premios
'                Flog.writeline " Resto = " & resto19
'                Flog.writeline "--------------------------------------------"
'                Flog.writeline " Topea contra Adicionales. Adicionales = " & arreglo(42)
'
'                'Topeo Adicionales
'                If resto19 > arreglo(42) Then
'                    resto19 = resto19 - arreglo(42)
'                    adicional = arreglo(42)
'                Else
'                    adicional = resto19
'                    resto19 = 0
'                End If
'                Flog.writeline " Resultado Adicionales = " & adicional
'                Flog.writeline " Resto = " & resto19
'                Flog.writeline "--------------------------------------------"
'                Flog.writeline " Topea contra Zona Desfav. Zona Desfav = " & arreglo(37)
'
'                'topeo zona desf.
'                If resto19 > arreglo(37) Then
'                    resto19 = resto19 - arreglo(37)
'                    zonadesf = arreglo(37)
'                Else
'                    zonadesf = resto19
'                    resto19 = 0
'                End If
'                Flog.writeline " Resultado Zona Desfav = " & zonadesf
'               ------------------------------------------------------------------------------------------------------------
'               ------------------------------------------------------------------------------------------------------------
'               FIN - SICOSS - 06/04/2009 - Martin Ferraro - Se quito el topeo----------------------------------------------
'               ------------------------------------------------------------------------------------------------------------
'               ------------------------------------------------------------------------------------------------------------
                
                
                
'               ------------------------------------------------------------------------------------------------------------
'               FIN - SICOSS - 06/04/2009 - Martin Ferraro - Se quito el topeo----------------------------------------------
'               ------------------------------------------------------------------------------------------------------------
                'SAC
                sac = arreglo(35)
                Flog.writeline " Resultado SAC = " & sac
                
                'lar
                lar = arreglo(38)
                Flog.writeline " Resultado LAR = " & lar
                
                'Sueldo
                sueldo = arreglo(34)
                Flog.writeline " Resultado Sueldo = " & sueldo
                
                'Horas Extras
                hrsextras = arreglo(36)
                Flog.writeline " Resultado Horas Extras = " & hrsextras
                
                'Premio
                premios = arreglo(43)
                Flog.writeline " Resultado Premio = " & premios
                
                'Adicionales
                adicional = arreglo(42)
                Flog.writeline " Resultado Adicionales = " & adicional
                
                'zona desf.
                zonadesf = arreglo(37)
                Flog.writeline " Resultado Zona Desfav = " & zonadesf
'               ------------------------------------------------------------------------------------------------------------
'               ------------------------------------------------------------------------------------------------------------
                
                
                Flog.writeline " Fin de Topes de Columnas"
                'Martin Ferraro - 12/07/2006
                Aux_adicional = Format(adicional, "#####0.00")
                Aux_premios = Format(premios, "#####0.00")
                Aux_Sue_Adic = Format(sueldo, "#####0.00")
                Aux_hrsextras = Format(hrsextras, "#####0.00")
                Aux_zonadesf = Format(zonadesf, "#####0.00")
                Aux_lar = Format(lar, "#####0.00")
                Aux_sac = Format(sac, "#####0.00")
                Aux_remdec788 = Format(arreglo(44), "#####0.00")
                Aux_remimpo7 = Format(arreglo(45), "#####0.00")
                
                
                If arreglo(39) <> 0 Then
                    'Para que nunca de negativo
                    Aux_DiasTrab = Format((arreglo(39) - arreglo(49)), "#####0.00")
                Else
                    'SUMA Y RESTA LOS DIAS TRABAJADOS DEL MES
                    Aux_DiasTrab = Aux_DiasTrab
                End If
                'FGZ - 28/12/2004
                'Valido que los dias trabajados no superen los 30 dias y no sean menores que cero
                'If Aux_DiasTrab > 30 Then
                '    Flog.writeline "Los dias trabajados superan los 30." & Aux_DiasTrab & ". Se topean en 30 "
                '    Aux_DiasTrab = Format(30, "#####0.00")
                'End If
                If Aux_DiasTrab < 0 Then
                    Flog.writeline "Los dias trabajados son negativos" & Aux_DiasTrab & ". Revisar el acumulador correspondiente, columna 39 del confrep "
                    Aux_DiasTrab = Format(0, "#####0.00")
                End If
                

                '04/09/2009 - Martin Ferraro - Redondeo de la cant de horas si es > a 1. Sino 1
                If (arreglo(50) > 0) Then
                    If (arreglo(50) <= 1) Then
                        Aux_canthsext = "1"
                    Else
                        Aux_canthsext = Format(Round(arreglo(50), 0), "##0")
                    End If
                Else
                    Aux_canthsext = "0"
                End If

                
                
                ' ----------------------------------------------------------------
            
                ' ----------------------------------------------------------------
                'FGZ - 01/09/2005
                'Restricciones
                Aplica = False
                For I = 0 To LS_Res
                    StrSql = " SELECT * FROM his_estructura "
                    StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
                    StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
                    StrSql = StrSql & " his_estructura.estrnro = " & Restricciones(I).Estrnro & " AND "
                    StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_fecha) & ") AND "
                    StrSql = StrSql & " ((" & ConvFecha(Aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                    StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
                    OpenRecordset StrSql, rs_Estructura
                    If Not rs_Estructura.EOF Then
                        Aplica = True
                        Valor = Restricciones(I).Valor
                        Exit For
                    End If
                Next I
                If Aplica Then
                    Aux_Rem_Total = Valor
                    AUX_Imp_SS = Valor
                    Aux_imp_ss_con = Valor
                    Aux_imprem3 = Valor
                    Aux_imprem4 = Valor
                    AUX_Imp_SS = Valor
                    Aux_Sue_Adic = Valor
                    '15/02/2008 - Martin Ferraro - Se agrego la restriccion a Remuneracion 5
                    Aux_Imprem5 = Valor
                    'Aux_sac = Valor
                    'Martin Ferraro - 12/07/2006
                    Aux_sac = Format(0, "#####0.00")
                    Aux_adicional = Format(0, "#####0.00")
                    Aux_premios = Format(0, "#####0.00")
                    Aux_hrsextras = Format(0, "#####0.00")
                    Aux_zonadesf = Format(0, "#####0.00")
                    Aux_lar = Format(0, "#####0.00")
                    
                    '07/05/2010 - Martin - Tomar de restricciones
                    'Aux_remdec788 = Format(0, "#####0.00")
                    'Aux_remimpo7 = Format(0, "#####0.00")
                    Aux_remdec788 = Valor
                    Aux_remimpo7 = Valor
                    
                    Aux_remImpo9 = Valor
                    
                
                End If
                ' ----------------------------------------------------------------
            
                ' ------------------------------------------------------------------------------------
                'FGZ - 23/08/2005
                'Cargo los datos para el detallado
                'Busco la cantidad de hijos discapacitados
                Det_CantHDisc = 0
                Det_Prenatal = 0
                Det_Escolar = 0
                
                Edad = 18 'Edad tope para considerar Escolar
                Estudia = True  'Si se contempla si estudia o no
                Nivel_Estudio = 2 'nivel de estudio considerado
                
                'Parentesco = 2 'Hijos
                StrSql = "SELECT parenro,terfecnac, faminc, familiar.ternro FROM familiar INNER JOIN tercero ON tercero.ternro = familiar.ternro"
                StrSql = StrSql & " WHERE (familiar.empleado =" & rs_Empleados!Ternro
                'StrSql = StrSql & " AND familiar.parenro = " & parentesco
                StrSql = StrSql & " AND familiar.famest = -1"
                StrSql = StrSql & " AND familiar.famsalario = -1)"
                StrSql = StrSql & " AND (familiar.famfecvto >=" & ConvFecha(Aux_fecha) & " OR familiar.famfecvto is null)"
                StrSql = StrSql & " Order by tercero.terfecnac DESC"
                If rs_Familiar.State = adStateOpen Then rs_Familiar.Close
                OpenRecordset StrSql, rs_Familiar
                Do While Not rs_Familiar.EOF
                    If rs_Familiar!parenro = 2 Then 'Hijo
                        edad_F = Calcular_Edad(rs_Familiar!terfecnac)
                        If CBool(rs_Familiar!faminc) Then
                            Det_CantHDisc = Det_CantHDisc + 1
                        End If
                        
                        'buscar el nivel de estudio
                        Fam_Estudia = False
                        StrSql = "SELECT * FROM estudio_actual WHERE ternro = " & rs_Familiar!Ternro
                        If rs_Estudio_Actual.State = adStateOpen Then rs_Estudio_Actual.Close
                        OpenRecordset StrSql, rs_Estudio_Actual
                        If Not rs_Estudio_Actual.EOF Then
                            If Not EsNulo(rs_Estudio_Actual!nivnro) Then
                                StrSql = "SELECT * FROM nivest WHERE nivnro =" & rs_Estudio_Actual!nivnro
                                If rs_Nivest.State = adStateOpen Then rs_Nivest.Close
                                OpenRecordset StrSql, rs_Nivest
                                If Not rs_Nivest.EOF Then
                                    Fam_Niv_Est = rs_Nivest!nivnro
                                    Fam_Estudia = True
                                Else
                                    Fam_Niv_Est = 0
                                End If
                            End If
                        End If
                        
                        If edad_F <= Edad Then
                            If Fam_Estudia Then
                                If Fam_Niv_Est <> 0 Then
                                    If Fam_Niv_Est <= Nivel_Estudio Then
                                        Det_Escolar = Det_Escolar + 1
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If rs_Familiar!parenro = 8 Then 'Prenatal
                            Det_Prenatal = Det_Prenatal + 1
                        End If
                    End If
                    
                    rs_Familiar.MoveNext
                Loop
                
                'Demas valores
                Det_PlusZonaDes = Format(arreglo(60), "#####0.00")
                Det_Ret_Jubilacion = Format(arreglo(61), "#####0.00")
                Det_Ret_Ley = Format(arreglo(62), "#####0.00")
                Det_Ret_OS = Format(arreglo(63), "#####0.00")
                Det_Ret_Ansal = Format(arreglo(64), "#####0.00")
                Det_Total_Ret_OS = Format(arreglo(65), "#####0.00")
                Det_Con_Jubilacion = Format(arreglo(66), "#####0.00")
                Det_Con_Ley = Format(arreglo(67), "#####0.00")
                Det_Con_FNE = Format(arreglo(68), "#####0.00")
                Det_Con_AsigFliares = Format(arreglo(69), "#####0.00")
                Det_Con_OS = Format(arreglo(70), "#####0.00")
                Det_Con_Ansal = Format(arreglo(71), "#####0.00")
                Det_Con_Total_OS = Format(arreglo(72), "#####0.00")
                
                '07/05/2010 - Martin - Tomar de restricciones
                If Not Aplica Then
                    Aux_remimpo6 = Format(arreglo(79), "#####0.00")
                Else
                    Aux_remimpo6 = Valor
                End If
                
'                Aux_adicional = Format(adicional, "#####0.00")
'                Aux_premios = Format(premios, "#####0.00")
'                Aux_Sue_Adic = Format(sueldo, "#####0.00")
'                Aux_hrsextras = Format(hrsextras, "#####0.00")
'                Aux_zonadesf = Format(zonadesf, "#####0.00")
'                Aux_lar = Format(lar, "#####0.00")
'                Aux_sac = Format(sac, "#####0.00")
'
'                Aux_remdec788 = Format(arreglo(44), "#####0.00")
'                Aux_remimpo7 = Format(arreglo(45), "#####0.00")

                '22/07/2010 - Martin Ferraro - Seguro de vida colectivo - V34
'                Aux_SegVida = "0"
'                StrSql = " SELECT * FROM his_estructura "
'                StrSql = StrSql & " WHERE his_estructura.ternro = " & rs_Empleados!Ternro & " AND "
'                StrSql = StrSql & " his_estructura.tenro = " & arreglo(86)
'                StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND "
'                StrSql = StrSql & " ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
'                OpenRecordset StrSql, rs_Estructura
'                If Not rs_Estructura.EOF Then
'                    Aux_SegVida = "1"
'                End If
'----------------------------------------------------------- Rafa -----------------------------------------------------------
                StrSql = "SELECT nrocod FROM tipoestructura "
                StrSql = StrSql & "INNER JOIN estructura ON estructura.tenro=tipoestructura.tenro "
                StrSql = StrSql & "INNER JOIN estr_cod ON estr_cod.estrnro=estructura.estrnro "
                StrSql = StrSql & "INNER JOIN tipocod ON tipocod.tcodnro = estr_cod.tcodnro "
                StrSql = StrSql & "INNER JOIN his_estructura ON his_estructura.estrnro=estructura.estrnro "
                StrSql = StrSql & "INNER JOIN empleado ON empleado.ternro = his_estructura.ternro "
                StrSql = StrSql & " AND his_estructura.ternro = " & rs_Empleados!Ternro
                StrSql = StrSql & " WHERE tipocod.tcodnro = 1 And estructura.Tenro = " & arreglo(86)
                'licho - 16/12/2014 - se agrego la condicion al historico * fecha
                StrSql = StrSql & " AND ((" & ConvFecha(Aux_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
                StrSql = StrSql & " ORDER BY his_estructura.htetdesde"
                OpenRecordset StrSql, rs_Estructura
                Flog.writeline "Tipo de Codigo " & StrSql
                If Not rs_Estructura.EOF Then
                    Aux_SegVida = rs_Estructura!nrocod
                Else
                    Aux_SegVida = ""
                    Flog.writeline "No tiene Tipo Estructura Asociado"
                End If
                ' ------------------------------------------------------------------------------------
            
            
                'inserto en repsijp
                StrSql = "INSERT INTO repsijp (pronroemp,empleg,cuil,cod_obra_social,cant_hijos"
                StrSql = StrSql & ",conyuges,adherentes,rebajapromovida,zona,actividad,reduccion"
                StrSql = StrSql & ",rem_total,imp_ss,aporte_voluntario,exc_ss,asig_fliar,exc_os"
                StrSql = StrSql & ",imp_os,adi_os,msr,porc_adi,apenom,cod_sitr,cod_cond,localidad"
                'StrSql = StrSql & ",cod_cont,con_fne,imp_ss_con,correspred,caprecomlrt"
                'FGZ - 23/08/2005
                'le saqué el con_fne porque se carga con los nuevos valores
                ' no se para que estaba ni que cargaba
                StrSql = StrSql & ",cod_cont,imp_ss_con,correspred,caprecomlrt"
                StrSql = StrSql & ",codsiniestro,imprem3,imprem4,tipempnro,apo_adi_os"
                StrSql = StrSql & ",cod_sitr1,diainisr1,cod_sitr2,diainisr2,cod_sitr3,diainisr3"
                StrSql = StrSql & ",sue_adic,sac,hrsextras,zonadesf,lar,diastrab,imprem5,xconv"
                StrSql = StrSql & ",pliqnro,lista_pronro,empresa,manual"
                
                StrSql = StrSql & ",canthdisc,prenatal,escolar,pluszonades"
                StrSql = StrSql & ",ret_jubilacion,ret_ley,ret_os,ret_ansal" ',ret_os"
                StrSql = StrSql & ",con_jubilacion,con_ley,con_fne,con_asigfliares"
                StrSql = StrSql & ",con_os,con_ansal,secrep"   ',con_os"
                
                'Version SIJP 26 y 27
                StrSql = StrSql & ",adicional,premios,remdec788,remimpo7,canthsext"
                
                StrSql = StrSql & ",remimpo6"
                
                'Version SICOSS
                StrSql = StrSql & ",ConcNoRem"
                StrSql = StrSql & ",Maternidad"
                StrSql = StrSql & ",RectifDeRem"
                
                'Version SICOSS
                StrSql = StrSql & ",remImpo9"
                StrSql = StrSql & ",tareaDiff"
                StrSql = StrSql & ",horTrabMes"
                
                'Version SICOSS 34
                StrSql = StrSql & ",SegVida"
                
                
                StrSql = StrSql & ") VALUES ("
                
                StrSql = StrSql & rs_Empleados!pronro & ","
                StrSql = StrSql & rs_Empleados!empleg & ","
                StrSql = StrSql & "'" & Aux_CUIL & "',"
                StrSql = StrSql & "'" & Aux_Cod_Obra_Social & "',"
                StrSql = StrSql & "'" & Aux_Cant_Hijos & "',"
                StrSql = StrSql & "'" & Aux_Conyuges & "',"
                StrSql = StrSql & "'" & Aux_Adherentes & "',"
                StrSql = StrSql & "'" & Aux_RebajaPromovida & "',"
                StrSql = StrSql & "'" & Aux_Zona & "',"
                StrSql = StrSql & "'" & Aux_Actividad & "',"
                StrSql = StrSql & "'" & Aux_Reduccion & "',"
                StrSql = StrSql & "'" & Aux_Rem_Total & "',"
                StrSql = StrSql & "'" & AUX_Imp_SS & "',"
                StrSql = StrSql & "'" & Aux_Aporte_Voluntario & "',"
                StrSql = StrSql & "'" & Aux_Exc_SS & "',"
                StrSql = StrSql & "'" & Aux_Asig_Fliar & "',"
                StrSql = StrSql & "'" & Aux_Exc_OS & "',"
                StrSql = StrSql & "'" & AUX_Imp_OS & "',"
                StrSql = StrSql & "'" & Aux_Adi_OS & "',"
                StrSql = StrSql & "'" & Aux_Msr & "',"
                
                StrSql = StrSql & "'" & Aux_Porc_Adi & "',"
                StrSql = StrSql & "'" & Aux_ApeNom & "',"
                StrSql = StrSql & "'" & Aux_Cod_sitr & "',"
                StrSql = StrSql & "'" & Aux_Cod_Cond & "',"
                StrSql = StrSql & "'" & Mid(Aux_Localidad, 1, 400) & "',"
                
                StrSql = StrSql & "'" & Aux_Cod_Cont & "',"
                'FGZ - 23/08/2005
                'le saqué el con_fne porque se carga con los nuevos valores
                ' no se para que estaba ni que cargaba
                'StrSql = StrSql & "'" & Aux_con_FNE & "',"
                StrSql = StrSql & "'" & Aux_imp_ss_con & "',"
                StrSql = StrSql & "'" & Aux_correspred & "',"
                StrSql = StrSql & "'" & Aux_caprecomlrt & "',"
                StrSql = StrSql & "'" & Aux_codsiniestro & "',"
                StrSql = StrSql & "'" & Aux_imprem3 & "',"
                StrSql = StrSql & "'" & Aux_imprem4 & "',"
                StrSql = StrSql & "'" & Aux_TipEmpNro & "',"
                StrSql = StrSql & "'" & Aux_Apo_Adi_OS & "',"
                StrSql = StrSql & "'" & Aux_Cod_sitr1 & "',"
                StrSql = StrSql & "'" & Aux_diainisr1 & "',"
                StrSql = StrSql & "'" & Aux_Cod_sitr2 & "',"
                StrSql = StrSql & "'" & Aux_diainisr2 & "',"
                StrSql = StrSql & "'" & Aux_Cod_sitr3 & "',"
                StrSql = StrSql & "'" & Aux_diainisr3 & "',"
                StrSql = StrSql & "'" & Aux_Sue_Adic & "',"
                StrSql = StrSql & "'" & Aux_sac & "',"
                StrSql = StrSql & "'" & Aux_hrsextras & "',"
                StrSql = StrSql & "'" & Aux_zonadesf & "',"
                StrSql = StrSql & "'" & Aux_lar & "',"
                StrSql = StrSql & "'" & Aux_DiasTrab & "',"
                StrSql = StrSql & "'" & Aux_Imprem5 & "',"
                StrSql = StrSql & "'" & Aux_XConv & "',"
                StrSql = StrSql & Nroliq & ","
                StrSql = StrSql & "'" & Lista_Pro_F & "',"
                StrSql = StrSql & Empresa & ","
                StrSql = StrSql & "0,"
                
                StrSql = StrSql & Det_CantHDisc & ","
                StrSql = StrSql & Det_Prenatal & ","
                StrSql = StrSql & Det_Escolar & ","
                StrSql = StrSql & Det_PlusZonaDes & ","
                
                StrSql = StrSql & Det_Ret_Jubilacion & ","
                StrSql = StrSql & Det_Ret_Ley & ","
                StrSql = StrSql & Det_Ret_OS & ","
                StrSql = StrSql & Det_Ret_Ansal & ","
                'StrSql = StrSql & Det_Total_Ret_OS & ","
                
                StrSql = StrSql & Det_Con_Jubilacion & ","
                StrSql = StrSql & Det_Con_Ley & ","
                StrSql = StrSql & Det_Con_FNE & ","
                StrSql = StrSql & Det_Con_AsigFliares & ","
                
                StrSql = StrSql & Det_Con_OS & ","
                StrSql = StrSql & Det_Con_Ansal & ","
                StrSql = StrSql & Aux_con_FNE & ","
                'StrSql = StrSql & Det_Con_Total_OS
                
                'Version SIJP 26
                StrSql = StrSql & "'" & Aux_adicional & "',"
                StrSql = StrSql & "'" & Aux_premios & "',"
                StrSql = StrSql & "'" & Aux_remdec788 & "',"
                StrSql = StrSql & "'" & Aux_remimpo7 & "',"
                StrSql = StrSql & "'" & Aux_canthsext & "',"
                StrSql = StrSql & "'" & Aux_remimpo6 & "',"
                
                'Version SICOSS
                StrSql = StrSql & "'" & Aux_ConcNoRem & "',"
                StrSql = StrSql & "'" & Aux_Maternidad & "',"
                StrSql = StrSql & "'" & Aux_RectifDeRem & "',"
                
                'Version 33 SICOSS
                StrSql = StrSql & "'" & Aux_remImpo9 & "',"
                StrSql = StrSql & "'" & Aux_tareaDiff & "',"
                StrSql = StrSql & "'" & Aux_horTrabMes & "',"
                
                'Version 34 SICOSS
                StrSql = StrSql & "'" & Aux_SegVida & "'"
                
                StrSql = StrSql & ")"
                Flog.writeline "Insertando : " & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
                                
                'restauro la empresa original
                Empresa = Empresa_Original
                
                ' ----------------------------------------------------------------
            
            End If 'If EsElUltimoEmpleado(rs_Empleados, UltimoEmpleado) Then
            
            Progreso = Progreso + IncPorc
            TiempoAcumulado = GetTickCount
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                     "' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            
            
            rs_Empleados.MoveNext
        Loop
    End If

    rs_Periodo.MoveNext
Loop

'Fin de la transaccion
10 If Not HuboError Then
    MyCommitTrans
Else
    MyRollbackTrans
End If


If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
If rs_Acu_liq.State = adStateOpen Then rs_Acu_liq.Close
If rs_Proceso.State = adStateOpen Then rs_Proceso.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
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

Set rs_Empleados = Nothing
Set rs_Acu_liq = Nothing
Set rs_Proceso = Nothing
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

Exit Sub

CE:
    Flog.writeline "=================================================================="
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "Ultimo SQl Ejecutado: " & StrSql
    Flog.writeline "=================================================================="
    MyRollbackTrans
    MyBeginTrans
    Progreso = Progreso + IncPorc
    TiempoAcumulado = GetTickCount
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
             ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
             "' WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
    
    HuboError = True
    Flog.writeline " Error: " & Err.Description

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


'EAM- Obtiene la situación de revista configurado en el confrep
Public Function Buscar_SituacionRevistaConfig(ByVal Ternro As Long, ByVal Fdesde As Date, ByVal Fhasta As Date) As Integer
 Dim rs_ConfrepAux As New ADODB.Recordset
 Dim rs_aux As New ADODB.Recordset
 Dim codigoExt As Long
    codigoExt = 0
    StrSql = "SELECT confval,confval2 FROM confrep WHERE repnro = 8 AND upper(conftipo) = 'REV'"
    OpenRecordset StrSql, rs_ConfrepAux

    Do While Not rs_ConfrepAux.EOF
        StrSql = "SELECT * FROM his_estructura WHERE ternro= " & Ternro & " AND his_estructura.estrnro = " & rs_ConfrepAux("confval") & _
                " AND (htethasta< " & ConvFecha(Fhasta) & " OR htethasta IS NULL) ORDER BY htethasta,htetdesde DESC"
        OpenRecordset StrSql, rs_aux
        
        If Not rs_aux.EOF Then
            'Buscar_SituacionRevistaConfig = rs_ConfrepAux("confval2")
            codigoExt = rs_ConfrepAux("confval2")
            Exit Do
        End If
        
        rs_ConfrepAux.MoveNext
    Loop
        
    Buscar_SituacionRevistaConfig = codigoExt
End Function

