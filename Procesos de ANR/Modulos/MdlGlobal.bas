Attribute VB_Name = "MdlGlobal"
    Option Explicit


Global Parametro As Double
Global conce As Long
Global Fecha_Inicio As Date
Global Fecha_Fin As Date
Global Empleado_Fecha_Inicio As Date
Global Empleado_Fecha_Fin As Date

Global guarda_nov As Boolean
Global NovedadesHist As Boolean
Global SoloLimpieza As Boolean

Global Valor As Double
Global Monto As Double
Global Retro As Date
Global Bien As Boolean
Global AFecha As Date

Global Valor_Ampo As Double
Global Valor_Ampo_Cont As Double

Global Cant_Diaria_Ampos_1 As Double
Global Cant_Diaria_Ampos_2 As Double
Global Cant_Diaria_Ampos_3 As Double
Global Cant_Diaria_Ampos_4 As Double
Global Cant_Diaria_Ampos_5 As Double

Global Cant_Ampo_Proporcionar_1 As Double
Global Cant_Ampo_Proporcionar_2 As Double
Global Cant_Ampo_Proporcionar_3 As Double
Global Cant_Ampo_Proporcionar_4 As Double
Global Cant_Ampo_Proporcionar_5 As Double

Global Ampo_Proporciona_1 As Boolean
Global Ampo_Proporciona_2 As Boolean
Global Ampo_Proporciona_3 As Boolean
Global Ampo_Proporciona_4 As Boolean
Global Ampo_Proporciona_5 As Boolean

Global Sumo_Cant_Ampo_Prop_1 As Boolean
Global Sumo_Cant_Ampo_Prop_2 As Boolean
Global Sumo_Cant_Ampo_Prop_3 As Boolean
Global Sumo_Cant_Ampo_Prop_4 As Boolean
Global Sumo_Cant_Ampo_Prop_5 As Boolean

Global Ampo_Max_1 As Double
Global Ampo_Max_2 As Double
Global Ampo_Max_3 As Double
Global Ampo_Max_4 As Double
Global Ampo_Max_5 As Double

Global Ampo_Min_1 As Double
Global Ampo_Min_2 As Double
Global Ampo_Min_3 As Double
Global Ampo_Min_4 As Double
Global Ampo_Min_5 As Double

'headcom.i
Global exito As Boolean

' varias
Global StrSql As String
Global StrSqlDatos As String
Global fs
Global Flog
Global FlogE
Global FlogP
Global rs As New ADODB.Recordset

Global Retroactivo As Boolean
Global pliqdesde As Long
Global pliqhasta As Long
Global Concepto_Retroactivo As Long
Global concepto_pliqdesde As Long
Global concepto_pliqhasta As Long

Global NroEmp As Long      ' empresa.empnro
Global NroEmple As Long    ' tercero.ternro
Global NroGrupo As Long    ' grpo.grunro
Global NroConce As Long    ' concepto.concnro
Global NroTpa As Long      ' tipopar.tpanro
Global NroCab As Long      ' cabliq.cliqnro
Global NroProg As Long     ' programa.prognro
Global tipoBus As Long     ' programa.tprognro
Global NroProc As String      ' rep06.pronro lista de procesos seleccionados separados por "-"
Global ListaNroProc As String ' lista de procesos seleccionados separados por ","
Global NroProcesoBatch As Long   'Nro de Proceso Batch generado
Global TipoProceso As Long       'Indica si el liquidador es Arg, Uru, Chile, etc.

' Registros Globales
Global buliq_proceso As New ADODB.Recordset
Global buliq_periodo As New ADODB.Recordset
Global buliq_impgralarg As New ADODB.Recordset
Global buliq_empleado As New ADODB.Recordset
Global buliq_tercero_emp As New ADODB.Recordset
Global buliq_cabliq As New ADODB.Recordset
Global rs_Buliq_Concepto As New ADODB.Recordset
Global rs_FunFormulas As New ADODB.Recordset
Global ErrorPosicion As Long
Global ErrorDescripcion As String
Global Texto As String
Global HACE_TRAZA As Boolean
Global FirmaActiva5 As Boolean
Global FirmaActiva15 As Boolean
Global FirmaActiva19 As Boolean
Global FirmaActiva20 As Boolean

'Variables de Progreso
Global CEmpleadosAProc As Long
Global CConceptosAProc As Long
Global IncPorc As Double
Global IncPorcEmpleado As Double
Global Progreso As Double

Global TiempoInicio As Long
Global TiempoFinal As Long
Global Fin As Long
Global Inicio As Long
Global TiempoAcumulado As Long

Global TiempoInicialProceso As Long
Global TiempoFinalProceso As Long
Global TiempoInicialEmpleado As Long
Global TiempoFinalEmpleado As Long
Global TiempoInicialConcepto As Long
Global TiempoFinalConcepto As Long
Global TiempoInicialParametro As Long
Global TiempoFinalParametro As Long
Global TiempoInicialBusqueda As Long
Global TiempoFinalBusqueda As Long
Global TiempoInicialFormula As Long
Global TiempoFinalFormula As Long

Global EmpleadoSinError As Boolean
Global HuboError As Boolean

Global Contador_updates_WF_impproarg As Long
Global Contador_updates_WF_impmesarg As Long
Global Contador_updates_acu_liq As Long
Global Contador_updates_acu_mes As Long

Global objconnProgreso As New ADODB.Connection
Global ContadorProgreso As Long
Global Cantidad_de_OpenRecordset As Long
Global Borrar_Estadisticas As Boolean
Global Legajo As Long
Global Etiqueta

'-----------------
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long


