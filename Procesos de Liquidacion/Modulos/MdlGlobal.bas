Attribute VB_Name = "MdlGlobal"
Option Explicit

Public Type tAcumulador
    acuNro As Variant
    acudesabr As Variant
    acusist As Variant
    acudesext As Variant
    acumes As Variant
    acutopea As Variant
    acudesborde As Variant
    acurecalculo As Variant
    acuimponible As Variant
    acuimpcont As Variant
    acusel1 As Variant
    acusel2 As Variant
    acusel3 As Variant
    acuppag As Variant
    acudepu As Variant
    acuhist As Variant
    acumanual As Variant
    acuimpri As Variant
    tacunro As Variant
    Empnro As Variant
    acuorden As Variant
    acuretro As Variant
    acunoneg As Variant
End Type

'Tabla For_tpa
Public Type TFor_tpa
    fornro As Variant
    tpanro As Variant
    ftentrada As Variant
    ftimprime As Variant
    ftorden As Variant
    ftobligatorio As Variant
    ftinicial As Variant
End Type

' Tabla Concepto
Public Type Tbuliq_concepto
    ConcNro  As Variant
    ConcCod As Variant
    concabr As Variant
    concorden As Variant
    tconnro As Variant
    concext As Variant
    concvalid As Variant
    concdesde As Variant
    conchasta As Variant
    concrepet As Variant
    concretro As Variant
    concniv As Variant
    fornro As Variant
    concimp As Variant
    codseguridad As Variant
    concusado As Variant
    concpuente As Variant
    Empnro As Variant
    Conccantdec As Variant
    Conctexto As Variant
    concautor As Variant
    concfecmodi As Variant
    Concajuste As Variant
End Type

Public Type TCft_Segun
    ConcNro  As Variant
    tpanro As Variant
    Nivel As Variant
    Origen As Variant
    Selecc As Variant
    fornro As Variant
    Entidad As Variant
End Type

Public Type TCge_Segun
    ConcNro As Variant
    Nivel As Variant
    Origen As Variant
    Entidad As Variant
End Type


Public Type TCon_for_tpa
    ConcNro As Variant
    tpanro As Variant
    depurable As Variant
    cftauto As Variant
    fornro As Variant
    Nivel As Variant
    Selecc As Variant
    Prognro As Variant
End Type

Public Type TPrograma
    Prognro As Variant
    Prognom As Variant
    Progdesc As Variant
    Tprognro As Variant
    Progarch As Variant
    Auxint1 As Variant
    Auxint2 As Variant
    Auxint3 As Variant
    Auxint4 As Variant
    Auxint5 As Variant
    Auxlog1 As Variant
    Auxlog2 As Variant
    Auxlog3 As Variant
    Auxlog4 As Variant
    Auxlog5 As Variant
    Auxchar1 As Variant
    Auxchar2 As Variant
    Auxchar3 As Variant
    Auxchar4 As Variant
    Auxchar5 As Variant
    Progarchest As Variant
    Progcache As Variant
    Progautor As Variant
    Progfecmodi As Variant
    Empnro As Variant
    Auxlog6 As Variant
    Auxlog7 As Variant
    Auxlog8 As Variant
    Auxlog9 As Variant
    Auxlog10 As Variant
    Auxlog11 As Variant
    Auxlog12 As Variant
    Auxchar As Variant
End Type

Public Type TEmpCabLiq
    'Ternro As Long
    cliqnro As Long
    Empleado As Long
End Type

Public Type tConcepto
    ConcCod As Variant
    ConcNro As Long
    tconnro As Long
    Concajuste As Boolean
    Conccantdec As Long
    concabr As String
    concretro As Boolean
    Conctexto As String
    fornro As Long
    Fortipo As Long
    Forexpresion As String
    Fordabr As String
    Forprog As String
    Seguir As Boolean
    NetoFijo As Double
End Type

Public Type Feriado
    Feriado  As Variant
End Type

Public Type TSanciones
    X As String
    Y As String
    Z As String
End Type

'Proporcion de contratos
Public Type TContrato
    Ternro As Variant
    proporcion As Variant
End Type


'Tabla temporal #wf_tpa
Public Type TWF_Tpa
    tipoparam As Variant
    ftorden As Variant
    Nombre As Variant
    Valor As Variant
    Fecha As Variant
    Calculado As Boolean
End Type


Public Type THisEstructura
    Tenro As Variant
    Estrnro As Variant
    htetdesde As Variant
    htethasta As Variant
End Type


Public Type TDistrib
    ID As Long
    Valor As Variant
End Type

Public Type TCImp_Alcance
    ConcNro As Variant
    tipo_dist As Variant
    ConcNro_depende As Variant
    Alcance As Variant
    Origen As Variant
End Type

Public Type T_Retenciones
    Retencion As Integer
End Type

Public Type T_Periodos
    PerMesDesde As Integer
    PerMesHasta As Integer
    PerMontoMensual As Integer
End Type

Public Type T_Detalles
    DetNombre As String
    DetValor As Integer
End Type

Public LimiteSuperior As Integer

Public Type T_RegPeriodo
    periodo_mesDesde As String
    periodo_meshasta As String
    periodo_montoMensual As String
    periodo_anio As String
End Type

Public Type T_RegDetalle
    detalle_nombre As String
    detalle_valor As String
End Type

Public Type T_RegingAp
    obrasoc As String
    segsocial As String
    sind As String
    ganbrut As String
    retgan As String
    retribNoHab As String
    Mes As Integer
    ajuste As String
End Type

Public Type T_RegGanLiq
    cuit As String
    Denominacion As String
    ingAp() As T_RegingAp
End Type

Public Type T_RegDeducciones
    Legajo As Integer
    tipo As Integer
    TipoDoc As String
    NroDoc As String
    Denominacion As String
    DescBasica As String
    DescAdicional As String
    MontoTotal As String
    Mes As String
    Mes_periodo As String 'MDF 28/07/2015----
    detallesdeduc() As T_RegDetalle
    periodosdeduc() As T_RegPeriodo
End Type

Public Type T_RegRetPerPago
    Legajo As Integer
    tipo As Integer
    TipoDoc As String
    NroDoc As String
    Denominacion As String
    DescBasica As String
    DescAdicional As String
    MontoTotal  As String
    detallesretpagos() As T_RegDetalle
    periodosretpagos() As T_RegPeriodo
End Type

Public Type T_RegCargaFamilia
    Legajo As Integer
    TipoDoc As String
    NroDoc As String
    Apellido As String
    Nombre As String
    FechaNac As String
    MesDesde As Integer
    MesHasta As Integer
    Parentesco As Integer
End Type

Public Type T_RegDatosAdicionales 'Carmen Quintero 20/09/2016---
    exencionGan2016SAC1 As String
End Type

Global Feriado_laborable As Boolean
Global Feriado_Por_Estructura As Boolean
Global Parametro As Double
Global conce As Long

' dbuf-liq.i

' Definicion de Variables globales
Global Fecha_Inicio As Date
Global Fecha_Fin As Date

'FGZ - 18/03/2004
Global Empleado_Fecha_Inicio As Date
Global Empleado_Fecha_Fin As Date

Global guarda_nov As Boolean
Global NovedadesHist As Boolean
Global SoloLimpieza As Boolean

' def-for.i
'Definici¢n de Par metros para los Pgmas. de F¢rmulas
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
Global fs 'As Scripting.FileSystemObject
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

'Global Monto_Proratear as double
'Global Monto_ya_prorratear As Boolean


Global NroEmp As Long      ' empresa.empnro
Global NroEmple As Long    ' tercero.ternro
Global NroGrupo As Long    ' grpo.grunro
Global NroConce As Long    ' concepto.concnro
Global NroTpa As Long      ' tipopar.tpanro
Global NroCab As Long      ' cabliq.cliqnro
Global NroProg As Long     ' programa.prognro
Global tipoBus As Long     ' programa.tprognro
'Global NroProc As Long      ' proceso.pronro
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
'' FGZ - 04/02/2004 -----------
'Global rs_NovGral As New ADODB.Recordset
'' FGZ - 04/02/2004 -----------


Global ErrorPosicion As Long
Global ErrorDescripcion As String

Global Texto As String

Global HACE_TRAZA As Boolean

Global FirmaActiva5 As Boolean
Global FirmaActiva15 As Boolean
Global FirmaActiva13 As Boolean
Global FirmaActiva19 As Boolean
Global FirmaActiva20 As Boolean
Global FirmaActiva165 As Boolean

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

'FGZ -  10/02/2004
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
Global HuboErrorInesperado As Boolean

'FGZ -  06/02/2004
Global Contador_updates_WF_impproarg As Long
Global Contador_updates_WF_impmesarg As Long
Global Contador_updates_acu_liq As Long
Global Contador_updates_acu_mes As Long

Global objconnProgreso As New ADODB.Connection
Global ContadorProgreso As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Global StrLog As String

Global Cantidad_de_Conceptos As Long
Global Concepto_Actual As Long
Global Arr_conceptos() As tConcepto

Global Arr_EmpCab() As TEmpCabLiq
Global Cantidad_de_Empleados As Long
Global Empleado_Actual As Long

Global Arr_Programa() As TPrograma

Global Arr_con_for_tpa() As TCon_for_tpa
'Mantiene el indice actual del confortpa
Global Indice_Actual As Long

Global Arr_Cge_Segun() As TCge_Segun
'Mantiene el indice actual del cge_Segun
Global Indice_Actual_Cge_Segun As Long

Global Arr_Cft_Segun() As TCft_Segun
'Mantiene el indice actual del cge_Segun
Global Indice_Actual_Cft_Segun As Long

Global Buliq_Concepto() As Tbuliq_concepto
'Mantiene el indice actual de buliq_concepto
Global Indice_Buliq_Concepto As Long

Global Arr_For_Tpa() As TFor_tpa
'Mantiene el indice actual del For_Tpa
Global Indice_Actual_For_Tpa As Long

Global Arr_Acumulador() As tAcumulador
'Mantiene el indice actual del Acumulador
Global Acumulador_Actual As Long

Global Arr_His_Estructura() As THisEstructura
Global Arr_His_EstructuraPer() As THisEstructura
Global Arr_His_EstructuraEmp() As THisEstructura
Global Arr_WF_TPA() As TWF_Tpa

Global Cantidad_de_OpenRecordset As Long
Global Borrar_Estadisticas As Boolean

'Indices Maximos de los arreglos globales
Global Max_Conceptos As Long
Global Max_Cabeceras As Long
Global Max_Programas As Long
Global Max_Con_For_Tpa As Long
Global Max_Cge_Segun As Long
Global Max_Cft_Segun As Long
Global Max_For_Tpa As Long
Global Max_Acumuladores As Long
Global Max_Contrato As Long
Global Max_WF_Tpa As Long
Global Max_His_Estructura As Long
'Global Max_His_EstructuraPer As Long
'Global Max_His_EstructuraEmp As Long
Global LI_WF_Tpa As Long
Global LS_WF_Tpa As Long


'Arreglo de Acumuladores por concepto
' dos dimensiones: Concepto, Acumulador
'Si el valor es -1 ==> el el concepto suma a ese acumulador
'Si el valor es 0 ==> el el concepto NO suma a ese acumulador
Global Arr_Con_Acum() As Long

'FGZ - 10/09/2004
Global HayAcuNoNeg As Boolean

'FGZ - 13/12/2013 -----------------
Global ArrNov() As TDistrib
'FGZ - 13/12/2013 -----------------

'FGZ - 12/10/2004
Global Cantidad_Warnings As Long

Global tplaorden As Long

'FGZ - 10/01/2005
Global Legajo As Long

'FGZ - 20/01/2005
Global Etiqueta

'FAF - 09/08/2006 - Para la Busqueda 50
Global proporciono_bus_DiasVac As Boolean

'FGZ - 27/04/2011 -----
Global Version_Valida As Boolean
'FGZ - 27/04/2011 -----

Global usuario As String

' FGZ - 19/05/2011 ---------------
'Para estadisticas
Global Cantidad_His_Estructura As Long
Global Cantidad_Detliq As Long
Global Cantidad_Select As Long
Global Cantidad_NovG As Long
Global Cantidad_NovE As Long
Global Cantidad_NovI As Long
Global Cantidad_NovA As Long
Global Cantidad_BusI As Long
Global Cantidad_Acumes As Long
Global Cantidad_Acu_liq As Long
Global Cantidad_WF_Tpa As Long
Global Cantidad_WF_Impproarg As Long
Global Cantidad_WF_Retro As Long

'FGZ - 10/02/2012 ------------
Global AnioActualpSAC As Long
Global mesActualpSAC As Long
'FGZ - 10/02/2012 ------------

Global arrConcImpAlcan() As TCImp_Alcance  'EAM- v(2.0) - CAS- 22808

Global TEEmpresa As String
Global TEFundo As String
Global TEParcela As String
Global TECentroCosto As String
Global TEActividad As String
Global DICC_ESTR
Global crear_diccionario_estr As Boolean
Global DICC_UNINRO
Global DICC_UNIDESC
    
Global rs_NovGral As New ADODB.Recordset       'EAM- (6.67) - cargo todas las novedades del empleado
