Attribute VB_Name = "MdlVersiones"
Option Explicit

'Global Const Version = "2.13"
'Global Const FechaModificacion = "13/07/2005"
'Global Const UltimaModificacion = "Nueva Busqueda. Dias Anuales para Vacaciones"

'Global Const Version = "2.14"
'Global Const FechaModificacion = "14/07/2005"
'Global Const UltimaModificacion = "Cambio en la generacion de sanciones de BAE"

'Global Const Version = "2.15"
'Global Const FechaModificacion = "21/07/2005"
'Global Const UltimaModificacion = "Se cambió todas las precisiones de los montos. Single por Double"

'Global Const Version = "2.16"
'Global Const FechaModificacion = "22/07/2005"
'Global Const UltimaModificacion = "Se Agregó el item 30 (Movilidad) como item fijo. Formula de Ganancias."

'Global Const Version = "2.17"
'Global Const FechaModificacion = "22/07/2005"
'Global Const UltimaModificacion = "Se Agregó Modulo de Formulas MDLFormulasGlencore. Modulo de Formulas ciutomizadas para Glencore."

'Global Const Version = "2.18"
'Global Const FechaModificacion = "25/07/2005"
'Global Const UltimaModificacion = "Se Agregaron 2 tipos de busquedas nuevos, customizadas para Glencore."
'Global Const UltimaModificacion1 = "    76 - Antiguedad del empleado (Glencore). bus_Anti_G"
'Global Const UltimaModificacion2 = "    77 - Acum. Mensual meses fijos. (Glencore).bus_Acum3_G"

'Global Const Version = "2.19"
'Global Const FechaModificacion = "26/07/2005"
'Global Const UltimaModificacion = "Se agregó el tipo de busqueda:"
'Global Const UltimaModificacion1 = "    78: SAC Proporcional tope de 30 x mes (bus_DiasSAC_Proporcional_Tope30)"
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.20"
'Global Const FechaModificacion = "26/07/2005"
'Global Const UltimaModificacion = "Se agregó el tipo de busqueda:"
'Global Const UltimaModificacion1 = "    79: Acum. Mensual meses Variables. (Glencore)"
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.21"
'Global Const FechaModificacion = "02/08/2005"
'Global Const UltimaModificacion = "Se agregó chequeo de tipo de concepto <> 5 en Imponibles"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.23"
'Global Const FechaModificacion = "03/08/2005"
'Global Const UltimaModificacion = "Busqueda de embargos"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.24"
'Global Const FechaModificacion = "04/08/2005"
'Global Const UltimaModificacion = "Se modificó el tipo de busqueda:"
'Global Const UltimaModificacion1 = "    78: SAC Proporcional tope de 30 x mes (bus_DiasSAC_Proporcional_Tope30)"
'Global Const UltimaModificacion2 = " Calculaba un dia de mas. Ahora hace 30 - dias no trabajados"

'Global Const Version = "2.25"
'Global Const FechaModificacion = "04/08/2005"
'Global Const UltimaModificacion = "Se modificó el tipo de busqueda:"
'Global Const UltimaModificacion1 = "    54: Dias de Ingreso"
'Global Const UltimaModificacion2 = " Cuando la baja es en el mes anterior a la liq. debe dar 0 dias de ingreso o 30 de inasistencia"

'Global Const Version = "2.26"
'Global Const FechaModificacion = "05/08/2005"
'Global Const UltimaModificacion = "Se modificó el tipo de busqueda:"
'Global Const UltimaModificacion1 = "    46: Dias habiles Mes otra liquidacion"
'Global Const UltimaModificacion2 = " Generacion de la traza."

'Global Const Version = "2.27"
'Global Const FechaModificacion = "08/08/2005"
'Global Const UltimaModificacion = "Se modificó el tipo de busqueda:"
'Global Const UltimaModificacion1 = "    66: BAE" 'Hecha por Javier en TTI
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.28"
'Global Const FechaModificacion = "16/08/2005"
'Global Const UltimaModificacion = "Se modificó que genere el acumulador de desborde cuando hay desborde"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.29"
'Global Const FechaModificacion = "23/08/2005"
'Global Const UltimaModificacion = "Se agregó el tipo de busqueda:"
'Global Const UltimaModificacion1 = "    80: Titulos"
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.30"
'Global Const FechaModificacion = "24/08/2005"
'Global Const UltimaModificacion = "Se agregó el tipo de busqueda:"
'Global Const UltimaModificacion1 = "    81: Dias de Ingreso Con Mes Anterior. (Customizacion para Teleperformance)"
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.31"
'Global Const FechaModificacion = "24/08/2005"
'Global Const UltimaModificacion = "Se modificó la formula de ganancias:"
'Global Const UltimaModificacion1 = "    Si  item_liq + Item_ddjj + item_oldLiq = 0 ==> tope = 0"
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.32"
'Global Const FechaModificacion = "25/08/2005"
'Global Const UltimaModificacion = "Cambios de tipo FLOAT(19) por FLOAT(63)"
'Global Const UltimaModificacion1 = "    ORACLE. Problema en la presicion con nros grandes"
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.33"
'Global Const FechaModificacion = "26/08/2005"
'Global Const UltimaModificacion = "Cambios en busqueda de embargos"
'Global Const UltimaModificacion1 = "    Cantidad de Cuotas"
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.34"
'Global Const FechaModificacion = "02/09/2005"
'Global Const UltimaModificacion = "Cambios en busqueda de embargos"
'Global Const UltimaModificacion1 = "    Desliquidacion de Cuotas"
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.35"
'Global Const FechaModificacion = "05/09/2005"
'Global Const UltimaModificacion = "Si el empleado no esta activo ==> seteo la fecha de baja"
'Global Const UltimaModificacion1 = "   "
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.36"
'Global Const FechaModificacion = "06/09/2005"
'Global Const UltimaModificacion = "Cambios en busqueda de Licencias"
'Global Const UltimaModificacion1 = "  Se agregó parametro de desborde "
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.37"
'Global Const FechaModificacion = "06/09/2005"
'Global Const UltimaModificacion = "Cambios en busqueda de Asignaciones Fliares"
'Global Const UltimaModificacion1 = "  Se agregó parametro de fecha hasta para calculo de edad "
'Global Const UltimaModificacion2 = ""

'Global Const Version = "2.38"
'Global Const FechaModificacion = "06/09/2005"
'Global Const UltimaModificacion = " " 'Cambios en busqueda de Dias de Ingreso
'Global Const UltimaModificacion1 = " " 'And FechaHasta <> buliq_proceso!profecfin
'Global Const UltimaModificacion2 = " " ' + bus_DiasSAC_Proporcional_Tope30

'Global Const Version = "2.39"
'Global Const FechaModificacion = "13/09/2005"
'Global Const UltimaModificacion = " " 'Formula de ganancias
'Global Const UltimaModificacion1 = " " 'se agregó el item Ganancia Bruta en el detalle (traza)
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.40"
'Global Const FechaModificacion = "13/09/2005"
'Global Const UltimaModificacion = " "  'If rs_Fases!bajfec < Empleado_Fecha_Fin And rs_Fases!bajfec > Empleado_Fecha_Inicio Then
'Global Const UltimaModificacion1 = " " '    Empleado_Fecha_Fin = rs_Fases!bajfec
'Global Const UltimaModificacion2 = " " 'End If

'Global Const Version = "2.41"
'Global Const FechaModificacion = "15/09/2005"
'Global Const UltimaModificacion = " "  'Busqueda de Prestamos
'Global Const UltimaModificacion1 = " " 'Se agregó la opcion de retornar: 1)Cuota Total, 2)Cuota Pura o 3)Solo los intereses
'Global Const UltimaModificacion2 = " " 'Por default busca 1).

'Global Const Version = "2.42"
'Global Const FechaModificacion = "20/09/2005"
'Global Const UltimaModificacion = " "  'Formula de provision de vacaciones - Glencore
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.43"
'Global Const FechaModificacion = "23/09/2005"
'Global Const UltimaModificacion = " "  'Cambio en la busqueda de vales
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.44"
'Global Const FechaModificacion = "03/10/2005"
'Global Const UltimaModificacion = " "  'Muestra el valor de los parametros cuando calcula la formula
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.45"
'Global Const FechaModificacion = "04/10/2005"
'Global Const UltimaModificacion = " "  'Traza_gan_item_top
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.46"
'Global Const FechaModificacion = "04/10/2005"
'Global Const UltimaModificacion = " "  'Traza.trafrecuencia con formato de 7 digitos redundantes
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.47"
'Global Const FechaModificacion = "05/10/2005"
'Global Const UltimaModificacion = " "  'Busqueda asig. Fliares.
'Global Const UltimaModificacion1 = " " 'Default de fecha hasta en calculo de edad
'Global Const UltimaModificacion2 = " " 'buliq_proceso!profecfin

'Global Const Version = "2.48"
'Global Const FechaModificacion = "14/10/2005"
'Global Const UltimaModificacion = " "  'Busqueda Licencias segun Periodo GTI. Agregado de logs.
'Global Const UltimaModificacion1 = " " 'Formula de ganancias. Items tipo tope = 5
'Global Const UltimaModificacion2 = " " 'If Items_TOPE(rs_item!itenro) < 0 Then 0 else Items_TOPE(rs_item!itenro) * rs_item!iteporctope / 100

'Global Const Version = "2.49"
'Global Const FechaModificacion = "19/10/2005"
'Global Const UltimaModificacion = " "  'Busqueda Licencias Integrada.
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.50"
'Global Const FechaModificacion = "19/10/2005"
'Global Const UltimaModificacion = " "  'Busqueda Licencias Integrada. Reviso lic marcadas por este proceso.
'Global Const UltimaModificacion1 = " " 'Busqueda de escalas. Dim Parametros() as integer por double
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.51"
'Global Const FechaModificacion = "20/10/2005"
'Global Const UltimaModificacion = " "  'Busqueda Licencias Integrada. Reviso lic marcadas por este proceso. Descuenta dif con ya marcadas.
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.52"
'Global Const FechaModificacion = "21/10/2005"
'Global Const UltimaModificacion = " "  'Busqueda Licencias Integrada. Reviso lic marcadas por este proceso. Descuenta dif con ya marcadas.
'Global Const UltimaModificacion1 = " " 'Busqueda Conceptos meses fijos. Si el mes de inicio es ninguno==> lo calcula de acuerdo al mes actual.
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.53"
'Global Const FechaModificacion = "25/10/2005"
'Global Const UltimaModificacion = " "  'Busqueda Dias SAC Proporcional Tope 30. Resto 1 al ultimo rango si no es completo.
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.54"
'Global Const FechaModificacion = "27/10/2005"
'Global Const UltimaModificacion = " "  'Busqueda Licencias Integrada.
'Global Const UltimaModificacion1 = " " 'If (SumaDias + Dias_Ya_Marcados) < DiasDelMes And Dias_Totales = (SumaDias + Dias_Ya_Marcados) Then
'Global Const UltimaModificacion2 = " " 'Quiere decir que trabajó algun dia ==> le descuento los dias trabajados a la ultima lic que se calcule

'Global Const Version = "2.55"
'Global Const FechaModificacion = "31/10/2005"
'Global Const UltimaModificacion = " "  'Busqueda Escalas.
'Global Const UltimaModificacion1 = " " 'Logs y control indice de parametros variables
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.56"
'Global Const FechaModificacion = "31/10/2005"
'Global Const UltimaModificacion = " "  'Tipo Busqueda 83 - Dias Habiles Trabajados.
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.57"
'Global Const FechaModificacion = "02/11/2005"
'Global Const UltimaModificacion = " "  'Modificacion en Tipo Busqueda 78
'Global Const UltimaModificacion1 = " " 'Modificacion en Tipo Busqueda 82
'Global Const UltimaModificacion2 = " " 'Modificacion en Formula de Ganancias
' Descripcion: Se agregaron 3 campos nuevos a traza_gan que estan relacionados con el F649.
'               traza_gan.deducciones decimal(19,4)
'               traza_gan.art23 decimal(19,4)
'               traza_gan.porcdeduc decimal(19,4)

'Global Const Version = "2.58"
'Global Const FechaModificacion = "08/11/2005"
'Global Const UltimaModificacion = " "  'Modificacion en Tipo Busqueda 78
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.59"
'Global Const FechaModificacion = "30/11/2005"
'Global Const UltimaModificacion = " "  'Procedimiento establecer_empresa y en consecuencia ...
'Global Const UltimaModificacion1 = " " '    Formulas : IRP, IRP_Franja, Ganancias, Ganancias_Schering
'Global Const UltimaModificacion2 = " " '               DesaProvSac, DesaProvVac. PorcPres

'Global Const Version = "2.60"
'Global Const FechaModificacion = "02/12/2005"
'Global Const UltimaModificacion = " "  'Tipo de busuqeda bus_Cotmon1.
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.61"
'Global Const FechaModificacion = "02/12/2005"
'Global Const UltimaModificacion = " "  'Busqueda de Pago/Dto de Vaccaciones.
'Global Const UltimaModificacion1 = " " 'Retorna 0 cuando No se encontraron Pagos/Dtos
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.62"
'Global Const FechaModificacion = "07/12/2005"
'Global Const UltimaModificacion = " "  'Busqueda de licencias(31) y licencias integradas(82)
'Global Const UltimaModificacion1 = " " 'Control si el empleado esta dado de baja en el periodo
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.63"
'Global Const FechaModificacion = "13/12/2005"
'Global Const UltimaModificacion = " "  'Busqueda de Vales
'Global Const UltimaModificacion1 = " " 'Se agrego la opcion 4: Segun fecha del proceso
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.64"
'Global Const FechaModificacion = "13/12/2005"
'Global Const UltimaModificacion = " "  'Se cambio la busqueda 78, SAC Proporcional tope de 30 x mes
'Global Const UltimaModificacion1 = " " 'No considerar la fecha de baja prevista
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.65"
'Global Const FechaModificacion = "05/01/2006"
'Global Const UltimaModificacion = " "  'Se cambio la busqueda 78, SAC Proporcional tope de 30 x mes
'Global Const UltimaModificacion1 = " " 'No considerar la fecha de baja fuese de un año o semestre anterior.
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.66"
'Global Const FechaModificacion = "13/01/2006"
'Global Const UltimaModificacion = " "  'Nuevo Tipo de Busqueda:
'Global Const UltimaModificacion1 = " " '    84 - Datos de Ganancias
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.67"
'Global Const FechaModificacion = "16/01/2006"
'Global Const UltimaModificacion = " "  'Modificacion Tipo de Busqueda:
'Global Const UltimaModificacion1 = " " '    51 - bus_Vac_No_Gozadas_Pendientes
'Global Const UltimaModificacion2 = " " ' La proporcion se debe hacer solo sobre el ultimo año

'Global Const Version = "2.68"
'Global Const FechaModificacion = "27/01/2006"
'Global Const UltimaModificacion = " "  'Modificacion Tipo de Busqueda:
'Global Const UltimaModificacion1 = " " '    7 - Acumualdores meses Fijos
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.69"
'Global Const FechaModificacion = "14/02/2006"
'Global Const UltimaModificacion = " "  'Nueva Formulña de Sistema para Divino SA (Uruguay):
'Global Const UltimaModificacion1 = " " ' for_Ribeteado
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.70"
'Global Const FechaModificacion = "14/02/2006"
'Global Const UltimaModificacion = " "  'Modificacion en Formula:
'Global Const UltimaModificacion1 = " " ' for_Ribeteado
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.71"
'Global Const FechaModificacion = "15/02/2006"
'Global Const UltimaModificacion = " "  'Modificacion Tipo de Busqueda:
'Global Const UltimaModificacion1 = " " ' 27 - Pago / Descuento de Licencias
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.72"
'Global Const FechaModificacion = "16/02/2006"
'Global Const UltimaModificacion = " "  'Modificacion Tipo de Busqueda:
'Global Const UltimaModificacion1 = " " ' 82 - Licencias Integradas. Andaba mal el desborde
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.73"
'Global Const FechaModificacion = "20/02/2006"
'Global Const UltimaModificacion = " "  'Modificacion Tipo de Busqueda:
'Global Const UltimaModificacion1 = " " ' 2 - Internas. Si el sql retorna eof ==> retrona true tambien
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.74"
'Global Const FechaModificacion = "20/03/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Modificacion Tipo de Busqueda:
'Global Const UltimaModificacion1 = " " ' 20 - Cotizacion de Monedas. Faltaba la funcion convfecha() en el sql
'Global Const UltimaModificacion2 = " " ' 66 - Busqueda de BAE. Procedimiento Generar_Sanciones. Rutina para busqueda anual
'                                       ' 45 - Asignaciones Familiares. Se agregó un control del periodo de estudio en el estudio actual
'                                       '      StrSql = StrSql & " AND estudio_actual.perescnro = " & Periodo_Escolar
'                                       ' Tambien se sacaron las advertencias (logs) del desmarcado de BAE cuando la tabla no existe

'Global Const Version = "2.75"
'Global Const FechaModificacion = "27/03/2006"   'HJI
'Global Const UltimaModificacion = " "  'Modificacion Tipo de Busqueda:
'Global Const UltimaModificacion1 = " " ' 66 - Busqueda de BAE. Procedimiento Generar_Sanciones. Rutina para busqueda anual
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.76"
'Global Const FechaModificacion = "18/04/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Modificacion en liqpro04:   'desmarcado de licencias y pagos/dtos
'Global Const UltimaModificacion1 = " " ' no desmarcaba pagos/dtos generados a partir de dias correspondientes de vacaciones
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.77"
'Global Const FechaModificacion = "08/05/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Evaluete de expresion con overflow
'Global Const UltimaModificacion1 = " " 'Busqueda 54 (dias de Ingreso). se agregó un parametro a la busqueda que denota si el ultimo dia se considera trabajado o no.
'                                        'Nuevos asp: busq_diasingreso_liq_00.asp, busq_diasingreso_liq_01.asp
'Global Const UltimaModificacion2 = " "  'Busqueda 81 (dias de Ingreso contempla mes anterior). se agregó un parametro a la busqueda que denota si el ultimo dia se considera trabajado o no.
'                                        'Nuevos asp: busq_diasingreso_liq_00.asp, busq_diasingreso_liq_01.asp

'Global Const Version = "2.78"
'Global Const FechaModificacion = "11/05/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Busqueda 78 (sac proporcional con tope 30 dias). Cuando la baja es en febrero estaba dando 1 o 2 dias de mas.
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.79"
'Global Const FechaModificacion = "18/05/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Busqueda 7 (Acum Meses Fijos).
'Global Const UltimaModificacion1 = " " '    se cambio cuando la busqueda no es a mes completo ==> DividePor lo paso en 0 salvo que
'                                       '    sea con fase activa ==> contempla la cantidad de meses del semestre/año - los meses fuera de fase
'Global Const UltimaModificacion2 = " " 'Operacion promedio AM_PROM.
'                                            'si el parametro DividePor viene en 0 ==> divido por cantProm

'Global Const Version = "2.80"
'Global Const FechaModificacion = "19/05/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Busqueda 7 (Acum Meses Fijos).
'Global Const UltimaModificacion1 = " " '    Anual, con periodo actual sin proces actual andaba mal
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.81"
'Global Const FechaModificacion = "23/05/2006"   'MB
'Global Const UltimaModificacion = " "  'Formula de Ganacias.
'Global Const UltimaModificacion1 = " " 'Impuestos a los debitos bancarios se puso como resta de imp determinado
'Global Const UltimaModificacion2 = " "

'....

'Global Const Version = "2.83"
'Global Const FechaModificacion = "05/06/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Nuevo tipo de Busuqeda 85: Fecha en formato numerico
'Global Const UltimaModificacion1 = " " 'Nuevo tipo de Busuqeda 86: Acum. Mensual meses variables con Ajuste de aumentos
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.84"
'Global Const FechaModificacion = "22/06/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Busqueda 78 (sac proporcional con tope 30 dias).
'Global Const UltimaModificacion1 = " " 'Cuando la fecha de fin de calculo es la fecha de fin de semestre se come un dia.
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.85"
'Global Const FechaModificacion = "23/06/2006"   'FGZ
'Global Const UltimaModificacion = " "  '82 'Licencias Integradas.
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "


'Global Const Version = "2.86"
'Global Const FechaModificacion = "26/06/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Busqueda 8 (Acum Meses Variables).
'Global Const UltimaModificacion1 = " " '    Cuando se elige la opcion promedio siempre busca las fases aun cuando se configure sin fase activa
'                                       '    sea con fase activa ==> contempla la cantidad de meses del semestre/año - los meses fuera de fase
'Global Const UltimaModificacion2 = " " 'Operacion promedio AM_PROM.
'                                            'si el parametro DividePor viene en 0 ==> divido por cantProm


'Global Const Version = "2.87"
'Global Const FechaModificacion = "26/06/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Tipo de Busuqeda 86: Acum. Mensual meses variables con Ajuste de aumentos
'Global Const UltimaModificacion1 = " " '    Operaciones Promedio(bug, Cuando se elige la opcion promedio siempre busca las fases aun cuando se configure sin fase activa) y promedio sin 0 (Agregado de logs)
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.88"
'Global Const FechaModificacion = "27/06/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Tipo de Busuqeda 86: Acum. Mensual meses variables con Ajuste de aumentos
'Global Const UltimaModificacion1 = " " '    Operaciones promedio sin 0 (no estaba acumulando bien)
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.89"
'Global Const FechaModificacion = "05/07/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Tipo de Busuqeda 54: Dias de Ingreso
'Global Const UltimaModificacion1 = " " '    Cuando la baja es posterior al fin del proceso ==> el ultimo dia siempre se considera trabajado
'Global Const UltimaModificacion2 = " " 'Tipo de Busuqeda 78 (sac proporcional con tope 30 dias).
'                                       '    Cuando la fecha de inicio de calculo es mayor a la fecha de inicio de semestre se come 2 dias.

'Global Const Version = "2.90"
'Global Const FechaModificacion = "07/07/2006"   'Martin Ferraro
'Global Const UltimaModificacion = " "  'En liqpro04 cuando se desmarcaban los vales, se desmarcaban todos y no los
'Global Const UltimaModificacion1 = " " 'los del empleado actual, marcando solo como liquidados los del vales del
'Global Const UltimaModificacion2 = " " 'ultimo empleado
                                      
'Global Const Version = "2.91"
'Global Const FechaModificacion = "24/07/2006"   'Fapitalle N.
'Global Const UltimaModificacion = " "  'Se agrego al tipo de busqueda Antigüedad del Empleado (10)
'Global Const UltimaModificacion1 = " " ' la opcion de calcularla a fin de la primera quincena
'Global Const UltimaModificacion2 = " " '
                                      
'Global Const Version = "2.92"
'Global Const FechaModificacion = "08/08/2006"   'Breglia M.
'Global Const UltimaModificacion = " "  'Se agrego al tipo de busqueda novgegi que inserte 0 en vigencia
'Global Const UltimaModificacion1 = " " ' para novedades historicas por estructura y globales
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "2.93"
'Global Const FechaModificacion = "09/08/2006"   'Fernando Favre
'Global Const UltimaModificacion = " "  'Tipo de Busqueda 22: Embargos
'Global Const UltimaModificacion1 = " " '  Cuando el embargo es de tipo 'Embargo Judicial por % Fijo' la ultima cuota la descuenta mal. Ej: le queda por pagar $100 y la cuota le quedo definida por $200, descuenta $200. Debe descontar $100
'Global Const UltimaModificacion2 = " " 'Tipo de Busqueda 50: Vacaciones no Gozadas
'                                       '  Si la antiguedad es <=6 redondea 2 veces.

'Global Const Version = "2.94"
'Global Const FechaModificacion = "17/08/2006"   'FGZ
'Global Const UltimaModificacion = " "  '82 'Licencias Integradas.
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.95"
'Global Const FechaModificacion = "24/08/2006"   'Martin Ferraro
'Global Const UltimaModificacion = " "  'Se modificaron las busquedas 31, 43, 48, 60, 82
'Global Const UltimaModificacion1 = " " 'para contemplar los estados de las licencias
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.96"
'Global Const FechaModificacion = "30/08/2006"   'FGZ
'Global Const UltimaModificacion = " "  'Modulo de clase CEval 'Funcion RED: estaba haciendo un Cint() y daba error con nros grandes se reemplazo por cdbl().
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "2.97"
'Global Const FechaModificacion = "04/09/2006"   'FGZ
'Global Const UltimaModificacion = " "   'Modulo MdlBuliq. Procedimiento Establecer_Empleado:
'Global Const UltimaModificacion1 = " "  'Cuando la fecha de baja del legajo que seguia era mayor que el anterior no seteaba la fecha de baja,
'Global Const UltimaModificacion2 = " "  'dejaba la fecha de baja del legajo anterior.

'Global Const Version = "2.98"
'Global Const FechaModificacion = "03/10/2006"   'FGZ
'Global Const UltimaModificacion = " "   ' 7 - Acumuladores Meses Fijos:
'Global Const UltimaModificacion1 = " "  '       le puse este control porque aveces se pasaba de meses. Cuando el mes de inicio es fijo y no son ni Julio ni Enero
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "2.99"
'Global Const FechaModificacion = "09/10/2006"   'Martin Ferraro
'Global Const UltimaModificacion = " "   ' En el modulo MdlTiposBusquedas, en la Subrutina "Generar_Sanciones" se agrego codigo generado por TTI
'Global Const UltimaModificacion1 = " "  ' esta entre lineas ----- Debug -- 17/07/2006
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "3.00"
'Global Const FechaModificacion = "12/10/2006"   'FGZ
'Global Const UltimaModificacion = " "   ''Modulo de clase CEval 'Funcion RED:
'Global Const UltimaModificacion1 = " "  '      en la version 2.96 cambié el Cint() por Cbdl() porque habia un error con nros grandes ==>
'Global Const UltimaModificacion2 = " "  '      el temas es que esa modificacion causó que cuando el redondeo es hacia arriba o hacia abajo no funcione
'                                        '      Lo que hice ahora es volver a reemplazar la funcion cdbl() por Fix()


'Global Const Version = "3.01"
'Global Const FechaModificacion = "24/10/2006"   'FGZ
'Global Const UltimaModificacion = " "   ''Modulo de tipos de busquedas:
'Global Const UltimaModificacion1 = " "  '      Tipo de Busqueda Antigüedad del Empleado (10)
'Global Const UltimaModificacion2 = " "  '           Cuando se agregó la opcion de calcularla a fin de la primera quincena(version 2.91)
                                                    'se introdujo un error que se propagó hasta la version 3.00
                                                    'Cuando se utilice la busqueda de antiguedad a fecha de alta reconocida ... NO FUNCIONA

                                        'Se corrigió el sub bus_Anti0

'Global Const Version = "3.02"
'Global Const FechaModificacion = "27/11/2006"   'Breglia Maximiliano
'Global Const UltimaModificacion = " "   'Nueva formula de Grossing UP
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "


'Global Const Version = "3.03"
'Global Const FechaModificacion = "26/01/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Nueva formula de Ganancias para Petroleros
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.04"
'Global Const FechaModificacion = "08/02/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Modificacion en formula de Ganancias para Petroleros
'Global Const UltimaModificacion1 = " "  'Se agregó el parametro 1024 para ajustar retenciones de meses anteriores
'Global Const UltimaModificacion2 = " "


'Global Const Version = "3.05"
'Global Const FechaModificacion = "08/02/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Modificacion en formula de Ganancias para Petroleros
'Global Const UltimaModificacion1 = " "  'Se corrigió el ajuste ajustar retenciones de meses anteriores
'Global Const UltimaModificacion2 = " "  'Se modificó la formula IRP

'Global Const Version = "3.06"
'Global Const FechaModificacion = "09/02/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Se modificó la formula IRP
'Global Const UltimaModificacion1 = " "  '
'Global Const UltimaModificacion2 = " "  '


'Global Const Version = "3.07"
'Global Const FechaModificacion = "09/02/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Se modificó la formula IRP, el calculo del porcentaje
'Global Const UltimaModificacion1 = " "  '
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "3.08"
'Global Const FechaModificacion = "26/02/2007"   'Martin Ferraro
'Global Const UltimaModificacion = " "   'Busqueda Vales - Se agrego opcion todos/revisados
'Global Const UltimaModificacion1 = " "  '
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "3.09"
'Global Const FechaModificacion = "08/03/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Busqueda BAE - Se modificó el procedimiento Generar_Sanciones
'Global Const UltimaModificacion1 = " "  '
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "3.10"
'Global Const FechaModificacion = "19/03/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Busqueda BAE - Se modificó el procedimiento Generar_Sanciones.
'Global Const UltimaModificacion1 = " "  '   Estaba calculando mal la antiguedad.
'Global Const UltimaModificacion2 = " "  '   Estaba calculando mal las penalidades anuales.

'Global Const Version = "3.11"
'Global Const FechaModificacion = "30/03/2007"   'Breglia Maximiliano
'Global Const UltimaModificacion = " "   'Busqueda Embargos - Se modificó porque activaba cq embargo y no los empleado cuando
'Global Const UltimaModificacion1 = " "  'habia embargos pendientes del mismo tipo
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "3.12"
'Global Const FechaModificacion = "19/04/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Liqpro06 - habia un error en el Topes de imponible para contribuciones.
'Global Const UltimaModificacion1 = " "  '               Hacia un .... ipamonto = " & Aux_impproarg2 - 1
'Global Const UltimaModificacion2 = " "  'Se le sacó

'Global Const Version = "3.13"
'Global Const FechaModificacion = "08/05/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Busqueda BAE -
'Global Const UltimaModificacion1 = " "  '   sub bus_BAE(). Cambio en la fecha hasta de penalidades para meses que no sean marzo
'Global Const UltimaModificacion2 = " "  '   Agregados de logs
''                                           sub generar_sanciones() cuando es anual La rotacion esta alreves

'Global Const Version = "3.14a"
'Global Const FechaModificacion = "22/06/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Nueva Formula : IRPF (Impuesto Renta a Personas Fisicas)
'Global Const UltimaModificacion1 = " "  'Nuevo tipo de busqueda:
'Global Const UltimaModificacion2 = " "  '   87: Deduccion Fliares para IRPF. (Customizacion para Uruguay)"
'                                        '   88: BPC para IRPF. (Customizacion para Uruguay)"

'Global Const Version = "3.15"
'Global Const FechaModificacion = "26/06/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Modificacion Liqpro06: Topeo imponible para SAC.
'Global Const UltimaModificacion1 = " "  '   Lo estaba haciendo mal cuando ya tenia algun imponible acumulado en meses anteriores
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "3.16"
'Global Const FechaModificacion = "26/06/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Nueva Formula : IRPF_SIMPLE (Basicamente = que IRPF pero sin traza_gan, ficharet, items, etc)
'Global Const UltimaModificacion1 = " "  'Modif Formula : IRPF
'Global Const UltimaModificacion2 = " "  '   No se restan otras deducciones para el calculo del impuesto

'Global Const Version = "3.17"
'Global Const FechaModificacion = "26/06/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Modificacion Liqpro06: Topeo imponible para SAC.
'Global Const UltimaModificacion1 = " "  '   Lo estaba haciendo mal cuando ya tenia algun imponible acumulado en meses anteriores
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "3.18"
'Global Const FechaModificacion = "27/06/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Modificacion Busquedas de Antiguedad:
'Global Const UltimaModificacion1 = " "  '   Estaba Calculando 1 dia de menos cuando la antiguedad era menor el año
'Global Const UltimaModificacion2 = " "  '   sub bus_Antiguedad(),
''                                           sub bus_Antiguedad_A_FechaAlta(),
''                                           sub bus_Antiguedad_FechaAltaReconocida(),
''                                           sub bus_Antfases,Bus_Edad_Empleado()

'Global Const Version = "3.19"
'Global Const FechaModificacion = "28/06/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Modif Formula : IRPF
'Global Const UltimaModificacion1 = " "  '   error de inicializacion de acumulador de proceso actual
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "3.20"
'Global Const FechaModificacion = "13/08/2007"   'FGZ
'Global Const UltimaModificacion = " "   'Modif liqpro06
'Global Const UltimaModificacion1 = " "  '   Problemas con topeo de imponible para conceptos de tipo 1.
'Global Const UltimaModificacion2 = " "  '   Estaba topeando mal cuando teniamos 2 quincenas de 15 dias cada una y la 1era Q no llega al tope pero
                                        '   la suma de la 1era + 2da supera el tope proporcianal ==> andaba mal

'Global Const Version = "3.21"
'Global Const FechaModificacion = "15/08/2007"   'Martin Ferraro
'Global Const UltimaModificacion = " "   ' la busqueda bus_Acum3 calculaba mal la cantidad de meses
'Global Const UltimaModificacion1 = " "  ' para Anual con periodo actual sin proces actual andaba mal
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "3.22"
'Global Const FechaModificacion = "21/09/2007"   'Martin Ferraro
'Global Const UltimaModificacion = " "   ' la busqueda de vales se modifico para que busque por tipos de vale
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.23"
'Global Const FechaModificacion = "21/09/2007"   'Martin Ferraro
'Global Const UltimaModificacion = " "    ' bus_AsignacionesFliares() - Problemas con Estudia
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.24"
'Global Const FechaModificacion = "27/09/2007"   'FGZ
'Global Const UltimaModificacion = " "    ' Modificacion formula de Ganancias para Petroleros
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.25"
'Global Const FechaModificacion = "19/10/2007"   'FGZ
'Global Const UltimaModificacion = " "    'Se agregó el tipo de busqueda:
'Global Const UltimaModificacion1 = " "   '89: Licencias en otros periodos"
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.26"
'Global Const FechaModificacion = "06/11/2007"   'FGZ
'Global Const UltimaModificacion = " "    'tipo de busqueda: Novedades de GTI
'Global Const UltimaModificacion1 = " "   '  Nunca devolvia TRUE
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.27"
'Global Const FechaModificacion = "09/11/2007"   'Breglia M
'Global Const UltimaModificacion = " "    'cuando actulizaba el proceso por acumul negativos daba dedlock
'Global Const UltimaModificacion1 = " "   'cuando se ejecutaban en parelelo varios procesos
'Global Const UltimaModificacion2 = " "   'se agregó el modulo de formulas de chile

'Global Const Version = "3.28"
'Global Const FechaModificacion = "29/11/2007" 'Martin Ferraro
'Global Const UltimaModificacion = " "    'Cambio en bus_grilla0
'Global Const UltimaModificacion1 = " "   'No especificaba cual de los ordenes tomar de la escala
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.29"
'Global Const FechaModificacion = "07/12/2007" 'Maximiliano Breglia
'Global Const UltimaModificacion = " "    'Cambio en busq dias ingreso mes anterior
'Global Const UltimaModificacion1 = " "   'Para que tome el ultimodia no trabajado
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.30"
'Global Const FechaModificacion = "10/12/2007" 'Martin Ferraro
'Global Const UltimaModificacion = " " 'busq de licencias en horas
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.31"
'Global Const FechaModificacion = "02/01/2008" 'Maximiliano Breglia
'Global Const UltimaModificacion = " "  'bus_acum4 restaba 2 años cuando la combinacion era 1 anio y 0 meses
'Global Const UltimaModificacion1 = " " 'y bus_Acum_Con_Ajuste lo mismo
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.32"
'Global Const FechaModificacion = "16/01/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Recalculo Impuesto Unico para Chile
'Global Const UltimaModificacion1 = " " 'Maxi 16/01/2008 error en tope de SAC cuando tiene algo acum en el semestre
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.33"
'Global Const FechaModificacion = "29/01/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Busquedas Recalculo Impuesto Unico
'Global Const UltimaModificacion1 = " " 'Modificacion Formula Impuesto Unico
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.34"
'Global Const FechaModificacion = "31/01/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Modificacion Busquedas Recalculo Impuesto Unico
'Global Const UltimaModificacion1 = " " 'Nueva Modificacion Formula Impuesto Unico
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.35"
'Global Const FechaModificacion = "01/02/2008" 'FGZ
'Global Const UltimaModificacion = " "  'Modificacion en topeo de imponibles... si los montos son negativos ==> los llevo a 0
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.36"
'Global Const FechaModificacion = "22/02/2008" 'Breglia Maximiliano
'Global Const UltimaModificacion = " "  'Formula de Ganancias - impuesto y debito bancarios
'Global Const UltimaModificacion1 = " " 'Si hay devolucion suma los impdebitos
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.37"
'Global Const FechaModificacion = "18/03/2008" 'Diego Rosso
'Global Const UltimaModificacion = " "  'Se agrego una nueva busqueda de dias llamada "bus_Cant_Dias_Prop()"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.38"
'Global Const FechaModificacion = "27/03/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Se modifico la busqueda de antiguedad para agregar la opcion Primer dia año del siguiente
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.39"
'Global Const FechaModificacion = "17/04/2008" 'FGZ
'Global Const UltimaModificacion = " "  'Se modifico la busqueda de bus_Vac_No_Gozadas_Pendientes
'Global Const UltimaModificacion1 = " " '        Estaba proporcionando mal los dias del ultimo año
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.40"
'Global Const FechaModificacion = "25/04/2008" 'FGZ
'Global Const UltimaModificacion = " "  'Se modifico la busqueda de bus_Vales
'Global Const UltimaModificacion1 = " " '        Estaba marcando todos los vales
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.41"
'Global Const FechaModificacion = "29/04/2008" 'FGZ
'Global Const UltimaModificacion = " "  'Se modifico el sub LiqPro06
'Global Const UltimaModificacion1 = " " '        'Se agregó detalle de log para que ayude a encontrar cual es el acumulador negativo
'Global Const UltimaModificacion2 = " "          'ACUMULADORES NEGATIVO ------> : " & Acum & " monto: " & Aux_Acu_Monto

'Global Const Version = "3.42"
'Global Const FechaModificacion = "13/05/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Se modifico la buqueda bus_Cant_Dias_Prop()
'Global Const UltimaModificacion1 = " " 'se agrego la opcion de controlar desde licencia
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.43"
'Global Const FechaModificacion = "29/05/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Se agrego orden a sql de setear AMPO
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.44"
'Global Const FechaModificacion = "11/06/2008" 'Breglia Maximiliano
'Global Const UltimaModificacion = " "  'Se modifico la busqueda de bus_Vales faltaba en 2 lugares en el toque del 3.40
'Global Const UltimaModificacion1 = " " 'Estaba marcando todos los vales de cualquier tipo si filtre por tipo en la busq
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.45"
'Global Const FechaModificacion = "25/06/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Se agrego la busq bus_AntEnEstructura()
'Global Const UltimaModificacion1 = " " 'bus_Licencias_Integradas(): se topea fecha fin a la fecha de baja
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.46"
'Global Const FechaModificacion = "10/07/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'bus_Cant_Dias_Prop() - Las licencias a descontar deben estar aprobadas y
'Global Const UltimaModificacion1 = " " 'en el caso de las licencias de accidente empresa y ART  además de aprobadas con
'Global Const UltimaModificacion2 = " " 'estado de alta.

'Global Const Version = "3.47"
'Global Const FechaModificacion = "18/07/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Se creo el tipo de busqueda 95 - "Angiguedad del empleado vacaciones" a partir
'Global Const UltimaModificacion1 = " " 'de la busque 10 - "Antiguedad del empleado" con la opcion primer dia de año siguiente
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "3.48"
'Global Const FechaModificacion = "08/08/2008" 'Breglia Maximiliano
'Global Const UltimaModificacion = " "  'Cambio en el tope mopre cuando tenia 2 liquidaciones en el mes se agrego el tope_mes
'Global Const UltimaModificacion1 = " " 'porque no lo tenia en cuenta y liquidaba doble el tope
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.49"
'Global Const FechaModificacion = "01/09/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Busqueda provisión Desaprovisión de Vacaciones
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.50"
'Global Const FechaModificacion = "02/09/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Cambios en log de busq embargos mensuales
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.51"
'Global Const FechaModificacion = "03/09/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Cambios en log de busq embargos mensuales y la desliq de embargos
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.52"
'Global Const FechaModificacion = "03/09/2008" 'Martin Ferraro
'Global Const UltimaModificacion = " "  'Cambios en log de busq embargos mensuales y la desliq de embargos
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "


'Global Const Version = "3.53"
'Global Const FechaModificacion = "24/09/2008" 'FGZ
'Global Const UltimaModificacion = " "  'Se modificó el liqpro06 en el topeo de imponibles para los conceptos de tipo 1
'Global Const UltimaModificacion1 = " " 'Corrige un detalle que quedó del la version 3.48
'Global Const UltimaModificacion2 = " " '  Se cambió el > por el >=
'                                       'Se modificó el tipo de busqueda 96 (provision desaprovision de vacaciones)

'Global Const Version = "3.54"
'Global Const FechaModificacion = "25/09/2008" 'FGZ
'Global Const UltimaModificacion = " "  'Se creo el tipo de busqueda 97 - "Nueva Angiguedad del empleado" a partir
'Global Const UltimaModificacion1 = " " 'de la busque 10 - "Antiguedad del empleado" con la diferencia en la forma de calculo
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "3.55"
'Global Const FechaModificacion = "25/09/2008"
'Global Const UltimaModificacion = " Martin Ferraro"  'Se creo el tipo de busqueda 98 - Busqueda de movimientos
'Global Const UltimaModificacion1 = " " 'Corrige un detalle que quedó del la version 3.48
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.56"
'Global Const FechaModificacion = "08/10/2008"
'Global Const UltimaModificacion = " Martin Ferraro"  'Se creo la formula for_irpf_diciembre para uruguay
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.57"
'Global Const FechaModificacion = "17/10/2008"
'Global Const UltimaModificacion = " FGZ"  'Se Modifico el tipo de busqueda 97 - "Nueva Antiguedad del empleado"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.58"
'Global Const FechaModificacion = "31/10/2008"
'Global Const UltimaModificacion = " MB"  'Se Modifico la formula de Grossing para permitir varios conceptos de gross Gym Cochera
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.59"
'Global Const FechaModificacion = "24/11/2008"
'Global Const UltimaModificacion = " Martin"  'Busqueda irpf_diciembre, se busca en escalas al 25/12
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.60 bis"
'Global Const FechaModificacion = "27/11/2008"
'Global Const UltimaModificacion = " Maxi - Martin"  'Sacar TRAZA de acumuladores e imponibles
'Global Const UltimaModificacion1 = " "              'Reusa Traza config por empresa
'Global Const UltimaModificacion2 = " "              'Irpf diciembre
'                                                    'Grossing para mas de un concepto

'Global Const Version = "3.61"
'Global Const FechaModificacion = "05/12/2008"
'Global Const UltimaModificacion = "Martin"  'No estaba la llamada de la formula de irpf diciembre
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
                                      

'Global Const Version = "3.62"
'Global Const FechaModificacion = "11/12/2008"
'Global Const UltimaModificacion = "Martin"  'Se pasaron las funciones de InsertarTraza y LimpiarTraza del modulo varios al del
'Global Const UltimaModificacion1 = " "      'liquidador poque usaban la variable global ReusaTraza que generaba problemas otros procesos
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.63"
'Global Const FechaModificacion = "17/12/2008"
'Global Const UltimaModificacion = "MB" 'Para el SAC de diciembre 2008 (item 50) se resta el monto para entrar a escala de deducciones
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.64"
'Global Const FechaModificacion = "19/01/2009"
'Global Const UltimaModificacion = "Martin" 'Cambios en la formula del impuesto unico para que tenga en cuenta el acu_mes cuando busca acumuladores de la liquidacion
'Global Const UltimaModificacion1 = " "     ' Dos nuevas formulas de Chile Imp Unico: for_RecalcConcepto  for_RecalcImpuestoUnico
'Global Const UltimaModificacion2 = " "
     
     
'Global Const Version = "3.65"
'Global Const FechaModificacion = "27/01/2009"
'Global Const UltimaModificacion = "Martin" 'for_ImpuestoUnico: se quito la traza
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
     
'Global Const Version = "3.66"
'Global Const FechaModificacion = "29/01/2009"
'Global Const UltimaModificacion = "Martin" 'Se creo la funcion EsModeloRecalculo para que no diera error en
'Global Const UltimaModificacion1 = " "     'en empresas que no son de chile
'Global Const UltimaModificacion2 = " "
    
'Global Const Version = "3.67"
'Global Const FechaModificacion = "19/02/2009"
'Global Const UltimaModificacion = "Maxi" 'Modificacion de formula de recalculo de Imp Unico
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.68"
'Global Const FechaModificacion = "25/02/2009"
'Global Const UltimaModificacion = "FGZ" 'Modificacion en busqueda de Acumuladores Fijos
'Global Const UltimaModificacion1 = " "  ' Cuando se seleccionaba MESUAL del Periodo/Semestre anterior hacia macanas con el mes de inicio
'Global Const UltimaModificacion2 = " "  ' Tambien se agregó Encriptacion de string de conexion

'Global Const Version = "3.69"
'Global Const FechaModificacion = "19/03/2009"
'Global Const UltimaModificacion = "MB" 'Modificacion en Tope Mopre por decimales
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.70"
'Global Const FechaModificacion = "27/04/2009"
'Global Const UltimaModificacion = "MB" 'error en tope de SAC cuando tiene algo acum en el semestre
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.71"
'Global Const FechaModificacion = "21/05/2009"
'Global Const UltimaModificacion = " FGZ"  'Se modificó el tipo de busqueda 98 - Busqueda de movimientos
'Global Const UltimaModificacion1 = " " 'Se busqcan solo movimientos que pertenecen a partes cerrados
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.72"
'Global Const FechaModificacion = "21/05/2009"
'Global Const UltimaModificacion = "Martin"  'Se modificó la formula de ganancias - Se amplio el limite de los
'Global Const UltimaModificacion1 = " "      'arreglos de items a 100
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.73"
'Global Const FechaModificacion = "11/06/2009"
'Global Const UltimaModificacion = "Martin"  'No lock para traza
'Global Const UltimaModificacion1 = " "      'Cambios en Busquedas de embargos
'Global Const UltimaModificacion2 = " "      'Busqueda de antiguedad: descontar licencias

'Global Const Version = "3.74"
'Global Const FechaModificacion = "11/06/2009"
'Global Const UltimaModificacion = "Martin"  'Cambio Mopre para SAC proporcional
'Global Const UltimaModificacion1 = " "      'Nueva formula de sistema de SAC no Remunerativo
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.75"
'Global Const FechaModificacion = "30/06/2009"
'Global Const UltimaModificacion = "Martin"  'Liqpro04 Desmarcaba todas las licencias y no las del empleado, error en borrado inmesarg
'Global Const UltimaModificacion1 = " "      'Liqpro04 Evalua si hay detliq antes de hacer todo el analisis de cosas a borrar
'Global Const UltimaModificacion2 = " "      'liqpro06 Se agrego usadebug a los logs del mopre

'Global Const Version = "3.76"
'Global Const FechaModificacion = "03/07/2009"
'Global Const UltimaModificacion = "Martin"  'Liqpro06 Cambio calculo mopre
'Global Const UltimaModificacion1 = " "      ' Items de ganancias en 100
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.77"
'Global Const FechaModificacion = "07/08/2009"
'Global Const UltimaModificacion = "MB"  'Cambio en grossing cuando itera sale del liqpro06 sin que termine de liquidar
'Global Const UltimaModificacion1 = " "  'Formula Gcias en Impuestos y debitos Bancarios Cuando el item=55 y el valor de DDJJ es x 17% y no 34%
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.78"
'Global Const FechaModificacion = "18/08/2009"
'Global Const UltimaModificacion = "MB"  'Cambio en formula de chile de recalculo de impuesto unico y recalc de conceptos
'Global Const UltimaModificacion1 = " "  ' para que tome recalculos sobre periodos ya recalculados
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.79"
'Global Const FechaModificacion = "26/08/2009"
'Global Const UltimaModificacion = "MB"  'Cambio en formula de chile de recalculo de impuesto unico con mas log y performance
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.80"
'Global Const FechaModificacion = "01/09/2009"
'Global Const UltimaModificacion = "MB"  'Se cambio la constante MaxIteraGross = 20 para que aumente las iteracion de Grossing
'Global Const UltimaModificacion1 = " "  'que estaba en 10 y en Chile no alcanzaba
'Global Const UltimaModificacion2 = " "

Global Const Version = "3.81"
Global Const FechaModificacion = "03/09/2009"
Global Const UltimaModificacion = "Martin" 'DiasAnualesVac() - Se sumo a la diferencia de fechas
Global Const UltimaModificacion1 = " "     'Se agrego transaccion en liqpro04
Global Const UltimaModificacion2 = " "
