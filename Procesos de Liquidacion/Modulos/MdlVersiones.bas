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

'Global Const Version = "3.81"
'Global Const FechaModificacion = "03/09/2009"
'Global Const UltimaModificacion = "Martin" 'DiasAnualesVac() - Se sumo a la diferencia de fechas
'Global Const UltimaModificacion1 = " "     'Se agrego transaccion en liqpro04
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.82"
'Global Const FechaModificacion = "27/10/2009"
'Global Const UltimaModificacion = "Martin" 'Ganancias cambio la forma de buscar en ficharet
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.83"
'Global Const FechaModificacion = "09/11/2009"
'Global Const UltimaModificacion = "Martin" 'Ganancias cambio la forma de buscar en ficharet y correcciones de lipro06
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.84"
'Global Const FechaModificacion = "12/11/2009"
'Global Const UltimaModificacion = "Martin" 'Cambios en los topeos de las busquedas de embargos mensuales y quincenales
'Global Const UltimaModificacion1 = " "     'Para Quincenales tomar 1/2 SMVM
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.85"
'Global Const FechaModificacion = "19/11/2009"
'Global Const UltimaModificacion = "Martin" 'Permitir que los embargos por porcentaje o monto sean en distinta moneda
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.86"
'Global Const FechaModificacion = "24/11/2009"
'Global Const UltimaModificacion = "MB" 'Recalculo de IU agregado de Zona Extrema
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.87"
'Global Const FechaModificacion = "05/01/2010"
'Global Const UltimaModificacion = "MB" 'Error en tope de sac cuando tenia algo en el semestre y no superaba el tope
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.88"
'Global Const FechaModificacion = "11/01/2010"
'Global Const UltimaModificacion = "MB" 'Error en creacion de tabla wf_escalautm para Oracle por case sensitive.
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.89"
'Global Const FechaModificacion = "12/01/2010"
'Global Const UltimaModificacion = "MB" 'Solos se crea la tabla wf_escalautm si es un modelo de Chile (tipoproceso = 3).
'Global Const UltimaModificacion1 = " " 'Limpieza de formulas viejas de chile solo quedaron las que se usan realmente
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.90"
'Global Const FechaModificacion = "21/01/2010"
'Global Const UltimaModificacion = "MB" 'Error en imponible de SAC cuando tenia 2 liq en mes y desborde del imponible.
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.91"
'Global Const FechaModificacion = "26/01/2010"
'Global Const UltimaModificacion = "MB" 'Cambio en AMPO se busca por fecha fin del proceso y no por fecha de pago
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.92"
'Global Const FechaModificacion = "24/02/2010"
'Global Const UltimaModificacion = "MB" 'Cambio en formula de recalculo de IU las deducciones se convierten a UTM historico
'Global Const UltimaModificacion1 = " " 'y redondeo sin decimales a los imponibles
'Global Const UltimaModificacion2 = " " 'Nueva busqueda dias habiles de las licencias
                                       'Correccion de la busqueda 54 - Dias de ingreso

'Global Const Version = "3.93"
'Global Const FechaModificacion = "17/03/2010"
'Global Const UltimaModificacion = "MB" 'Cambio en formula de ganancias se agrego un parametro para indicar si prorratea o no
'Global Const UltimaModificacion1 = " " 'parametro 1006 - 0 prorratea y 1 no prorratea por mas que el acum/cpto este indicado como q prorratea
'Global Const UltimaModificacion2 = " "


'Global Const Version = "3.94"
'Global Const FechaModificacion = "30/03/2010"
'Global Const UltimaModificacion = "Martin"  ' Se termino busqueda de premio medicus, la busqueda de familiares con guarderia y  formula
'Global Const UltimaModificacion1 = " "      ' de premio mensual
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.95"
'Global Const FechaModificacion = "08/04/2010"
'Global Const UltimaModificacion = "MB"  ' Se cambio la formulas de Recalculo de Cptos e ImpUnico para que divida x los meses trabajados y no por la
'Global Const UltimaModificacion1 = " "      ' cantidad de periodos
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.96"
'Global Const FechaModificacion = "30/04/2010"
'Global Const UltimaModificacion = "Martin Ferraro" 'Correccion en la busqueda bus_DiasSAC_Proporcional_Tope30()
'Global Const UltimaModificacion1 = " "             'Cuando la fase termina ultimo dia del mes descontaba un dia
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.97"
'Global Const FechaModificacion = "04/05/2010"
'Global Const UltimaModificacion = "MB" 'Correccion topeo de SAC con monto de gratif en el semestre sin cantidad de dias
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.98"
'Global Const FechaModificacion = "13/05/2010"
'Global Const UltimaModificacion = "MB" 'Cambios en el liqpro04 agregue mas log para el recalculo de IU
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "3.99"
'Global Const FechaModificacion = "18/05/2010"
'Global Const UltimaModificacion = "MB" 'Cambios en retroactivos con novedades de ajuste porque no marcaba el detliq con retro
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.00"
'Global Const FechaModificacion = "27/05/2010"
'Global Const UltimaModificacion = "Martin" 'Se modifico bus_FamGuarderia() para que compare la fecha de vencimiento contra la hasta del periodo y no
'Global Const UltimaModificacion1 = " "     ' contra la de pago del proceso
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.01"
'Global Const FechaModificacion = "16/06/2010"
'Global Const UltimaModificacion = "Martin" '1 - Se agrego la funcion Top en dataAccess para usar en la funcion de reutilizacion de traza que tardaba mucho
'Global Const UltimaModificacion1 = " "     '2 - Modificacion de formula for_irpf_simple. Se agrego param opcional multiplicador que multiplica los topes de cada escala
'Global Const UltimaModificacion2 = " "     '    por el valor y se agrego traza para valores intermedios
                                           '3 - Nueva busqueda para anticipos Multivoice bus_AnticipoMV
                                           '4 - Modificaciones para poder manejar retroactividades con tablas temp (Tb CrearRetro)
                                           '5 - Nueva Version del Liqpro15
                                           '6 - Se agrego Busqueda bus_DiasHabLic Dias habiles en licencias
                                           '7 - Se agrego Busqueda bus_ConcAcumAnioMesFijo
                                           '8 - Se agrego Busqueda bus_ConcAcumAniVaroMesFijo
                                           '9 - Se modificaron todos los AM_Prom (menos AM_PromSin0_Nuevo y AM_Prom_Nuevo)porque sumaba uno mas cuando incluia proceso actual
                                           '10 - bus_DiasSAC_Proporcional_Tope30() - No consideraba el ultimo dia como trabajado
                                           
'Global Const Version = "4.02"
'Global Const FechaModificacion = "08/07/2010"
'Global Const UltimaModificacion = "Martin" '1 - bus_Licencias_Integradas() - Se agrego a la busqueda del periodo de gti la parte de alcance de estr
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "


'Global Const Version = "4.03"
'Global Const FechaModificacion = "20/07/2010"
'Global Const UltimaModificacion = "Martin" '1 - bus_Licencias_Integradas() se corrigio error de la version anterior
'Global Const UltimaModificacion1 = " "     '2 - Se agrego tope en la busqueda de prestamos
'Global Const UltimaModificacion2 = " "     '3 - Se quito escritura de log en la evaluacion de formulas (CEval)


'Global Const Version = "4.04"
'Global Const FechaModificacion = "29/07/2010"
'Global Const UltimaModificacion = "Martin" '1 - Liqpro15 - Se modifico para proporcionar las novedades retroactivas
'Global Const UltimaModificacion1 = " "     '2 - bus_DiasSAC_Proporcional_Tope30() - Error cuando calculaba en febrero - Cuando entraba el 15/02 descontaba 16 dias en vez de 14
'Global Const UltimaModificacion2 = " "     '3 - bus_DiasSAC_Proporcional() - Control de If Fec_Fin_Calc < Fec_Ini_Calc Then Fec_Fin_Calc = Fec_Fin_Sem

'Global Const Version = "4.05"
'Global Const FechaModificacion = "04/08/2010"
'Global Const UltimaModificacion = "Martin" '1 - bus_Licencias_Integradas() - Se creo variable DiasTopeMes guarda el tope de 30 o 15 dias.
'Global Const UltimaModificacion1 = " "     '    Si la cantidad de dias del mes es <= lo considera quincenal y el tope es de 15, sino 30
'Global Const UltimaModificacion2 = " "     '2 - La parte de desmarcar prestamos se paso a un sub con manejador de errores

'Global Const Version = "4.06"
'Global Const FechaModificacion = "10/08/2010"
'Global Const UltimaModificacion = "MB" ' Cambio en for_ganancias para que guarde la cuota prorrateable en la traza y en traza_gan_item_top
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.07"
'Global Const FechaModificacion = "24/08/2010"
'Global Const UltimaModificacion = "MB" ' Cambio en bus_Acum_Con_Ajuste, AM_Max_Aju, AM_Prom_Aju, AM_promsin0_Aju AM_Min_Aju
'Global Const UltimaModificacion1 = " " ' Para el CAS-11070-MONRESA-REMUNERACIONES VARIABLES AJ X INFLACION
'Global Const UltimaModificacion2 = " " ' Se modifico la formula de gcias para que no de error cuando hay DDJJ de item 3 con Cuit

'Global Const Version = "4.08"
'Global Const FechaModificacion = "23/09/2010"
'Global Const UltimaModificacion = "MB" ' Nueva busqueda bus_Feriados() para feriados no laborable y laborables para Monresa
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "4.09"
'Global Const FechaModificacion = "07/10/2010"
'Global Const UltimaModificacion = "MB" ' Cambio en la busqueda de sac porque daba 1 dia menos
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "4.10"
'Global Const FechaModificacion = "19/10/2010"
'Global Const UltimaModificacion = "MB" ' Cambio en la busqueda bus_Acum_Con_Ajuste, nueva busqueda de desgloce de horas gti bus_Desg_Horas
'Global Const UltimaModificacion1 = " " ' cambio en bus_Fecha que no andaba si no se liquidaba con debug
'Global Const UltimaModificacion2 = " " ' nueva busqueda bus_Acum3_Con_Ajuste ajuste para acum fijos CAS-11070-MONRESA-REMUNERACIONES VARIABLES AJ X INFLACION

'Global Const Version = "4.11"
'Global Const FechaModificacion = "18/11/2010"
'Global Const UltimaModificacion = "MB" ' Nueva busqueda Licencias del mes siguiente para Teletech
'Global Const UltimaModificacion1 = " " '
'Global Const UltimaModificacion2 = " " '

'Global Const Version = "4.12"
'Global Const FechaModificacion = "01/12/2010"
'Global Const UltimaModificacion = "MB" ' Cambio en bus_Vac_No_Gozadas_Pendientes para que no redondee
'Global Const UltimaModificacion1 = " " ' se agrego un parametro que indica si redondea o no Arr_Programa(NroProg).Auxlog1
'Global Const UltimaModificacion2 = " " ' Nueva busqueda Bus_NovRetro para el Simulador agregado de retroactividad a cada periodo
                                       ' For_gcias Cuando hay dif de plan item13 y devuelve tiene que restar el valor liquidado por eso saco el ABS de LIQ en tipo tope 5
 
'Global Const Version = "4.13"
'Global Const FechaModificacion = "04/01/2011"
'Global Const UltimaModificacion = "MB" ' Cambio en bus_Desg_Horas para Monresa para que traiga la dif de valor de hs de categ del desglose vs la cat actual
'Global Const UltimaModificacion1 = " " ' Cambio en formula de gcias no guardaba las ddjj de item1 de otras empresas para mostrar en F649
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.14"
'Global Const FechaModificacion = "07/01/2011"
'Global Const UltimaModificacion = "MB" ' Nueva Busqueda Periodos de Recalculo de IU bus_periodos_recalc
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.15"
'Global Const FechaModificacion = "14/01/2011"
'Global Const UltimaModificacion = "MB" ' Modif Busquedas de Ajuste de Monresa y de Desglose de Horas 2
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.16"
'Global Const FechaModificacion = "20/01/2011"
'Global Const UltimaModificacion = "MB" ' Nueva Busqueda bus_present_licenfer para Sidersa
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.17"
'Global Const FechaModificacion = "10/02/2011"
'Global Const UltimaModificacion = "FGZ" ' Busqueda bus_novretro (le agregué la marca de la novedad)
'Global Const UltimaModificacion1 = " "  ' liqpro04 (desmarcar las novedades retroactivas)
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.18"
'Global Const FechaModificacion = "11/02/2011"
'Global Const UltimaModificacion = "FGZ" ' Busqueda bus_DiasSAC_Proporcional_Tope30() - Error cuando calculaba en febrero - Cuando entraba antes y cuand tenia una baja en el mes
'Global Const UltimaModificacion1 = " "  ' Ademas cuando la baja era un dia 31 no estaba topeando bien
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "4.19"
'Global Const FechaModificacion = "17/02/2011"
'Global Const UltimaModificacion = "MB" ' Cambio en formula de gcias no tomaba las ddjj de item2 del mes
'Global Const UltimaModificacion1 = " "  'Cambio en formula de recalculo de IU de Chile para cuando no trabajo un periodo
'Global Const UltimaModificacion2 = " "  'nueva busqueda de hs de licencias bus_Licencias_parciales para Fundacion San Cayetano

'Global Const Version = "4.20"
'Global Const FechaModificacion = "18/02/2011"
'Global Const UltimaModificacion = "MB" ' Cambio en la busqueda de licencias integradas en febrero daba un dia de mas
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.21"
'Global Const FechaModificacion = "05/04/2011"
'Global Const UltimaModificacion = "MB" 'busqueda de hs de licencias bus_Licencias_parciales para FSC que marque la lic para no la tome otra busqueda
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.22"
'Global Const FechaModificacion = "25/04/2011"
'Global Const UltimaModificacion = "FGZ" ' Busqueda de conceptos meses fijos (estaba mal seteada una variable cuando utilizaba la opcion de incluir periodo y proceso actual)
'Global Const UltimaModificacion1 = " "  'Busqueda bus_Desg_Horas    Custom Monresa
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.23"
'Global Const FechaModificacion = "27/04/2011"
'Global Const UltimaModificacion = "FGZ" 'Busqueda bus_DiasSAC_Proporcional_Tope30()
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.24"
'Global Const FechaModificacion = "09/05/2011"
'Global Const UltimaModificacion = "FGZ" 'Nuevo tipo de Busqueda bus_Francos_noGozados()
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "4.25"
'Global Const FechaModificacion = "18/05/2011"
'Global Const UltimaModificacion = "MB" 'Recalculo de CHile Se saco el ajuste por UTM de cada cpto reliquidado de deducciones historico
'Global Const UltimaModificacion1 = " "  'Cambio en bus_Francos_noGozados de Winsy
'Global Const UltimaModificacion2 = " "

'===========================================================
'Global Const Version = "5.00"
'Global Const FechaModificacion = "26/05/2011"
'Global Const UltimaModificacion = "FGZ" 'Mejoras de Performance
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "5.01"
'Global Const FechaModificacion = "10/06/2011"
'Global Const UltimaModificacion = "FGZ" 'Mejoras de Performance
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "5.02"
'Global Const FechaModificacion = "22/06/2011"
'Global Const UltimaModificacion = "FGZ" 'Cambio en bus_Francos_noGozados
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "

'Global Const Version = "5.03"
'Global Const FechaModificacion = "27/06/2011"
'Global Const UltimaModificacion = "FGZ" 'Cambio en 102 bus_DiasHabLic: Dias habiles en licencias
'Global Const UltimaModificacion1 = " "  ' Se le agregó control de dias feriados
'Global Const UltimaModificacion2 = " "

'Global Const Version = "5.04"
'Global Const FechaModificacion = "30/06/2011"
'Global Const UltimaModificacion = "FGZ" 'Cambio en la formula for_RecalcImpuestoUnico
'Global Const UltimaModificacion1 = " "  ' Habia mal una referencia a un campo
'Global Const UltimaModificacion2 = " "  '       Estaba difimponible y es difimponibleact

'Global Const Version = "5.05"
'Global Const FechaModificacion = "08/07/2011"
'Global Const UltimaModificacion = "FGZ" 'Cambio en busqueda de embargos bus_embargos
'Global Const UltimaModificacion1 = " "  ' Se le agregó logica para poder calcular el SMVM proporcional a la carga horaria
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "5.06"
'Global Const FechaModificacion = "13/07/2011"
'Global Const UltimaModificacion = "FGZ" 'Cambio en busqueda de embargos bus_ParametroConcepto
'Global Const UltimaModificacion1 = " "  '
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "5.07"
'Global Const FechaModificacion = "21/07/2011"
'Global Const UltimaModificacion = "FGZ" 'Cambio en busqueda de embargos bus_embargos
'Global Const UltimaModificacion1 = " "  '   el SMVM proporcional proporciona mensualmente
'Global Const UltimaModificacion2 = " "  ' Se cambió la formula de Chile for_RecalcConcepto
''                                               inserto pero pongo en 0 los campos difimponible y rentaimpoact que utiliza el recalculo de impuesto único
''                                         Se modificó la busqueda de noveedades Bus_NovGegi
''                                               Habia un problema cuando se guardan novedades historicas para las novedades globales y cuando se usa cache

'Global Const Version = "5.08"
'Global Const FechaModificacion = "28/07/2011"
'Global Const UltimaModificacion = "FGZ" 'Mejora en el control de tiempo transcurrido
'Global Const UltimaModificacion1 = " "  '   Cuando solo desliquida no actualiza el tiempo acumulado y luego el tiempo transcurrido da negativo
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "5.09"
'Global Const FechaModificacion = "15/09/2011"
'Global Const UltimaModificacion = "FGZ" 'Se cambió la formula de Chile for_RecalcImpuestoUnico
'Global Const UltimaModificacion1 = " "  '   Se cambió este calculo de zona extrema
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "5.10"
'Global Const FechaModificacion = "21/09/2011"
'Global Const UltimaModificacion = "FGZ" 'bus_Licencias_Integradas() -
'Global Const UltimaModificacion1 = " "  '   Problemas para determinar el periodo de GTI cuando hay alcance por estructura
'Global Const UltimaModificacion2 = " "  '

'Global Const Version = "5.11"
'Global Const FechaModificacion = "04/10/2011"
'Global Const UltimaModificacion = "FGZ" 'Cambio en busqueda de embargos bus_embargos
'Global Const UltimaModificacion1 = " "  '   para que marque los embargos con el nro de proceso de liquidacion y el estado anterior
'Global Const UltimaModificacion2 = " "  '   cuando se activa un prestamo que estaba en estado de Espera
''                                         Ademas se modificó es desmarcado de embargo para restaure el estado anterior del embargo cuando el mismo fué activado automaticamente por una liquidacion


'Global Const Version = "5.12"
'Global Const FechaModificacion = "08/11/2011"
'Global Const UltimaModificacion = "FGZ" '
'Global Const UltimaModificacion1 = " "  '
'Global Const UltimaModificacion2 = " "  ' Error en desmarcado de prestamo. Estaba mal referenciado un campo.

'Global Const Version = "5.13"
'Global Const FechaModificacion = "16/11/2011"
'Global Const UltimaModificacion = "FGZ" '
'Global Const UltimaModificacion1 = " "  ' Nueva Busqueda de Curvas Estacionales
'Global Const UltimaModificacion2 = " "  ' busqueda bus_DiasHabiles_Trabajados. Mejoras de performance

'Global Const Version = "5.14"
'Global Const FechaModificacion = "21/12/2011"
'Global Const UltimaModificacion = "FGZ" '
'Global Const UltimaModificacion1 = " "  'busqueda bus_Prestamos. Mejoras de performance
'Global Const UltimaModificacion2 = " "  'busqueda bus_vales. Mejoras de performance

'Global Const Version = "5.15"
'Global Const FechaModificacion = "13/02/2012"
'Global Const UltimaModificacion = "FGZ" '
'Global Const UltimaModificacion1 = " "  'Nuevo tipo de Busqueda
'Global Const UltimaModificacion2 = " "  '       115  - busqueda bus_AcumparaSAC_RyNR. Busqueda de Acumuladores de meses fijos para sac (Rem y No Rem)

'Global Const Version = "5.16"
'Global Const FechaModificacion = "19/04/2012"
'Global Const UltimaModificacion = "JAZ" ' Juan Zamarbide
'Global Const UltimaModificacion1 = " "  ' CAS-14735 - H&A - Error Busqueda Segun Fases
'Global Const UltimaModificacion2 = " "  ' Se cambió la lógica de la búsqueda por que no cumplía con todos los criterios de la misma. Se renombró a la vieja búsqueda como bus_DiasEnMesSegunFase_OLD.

'Global Const Version = "5.17"
'Global Const FechaModificacion = "23/04/2012"
'Global Const UltimaModificacion = "JAZ" ' Juan Zamarbide
'Global Const UltimaModificacion1 = " "  ' CAS-14735 - H&A - Error Busqueda Segun Fases
'Global Const UltimaModificacion2 = " " ' Corrección de error detectado en formulario de Rechazo

'Global Const Version = "5.18"
'Global Const FechaModificacion = "03/05/2012"
'Global Const UltimaModificacion = "JAZ" ' Juan Zamarbide
'Global Const UltimaModificacion1 = " "  ' CAS-14735 - H&A - Error Busqueda Segun Fases
'Global Const UltimaModificacion2 = " " ' Corrección de error detectado en formulario de Rechazo - Faltaban contemplar casos en la Sql de la búsqueda

'Global Const Version = "5.19"
'Global Const FechaModificacion = "14/05/2012"
'Global Const UltimaModificacion = "Lisandro Moro"
'Global Const UltimaModificacion1 = " Nivelacion con simulador - CAS-13713 - MONRESA - Gestion Presupuestaria - Simulaciones multiples "
'Global Const UltimaModificacion2 = " Correccion bus_Concep3, faltaban condiciones al sql - empleado y concepto "

'Global Const Version = "5.20"
'Global Const FechaModificacion = "27/06/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " Mejoras de performance."
'Global Const UltimaModificacion2 = " "
''se hicieron varias mejoras para tratar de manejar problemas de interloqueos
''                               cuando se porcesan simultaneamente varios procesos de liquidacion
''                               Modo transaccional con nivel de aislamiento Read Committed Snapshot (Snapshot)
''                                                              sub Batliq06.
''                                                              sub Liqpro16.
''                                                              sub Liqpro04. Desmarcado de Embargos, Vales, Licencias y Pagos/Dtos de vacaciones
''                                                              sub Liqpro06.
''                                                              Busqueda Licencias integradas (82): update with rowlock
''                                                              Busquedas en general
''                                                              Formulas en general + for_grossnew
'''                                         Se modificó la busqueda de noveedades Bus_NovGegi
'''                                               Habia un problema cuando se guardan novedades historicas para las novedades globales y cuando se usa cache


'Global Const Version = "5.20"
'Global Const FechaModificacion = "30/07/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " Mejoras de performance."
'Global Const UltimaModificacion2 = " "
''se hicieron varias mejoras para tratar de manejar problemas de interloqueos
''                               cuando se porcesan simultaneamente varios procesos de liquidacion
''                               Modo transaccional con nivel de aislamiento Read Committed Snapshot (Snapshot)
''                                                              sub Batliq06.
''                                                              sub Liqpro16.
''                                                              sub Liqpro04. Desmarcado de Embargos, Vales, Licencias y Pagos/Dtos de vacaciones
''                                                              sub Liqpro06.
''                                                              Busqueda Licencias integradas (82): update with rowlock
''                                                              Busquedas en general
''                                                              Formulas en general + for_grossnew
''                                         Se modificó la busqueda de noveedades Bus_NovGegi
''                                               Habia un problema cuando se guardan novedades historicas para las novedades globales y cuando se usa cache
''                                         Se modificó la busqueda de prestamos
''                                         Se modificó las formulas de Uruguay for_irpf y for_irpf_diciembre
''



'Global Const Version = "5.21"
'Global Const FechaModificacion = "18/09/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   busqueda  bus_Licencias_parciales()
''                                      Controla la hora desde y hasta solo si la lic es parcial fija
''                                   'CAS-16863 - Sykes - Nuevo tipo de Busqueda 116 - Incapacidades
''                                   Busqueda de novedades retroactivas (bus_novretro)
''                                           marco las lic_pagas que originaron una nov retroactiva


'Global Const Version = "5.22"
'Global Const FechaModificacion = "28/09/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS- 16688- SANTANDER CHILE- ERROR EN RELIQUIDACION
''                                   Se modificó la formula interna de Recalculo de impuesto Unico (Chile) - for_RecalcImpuestoUnico
''                                           El valor que  se esta guardando en la tabla Impuni_cab.rentaimpopact es el valor de Tributable menos Zona Extrema,
''                                           El valor que debería ser guardado es Nuevo Tributable.
''                                   Se modificó la busqueda bus_ValorEnOtroLegajo. Habia una referencia a un campo que no estaba en el select


'Global Const Version = "5.23"
'Global Const FechaModificacion = "04/10/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS-16833 - Vision Outsourcers - Chacomer Py - Interfaz Comisiones Produccion
''                                   Se creó la formula interna de para el calculo de comisiones for_comision
'

'Global Const Version = "5.24"
'Global Const FechaModificacion = "24/10/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   'CAS-16863 - Sykes - Tipo de Busqueda 116 - Incapacidades
''                                   Habia un problema con el registro de la fecha del dia pago cuando la licencia comienza fuera de las fechas del proceso

'Global Const Version = "5.25"
'Global Const FechaModificacion = "29/10/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   'CAS-16863 - Sykes - Tipo de Busqueda 116 - Incapacidades
''                                   Habia un problema con el registro de la fecha del dia pago cuando la licencia comienza fuera de las fechas del proceso

'Global Const Version = "5.26"
'Global Const FechaModificacion = "05/11/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   'CAS-16863 - Sykes - Tipo de Busqueda 116 - Incapacidades
''                                   Correccion a la forma de buscar licencias anteriores.

'Global Const Version = "5.27"
'Global Const FechaModificacion = "07/11/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   'CAS-16863 - Sykes - Tipo de Busqueda 116 - Incapacidades
''   FGZ - 07/11/2012 - se agregó control de dia laborable en las fechas de la licencia.



'Global Const Version = "5.28"
'Global Const FechaModificacion = "16/11/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   'CAS-17605 - ANDREANI - Bug en Liquidador - Tipo de Busqueda 50 - bus_Vac_No_Gozadas
''                                   'Habia un campo mal referenciado. lic_vacacion.vacanio , el campo esta en otra tabla vacacion.vacanio

'Global Const Version = "5.29"
'Global Const FechaModificacion = "23/11/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS - 17424 -  GRUPO CARGO - Upgrade a  R3 de RH Pro X2 - QA
''                                   Tipo de Busqueda 90 - bus_Licencias_Horas
''                                   'Modulo Varios.bas Estaba usando una funcion  CantidadDeHorasDias() en la cual estaba mal definido el objeto feriado.
''                                   'Modulo Varios.bas Estaba usando una funcion  CantidadDeDiasHab() en la cual estaba mal definido el objeto feriado.


'Global Const Version = "5.30"
'Global Const FechaModificacion = "30/11/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS-17697 - Sykes - Simulador - Error Retroactivos
''                                   Nivelacion de Version. Para mantener las mismas versiones en Liquidador y
''                                   todos los procesos simulación se registra la versión aunque no hay modificaciones de código.


'Global Const Version = "5.31"
'Global Const FechaModificacion = "18/12/2012"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS-17697 - Sykes - Tipo de Busqueda 116 - Incapacidades
''                                   Habia un ordenamiento incorrecto de las licencias cuando controlaba lic no pagas

'==============================================
'Aun no liberad oficialmente (se generó para salir de un apuro con SEDAMIL y Santander Chile)
'Global Const Version = "5.32"
'Global Const FechaModificacion = "28/12/2012"

'Global Const Version = "5.32"
'Global Const FechaModificacion = "08/01/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS-17560 - RH Pro Consulting - Santander Santiago - Error Recalculo Impuesto Unico
''                                   Se modificó la formula interna de Recalculo de impuesto Unico (Chile) - for_RecalcImpuestoUnico
''   Ademas
''                                   'Bug en Liquidador - Tipo de Busqueda 59 - bus_Vac_No_Gozadas_A_Pagar
''                                   'Habia un campo mal referenciado. lic_vacacion.vacanio , el campo esta en otra tabla vacacion.vacanio
''   Ademas
''                                   CAS-18029 - GC - Cardiff - Nuevo item de Ganancias
''                                   FOR_GANANCIAS
''                                       Impuestos y debitos Bancarios
''                                           Se agregó el ITEM 23 Impuesto Deb y Creditos (100%)

'Global Const Version = "5.33"
'Global Const FechaModificacion = "14/01/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS-17686 - CDA – Mejoras para la opción de carga de Gastos
''                                   Nuevo tipo de Busqueda 117 - Gastos
''
''                                   LIQPRO04    Se agregó call Desmarcar_gastos
''                                   Se creó sub para desmarcar gastos marcados por una liquidacion
''   Ademas
''                                   FOR_GANANCIAS
''                                       Impuestos y debitos Bancarios
''                                           Se cambió el ITEM 23 por el item 56

'Global Const Version = "5.34"
'Global Const FechaModificacion = "13/02/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS-17686 - CDA – Mejoras para la opción de carga de Gastos
''                                   Tipo de Busqueda 117 - Gastos
''                                   Error en la busqueda de gastos ya pagados
''   Ademas
''                                   for_RecalcImpuestoUnico
''                                       Se agregó el calculo del valor menor de MontoZonaExt1 y MontoZonaExt2

'Global Const Version = "5.35"
'Global Const FechaModificacion = "26/02/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS-16441 - H&A - Perú - Distribucion e utilidades
''                                   Nuevo Tipo de Busqueda 118 - Utilidades
''                                   modificacion en sub Establecer_Proceso del modeulo mdlbuliq

'Global Const Version = "5.36"
'Global Const FechaModificacion = "04/03/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS-18103 - SYKES COSTA RICA - Registro de Cancelacion de Saldo Vac. Por Liquidacion
''                                   Nuevo Tipo de Busqueda 119 - bus_SaldoVac_CR()


'Global Const Version = "5.37"
'Global Const FechaModificacion = "15/03/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   CAS-16441 - H&A - Perú - Distribucion e utilidades
''                                   Tipo de Busqueda 118 - Utilidades
''                                       Se le agregaron las opciones de resultado para CargasdeFlia

' **************
' Version beta liberada exclusivamente para Chile por Recalculo de impuesto Unico
'Global Const Version = "5.38"
'Global Const FechaModificacion = "03/04/2013"
'Global Const UltimaModificacion = "FGZ"
' **************


'Global Const Version = "5.38"
'Global Const FechaModificacion = "07/06/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                                   Por caso CAS-18643 - CDA - Custom Legal - Formula de Embargo
''                                       Tipo de Busqueda 22: Embargos
''                                       Cuando el embargo es de tipo 'Embargo Judicial' No tiene en cuenta otras liquidaciones en el mes.
''
''Ortos cambios --------------------------------------
''                                   CAS-16441 - H&A - Perú - CTS
''                                       Nuevo Tipo de Busqueda 120 - CTS - Tiempo Efectivamente Trabajado -
''                                           bus_CTS_Tiempo_Trabajado()
''                                       Nuevo Tipo de Busqueda 121 - CTS - Acum. Mensual meses fijos -
''                                           bus_Acum_CTS()
''                                       Nuevo Tipo de Busqueda 122 - CTS - Conceptos Mese Fijos -
''                                           bus_Concep_CTS()
''                                       Nuevo Tipo de Busqueda 123 - Acum. Mensual meses fijos Desde Hasta
''                                           bus_Acum_FijosDesdeHasta()
'' Ademas
''                                   Por caso CAS-18895 - RH Pro Consulting - Santander - Chile
''                                       Se modificó la formula interna de Recalculo de impuesto Unico (Chile) - for_RecalcImpuestoUnico
''                                           Se cambió el calculo de la zona Extrema2 para el calculo del nuevo impuesto
'' Otras
''                                   Por caso CAS-19684 - HORWATH LITORAL - AMR -  Error busqueda de antiguedad
''                                       Modificacion Tipo de Busqueda 97: Antiguedad Nueva. sub bus_Anti_Nueva().
''                                           Se mofificó el sub bus_Antiguedad2, cuando habia una sola fase no estaba topeando a 30 dias cuando el resultado no era en dias,
''                                           en lugar de dar un mes mas, daba resultados como 11 meses y 30 dias. Debiera dar 12 mes o 1 año
'' Otras
''                                   Por caso CAS-19679 - GESTION COMPARTIDA - Error en liquidador item 13 ganancias
''                                       Modificacion en formula for_ganancias.
''                                       Para los items de tipo 5,como el item 13, se restauró el ABS de LIQ que se habia sacado en la V4.12 en 01/12/2010.


'Global Const Version = "5.39"
'Global Const FechaModificacion = "26/06/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                               CAS-16441 - H&A - PERU - BUSQUEDA PARA CTS
''                                       Tipo de Busqueda 120 - CTS - Tiempo Efectivamente Trabajado - bus_CTS_Tiempo_Trabajado()
''                                           Estaba armando mal las fechas del periodo en los meses de Noviembre y Diciembre
''
''Otros cambios
''                               Cuando un parametro se resuelve por novedad no está poniendo la descripcion correcta en la Traza
''

'Global Const Version = "5.40"
'Global Const FechaModificacion = "11/07/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                               CAS-18773 - CDA - Customs Varias
''                                   Tipo de Busqueda 117 - Gastos.
''                                   Se le agregó la generacion de retroactividad para los gastos. Funcionan de la misma  forma que las novedades retroactivas.
''Ademas
''                                   Por caso CAS-19684 - HORWATH LITORAL - AMR -  Error busqueda de antiguedad
''                                       Modificacion Tipo de Busqueda 97: Antiguedad Nueva. sub bus_Antiguedad2().
''                                       Se analiza si las fases cortadas en realidad son continuas y el resultado es en años ==> se calcula a año completo recien al mismo dia del siguiente año
''                                           Ejemplo del 01/01/2013 al 31/12/2013 NO hay un año recien cumple el año el 01/01/2014


'Global Const Version = "5.41"
'Global Const FechaModificacion = "05/08/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                               CAS-18722 - HORWATH LITORAL - AMR - Busqueda de Vacaciones
''                                       Nuevo Tipo de Busqueda 124 - Dias Corresp - Control Baja
''                                           bus_DiasVac_RecPorBaja()
''Ademas ....
''                               Estaba mal el comentario de la validacion de la estructura de la BD, los campos agregados no son esos sino gasretro, pliqdesde, pliqhasta en la tabla gastos
''
''                               En la seccion final de proceso, se imprimirá en el log siempre las estadisticas de cantidad de tiempo de procesamiento y tiempo por empleado.
''                               independientemente de si se usa beddug o no

 
'Global Const Version = "5.42"
'Global Const FechaModificacion = "09/08/2013"
'Global Const UltimaModificacion = "FGZ"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                               CAS-20769 - SYKES - Error calculo de embargos
''                                       Tipo de Busqueda 22: Embargos
''                                       Producto del ultimo cambio legal en la V5.38 quedó mal un calculo cuando el embargo NO es de ley.
''Ademas...
''                               Tipo de Busqueda 119: Saldo Vacaciones con Venta (CR). bus_SaldoVac_CR()
''                               Se le agregó detalle de log cuando inserta en la tabla vacvendidos

'Global Const Version = "5.43"
'Global Const FechaModificacion = "12/08/2013"
'Global Const UltimaModificacion = "EAM"
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                               Tipo de Busqueda 119: Saldo Vacaciones con Venta (CR). bus_SaldoVac_CR()
''                               Se modifico la sql que busca licencia. Ahora busca todas las licencias del empleado y no a una fecha


'Global Const Version = "5.44"
'Global Const FechaModificacion = "02/09/2013"
'Global Const UltimaModificacion = " " 'FGZ - EAM (se envía con el caso - CAS-18103 - SYKES COSTA RICA - Registro de Cancelacion de Saldo Vac. Por Liquidacion [Entrega 3])
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                               Tipo de Busqueda 22: Busqueda de Embargos. Se restauró temporalmente la version anterior a Junio 2013 por problemas
''                               Se agrego la función Desmarcar_VentaVac() para de desliquidacion en la funcion liqpro04.
''                               Se modificó la busqueda de licencia en venta de vacaciones por días habiles y no por el campo cantdias.'
''                               La formula nueva de ganancias FOR_GANANCIAS2013

'Global Const Version = "5.45"
'Global Const FechaModificacion = "18/09/2013"
'Global Const UltimaModificacion = " " 'CAS-21112 - H&A - LIQ - Cambio legal Ganancias Argentina
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                               Se modifico la formular nuevamente de ganancia FOR_GANANCIAS2013
''                                      Se modifico el formula ganancia el calculo del item 17
''                                      Se corrigieron errores en las funciones AM_Min, del modulo de tipo de busqueda. Falto declarar la variable ammes en la sql.
''                               Tipo de Busqueda 22: Busqueda de Embargos. Para quincenales se restauró la version anterior a Junio 2013 por problemas.
''                               Tipo de Busqueda 119: Saldo Vacaciones con Venta (CR). bus_SaldoVac_CR. Se agregó el control sobre los dias de beneficio.


'Global Const Version = "5.45B"
'Global Const FechaModificacion = "20/09/2013"
'Global Const UltimaModificacion = " " 'CAS-21118 - H&A - LIQ - Cambio legal Ganancias Argentina
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''                               Se modifico la formular nuevamente de ganancia FOR_GANANCIAS2013
''                                       Habia quedado un problema en la liquidacion de Agosto con mes de retencion en septiembre.
''                               Se agrego en la funciona de saldo de vacaciones para que busque días hábiles.


'Global Const Version = "5.46"
'Global Const FechaModificacion = "23/09/2013"
'Global Const UltimaModificacion = " " 'CAS-16441 - H&A - NACIONALIZACION PERÚ - BUSQUEDA PROMEDIO
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "
''
''Ademas
''                               Tipo de Busqueda 22: Busqueda de Embargos. Se reformularon las modificaciones del cambio legal de la version 5.38. Tanto para quincenales como mensuales.
''
''                               Se modifico la formular nuevamente de ganancia FOR_GANANCIAS2013
''                                       Habia quedado un problema en la liquidacion de Agosto con mes de retencion en septiembre.
''
''                                       Modificacion Tipo de Busqueda 97: Antiguedad Nueva. sub bus_Anti_Nueva().
''                                           Se mofificó el sub bus_Antiguedad2, estaba calculando mal a fin de año
''
''Ademas
''                                        Modificacion Tipo de Busqueda 7:   Acum. Mensual Meses Fijos
''                                           Se le agregó un parametro opcional de "Cantidad Minima de Acums". Funcional solo para operaciones. Suma, promedio y Promedio sin 0.
''                                        Modificacion Tipo de Busqueda 8:   Acum.Mens.Meses Variables
''                                           Se le agregó un parametro opcional de "Cantidad Minima de Acums". Funcional solo para operaciones. Suma, promedio y Promedio sin 0.
''                                        Modificacion Tipo de Busqueda 77:  Acum Mensual Fijos (Glencore)
''                                           Se le agregó un parametro opcional de "Cantidad Minima de Acums". Funcional solo para operaciones. Suma, promedio y Promedio sin 0.
''                                        Modificacion Tipo de Busqueda 79:  Acum Meses Var (Glencore)
''                                           Se le agregó un parametro opcional de "Cantidad Minima de Acums". Funcional solo para operaciones. Suma, promedio y Promedio sin 0.
''                                        Modificacion Tipo de Busqueda 86:  Acum meses variable con ajuste
''                                           Se le agregó un parametro opcional de "Cantidad Minima de Acums". Funcional solo para operaciones. Suma, promedio y Promedio sin 0.
''                                        Modificacion Tipo de Busqueda 107: Acum Mensual Mes Fijo Ajustado
''                                           Se le agregó un parametro opcional de "Cantidad Minima de Acums". Funcional solo para operaciones. Suma, promedio y Promedio sin 0.
''                                        Modificacion Tipo de Busqueda 121: CTS - Acumuladores Meses Fijos
''                                           Se le agregó un parametro opcional de "Cantidad Minima de Acums". Funcional solo para operaciones. Suma, promedio y Promedio sin 0.
''                                        Modificacion Tipo de Busqueda 123: Acumuladores Meses Desde-Hasta
''                                           Se le agregó un parametro opcional de "Cantidad Minima de Acums". Funcional solo para operaciones. Suma, promedio y Promedio sin 0.
'

'Global Const Version = "5.47"
'Global Const FechaModificacion = "02/10/2013"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'CAS-21616 - NGA - DDEE en F649 luego de cambio de ganancias
''                               Se modifico la formular nuevamente de ganancia FOR_GANANCIAS2013
''                                       Habia quedado un problema grabando traza y traza_gan cuando el bruto es menor a 15000 y no corresponde retencion.
''                                          Esto impactaba en F649.

'Global Const Version = "5.48"
'Global Const FechaModificacion = "08/10/2013"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'CAS-21616 - NGA - DDEE en F649 luego de cambio de ganancias
''                               Se modifico la formula nuevamente de ganancia FOR_GANANCIAS2013
''                                       Los item 29, 55 y 56 se tenian en cuenta en liq finales y diciembre pero en caso de que no sea final ni diciembre NO se deben tener en cuenta en el calculo.


'Global Const Version = "5.49"
'Global Const FechaModificacion = "17/10/2013"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'CAS-21616 - NGA - DDEE en F649 luego de cambio de ganancias
''                               Nueva correccion a la formula de ganancia FOR_GANANCIAS2013
''                                       Los item 29, 55 y 56 se tenian en cuenta en liq finales y diciembre pero en caso de que no sea final ni diciembre NO se deben tener en cuenta en el calculo.
''                                       Se topea contra el impuesto anual por escala y no contra lo que se retiene en el mes


'Global Const Version = "5.50"
'Global Const FechaModificacion = "22/10/2013"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'CAS-21979 - PWC CL - Modificacion calculo reliquidacion con adicional de salud
''                               Se modifico la formula de Chile for_RecalcConcepto()


'Global Const Version = "5.51"
'Global Const FechaModificacion = "21/11/2013"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'Caso Nro    CAS-21616 - NGA - F649 DDEE 1 er segmento
''                       Nueva correccion a la formula de ganancia FOR_GANANCIAS2013
''                          Aquellos que perciban menos de $15.000.- de Enero a Agosto, la Deducción Especial (ITEM 16) se calcule de acuerdo a lo
''                           establecido en el Dec.1242/13 Art.1, que dice lo siguiente:
''                               Artículo 1° - lncreméntase, respecto de las rentas mencionadas en los incisos a), b) y c) del artículo 79 de la Ley de Impuesto a las Ganancias,
''                               texto ordenado en 1997, y sus modificaciones, la deducción especial establecida en el inciso c) del artículo 23 de dicha Ley,
''                               hasta un monto equivalente al que surja de restar a la ganancia neta sujeta a impuesto las deducciones de los incisos a) y b) del mencionado artículo 23.
''
''
''                       Ademas se le agregaron opciones para recoleccion de datos estadisticos de ejecucion del liquidador (Campos y tabla nueva)
''

'Global Const Version = "5.52"
'Global Const FechaModificacion = "03/12/2013"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'Caso Nro    CAS-21616 - NGA - F649 DDEE 1 er segmento
''                       Nueva correccion a la formula de ganancia FOR_GANANCIAS2013
''                               No se estaba guardando el detalle de la traza del item
'

'Global Const Version = "5.53"
'Global Const FechaModificacion = "13/12/2013"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'CAS-22801 - SIDERSA - Error en Liqudiación de un Empleado
''
''                   Tipo de Busqueda 111: Presentismo Licencia por Enfermedad - Sidersa. bus_present_licenfer.  Se corrigieron varias referencias incorrectas a un campo de emp_lic (referenciaba emplicnro cuando el campo se llama emp_licnro)


'Global Const Version = "5.54"
'Global Const FechaModificacion = "02/01/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'CAS-22983 - Bodegas Norton - LIQ - Calculo Imp a las Ganancias
''
''                       Nueva correccion a la formula de ganancia FOR_GANANCIAS2013
''                               Problemas con el ítem 10, 11,12, 16 y 17 del Imp. a las Ganancias cuando la fecha de pago es Enero 2014 (liquidando Diciembre 2013 o posterior).
''

'Global Const Version = "5.55"
'Global Const FechaModificacion = "23/01/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'Varias modificaciones
''
''CAS-21979 - PWC CL - Modificacion calculo reliquidacion con adicional de salud
''                               Se modifico la formula de Chile for_RecalcConcepto()
''
''   Ademas
''CAS-22983 - Bodegas Norton - LIQ - Calculo Imp a las Ganancias
''                       Nueva correccion a la formula de ganancia FOR_GANANCIAS2013
''                               Problemas con el ítem 10, 11,12, 16 y 17 del Imp. a las Ganancias cuando la fecha de pago es Febrero 2014 (liquidando Diciembre 2013 o posterior).


'Global Const Version = "5.56"
'Global Const FechaModificacion = "05/02/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'Varias modificaciones
''   CAS-22983 - Bodega Norton - LIQ - Ganancias - Ajuste anual 2013
''                       Nueva correccion a la formula de ganancia FOR_GANANCIAS2013
''                               Problemas con el ítem 16 del Imp. a las Ganancias con Zona desfavorable y mayor a 25000.
''       Rectificativa 2013
''                           Deducciones especiales:
''                               Si hubo retenciones entre Enero y Agosto 2013 ==> debo recalcular deduccion especial y Ganancia Imponible
''                           Item 31
''                               Se agrega un parametro extra a la formula tal que si está configurado en valor <> 0  aplica incremento (igual que item 17)
''                               Para los que no configuren el parametro o esté configurado en 0 no se aplicará incremento del item
''                           Items 10,11 y 12
''                               Hay que hacer una analisis mes por mes dado que lo valores de la escala fueron variando y el calculo estandar no es exacto.
''                               Se divide el calculo en 3 zonas. De Enero a Febrero, de Marzo a Agosto y de Septiembre a Diciembre
''
''   Ademas
'    'CAS-19564 - Raffo - BUG en Busqueda SAC
''                       Tipo de Busqueda 78: SAC Proporcional tope de 30 x mes (bus_DiasSAC_Proporcional_Tope30)"
''                               Se le agregò el control para que solo descuenta licencias en estado aprobada.
''   Ademas
''   CAS-23019 - RHPro Consulting - Santander - Error Tabla Impunicab
''                       Se modificó la formula interna de Recalculo de impuesto Unico (Chile) - for_RecalcImpuestoUnico
''                       Se cambió el calculo de la zona Extrema2 para el calculo del nuevo impuesto.
''                           El valor de zona extrema se estaba arrastrando en todos los periodos analizados y no se estaba limpiando de periodo en periodo.
''   Ademas
''   CAS-21979 - PWC CL - Modificacion calculo reliquidacion con adicional de salud
''                               Se modifico la formula de Chile for_RecalcConcepto()
''   Ademas
''                       Nueva Tipo de Busqueda 125: Vacaciones no Gozadas Pendientes (El Salvador). Sub bus_Vac_No_Gozadas_Pendientes_SV
''                               calcula la proporcion de días de vacaciones sin gozar a la fecha de fin del proceso.
''                       Nueva Tipo de Busqueda 126: Antiguedad Aniversario. Sun bus_ant_aniversario_fin_proceso_SV:
''                               esta busqueda calcula la antiguedad aniversario hasta la fecha de fin del proceso de liquidacion.
''                       Nueva Tipo de Busqueda 127: Antiguedad Aguinaldo SV. bus_aguinaldo_SV
''                               Calcula la antiguedad para el aguinaldo del salvador. Aniversario, con tope a 365 días.
''

'Global Const Version = "5.57"
'Global Const FechaModificacion = "11/02/2014"
'Global Const UltimaModificacion = " " 'FGZ - EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'Varias modificaciones
''CAS-22822 - Raffo - Custom en calculo de antiguedad
''
''
''                       Nueva Tipo de Busqueda 128: Licencias Integradas por fecha de certificado (El Salvador). Sub bus_Licencias_Integradas_Certif
''                               calcula la cantidad de dias de licencias de las licencias cuya fecha de certificado cae dentro del proceso analizado.
''                                  Simil busqueda tipo 82
''                       Nueva Tipo de Busqueda 129: Licencias Parciales en horas por fecha de Certificado (El Salvador). Sub bus_Licencias_parciales_Certif
''                               calcula la cantidad de dias de licencias de las licencias cuya fecha de certificado cae dentro del proceso analizado.
''                                  Simil busqueda tipo 112
''Margiotta, Emanuel- CAS- 24100 Sykes El Salvador
''   Tipo de Busqueda 82: se agrego una nueva opcion que busca las licencias completas, parciales o todas (EAM)


'Global Const Version = "5.58"
'Global Const FechaModificacion = "21/02/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'Varias modificaciones
''   CAS-22983 - Bodega Norton - LIQ - Ganancias - Ajuste anual 2013
''                           Items 10,11 y 12
''                               Ajuste en el calculos de los items entre Septiembre a Diciembre
''
''   Ademas

'Global Const Version = "5.59"
'Global Const FechaModificacion = "24/02/2014"
'Global Const UltimaModificacion = " " 'EAM - FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-22983 - Bodega Norton - LIQ - Ganancias - Ajuste anual 2013
''                           Se corrigio un error en un sql de Ganancias2013
''
''                           Items 10,11 y 12
''                               Ajuste en el calculos de los items entre Septiembre a Diciembre

'Global Const Version = "5.60"
'Global Const FechaModificacion = "25/02/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-22983 - Bodega Norton - LIQ - Ganancias - Ajuste anual 2013
''                               Ajuste en el calculo 2013. para los que hubiesen tenido retenciones antes de Septiembre 2013 y estaban por debajo de 15000
''
''   Ademas
''           Limpieza de las trazas en las formulas estaba limpiando mal
''           Se cambio El uso de la funcion LimpiarTraza por LimpiarTrazaConcepto.
''
'' Ademas
''           CAS-22822 - Raffo - Custom en calculo de antiguedad
''               Se corrigió el tipo de busqueda 97 (Nueva Antiguedad). Cuando la fecha de corte es anterior a la fecha de alta del empleado estaba dando negativo
''
'' Ademas
''           CAS-23723 - G.Compartida - Inconvenientes con las Búsquedas Automaticas
''               Tipo de Busqueda 22: Acumuladores Imponibles Mensuales. Se modificaron todos los subs de suma, minimo, maximo y promedio para cuando la cantidad de meses a buscar es o.

'Global Const Version = "5.61"
'Global Const FechaModificacion = "28/02/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-22822 - Raffo - Custom en calculo de antiguedad
''               Se corrigió el tipo de busqueda 97 (Nueva Antiguedad). Cuando la fecha de corte es anterior a la fecha de alta del empleado estaba dando negativo
''
''           CAS-23106 - HORWATH LITORAL - AMR - Modificacion Busqueda VNG
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja.
''                   Se le agregó la opcion de que descuente dias de vacaciones ya gozados. Ademas se le sacó el redondeo.

'Global Const Version = "5.62"
'Global Const FechaModificacion = "05/03/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-21182 - H&A - ERROR EN CALCULO DE EMBARGOS
''               Tipo de Busqueda 22: Busqueda de Embargos. Para cuando es quincenal, en segunda quincena, se debe tomar el SMVM entero y no la mitad como hace actualmente.


'Global Const Version = "5.63"
'Global Const FechaModificacion = "05/03/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-22822 - Raffo - Custom en calculo de antiguedad
''               Se corrigió el tipo de busqueda 97 (Nueva Antiguedad). Cuando la fecha de corte es anterior a la fecha de alta del empleado estaba dando negativo

'Global Const Version = "5.64"
'Global Const FechaModificacion = "07/03/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-21979 - PWC CL - Modificacion calculo reliquidacion con adicional de salud
''                               Se modifico la formula de Chile for_RecalcConcepto()

'Global Const Version = "5.65"
'Global Const FechaModificacion = "17/03/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-24465 - SYKES SV - LIQ - Busqueda de acumulado diario en periodo de liq
''                       Nueva Tipo de Busqueda 130: Acumulado Diario de Horas
''                               calcula la suma de horas del acumulado diario entre las fechas del proceso de liquidacion.
''Ademas
''   CAS-21979 - PWC CL - Modificacion calculo reliquidacion con adicional de salud
''                               Se modifico la formula de Chile for_RecalcConcepto()

'Global Const Version = "5.66"
'Global Const FechaModificacion = "19/03/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-24465 - SYKES SV - LIQ - Busqueda de acumulado diario en periodo de liq
''                       Tipo de Busqueda 130: Acumulado Diario de Horas
''                               Se corrigió problema levantando parametro de periodo de analisis.

'Global Const Version = "5.67"
'Global Const FechaModificacion = "19/03/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-24465 - SYKES SV - LIQ - Busqueda de acumulado diario en periodo de liq
''                       Tipo de Busqueda 130: Acumulado Diario de Horas
''                               Se corrigió problema levantando parametro de periodo de analisis.

'Global Const Version = "5.68"
'Global Const FechaModificacion = "31/03/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-23106 - HORWATH LITORAL - AMR - Modificacion Busqueda VNG
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja.
''                   Se le corrigiò la definicion de la variable de cantidad porque estaba como entera y por la quita del redondeo anterior debe quedar como double.


'Global Const Version = "5.69"
'Global Const FechaModificacion = "04/04/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-24432 - TELEFAX - MONRESA - Error busqueda de liquidacion
''               Tipo de Buqueda 7: 'Acumuladores Mensuales Meses Fijos. (sub Bus_Acum3)
''                                   'Estaba teniendo en cuenta un mes menos cuando se buscaba Anual sin incluir periodo ni proceso actual



'Global Const Version = "5.70"
'Global Const FechaModificacion = "08/04/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-24955 - H&A - Ganancias cambio de criterio item 31 aumento de 20% segunda franja
''                           Nuevo cambio en formula de Ganancias for_Ganancias2013
''                           Items 31
''                               Ahora funciona igual que el item 17 con la salvedad que se topea al valor de la escala

'Global Const Version = "5.71"
'Global Const FechaModificacion = "09/04/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-24955 - H&A - Ganancias cambio de criterio item 31 aumento de 20% segunda franja
''                           Nuevo cambio en formula de Ganancias for_Ganancias2013
''                           Items 31. Se topea la DDJJ contra la escala incrementeada
''Ademas
''           CAS-24432 - TELEFAX - MONRESA - Error busqueda de liquidacion
''               Tipo de Buqueda 7: 'Acumuladores Mensuales Meses Fijos. (sub Bus_Acum3)
''                                   'Estaba teniendo en cuenta un mes menos cuando se buscaba Anual sin incluir periodo ni proceso actual


'Global Const Version = "5.72"
'Global Const FechaModificacion = "09/04/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-23307 - Raffo - Custom Comisiones APM
''               Nueva Formula : FOR_COMISIONES ()
'' Ademas
''           CAS-22808 - SGS - Distribución Contable
''                   Busqueda de Novedades (Nov_GEGI): Chequea si las novedades individuales tienen distribucion e inserta distribucion del concepto
''                   Novedades de Ajuste (Machetazo): Chequea si las novedades de ajuste tienen distribucion e inserta distribucion del concepto
''


'Global Const Version = "5.73"
'Global Const FechaModificacion = "15/04/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-25034 - H&A - LIQ - Ganancias - Bienes Personales
''               Se requiere modificar la fórmula de Ganancias, para los casos de menos de 15000 con retenciones hasta Agosto 2013,
''               para que calcule la ganancia imponible a Dic/2013 determinando el mismo impuesto que en Agosto 2013.


'Global Const Version = "5.74"
'Global Const FechaModificacion = "30/04/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-23307 - Raffo - Custom Comisiones
''               Formula : FOR_COMISIONES (). No estaba controlando la vigencia de las escalas.
''
''Ademas
''           CAS-25331 - H&A - ERROR EN LIQUIDADOR 5.73 AJUSTE DE GANANCIAS ITEM 31
''                           formula de Ganancias for_Ganancias2013
''                           Items 31. está calculando el valor topeado aun cuando no hay DDJJ y en ese caso debe dar 0
''Ademas
''           CAS-23106 - HORWATH LITORAL - AMR - Modificacion Busqueda VNG
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja.
''                   Se le agregó detalle de log a la busqueda.


'Global Const Version = "5.75"
'Global Const FechaModificacion = "12/05/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-25331 - H&A - Error liquidador 5.74 item 17 y 31
''                           formula de Ganancias for_Ganancias2013
''                           Items 31. Cuando el valor de DDJJ es menor al tope no está dando el resultado correcto.
''
''           CAS-25331 - H&A - Error en segmentacion de ganancias para nuevos ingresos
''                           formula de Ganancias for_Ganancias2013
''
''Ademas
''               URUGUAY
''               Formulas : FOR_IRPF (), for_irpf_simple() y for_irpf_diciembre(). Se redefinieron los tamaños de arreglos de DDJJ.
''
''               CHILE
''               Formulas : for_ImpuestoUnico(), . Se redefinieron los tamaños de arreglos de DDJJ.

'Global Const Version = "5.76"
'Global Const FechaModificacion = "14/05/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-24145 - Santander Uruguay - Busqueda Liq
''                       Nueva Tipo de Busqueda 131: Antiguedad con Redondeo cada 6 meses

'Global Const Version = "5.77"
'Global Const FechaModificacion = "19/05/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-25371 -  Nacionalización Uruguay - Funcionalidad Paros
''                       Nueva Tipo de Busqueda 132: Horas de Paros Sindicales


'Global Const Version = "5.78"
'Global Const FechaModificacion = "26/05/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-25682 - Raffo - Bug en liquidación de haberes de empleados nuevos en el proceso
''                           formula de Ganancias for_Ganancias2013
''           Ademas se agregó un control sobre el borrado de ficharet cuando no hay detalles de liquidacion en el proceso (sub LIQPRO04()
''

'Global Const Version = "5.79"
'Global Const FechaModificacion = "26/06/2014"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-24145 - Santander Uruguay  - Busqueda de liquidación 131
''          Busqueda 131: Se modifico la búsqueda para que devuelva resultados en días - mes - año


'Global Const Version = "5.80"
'Global Const FechaModificacion = "08/07/2014"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-21788 - Sykes SV - Modificacion búsqueda 82 - Licencias integral
''          Busqueda 82: Se agrego la opcion para los periodos de GTI buscar por 2 criterios. 1) Primer perido de GTI. 2) Todos los periodos de GTI
''          Busqueda 112: Se agrego la opcion para los periodos de GTI buscar por 2 criterios. 1) Primer perido de GTI. 2) Todos los periodos de GTI
''          Busqueda 126: Se agrego una validación en la busqueda de antiguedad que controla si la fecha de fin es mayor a la del proceso, se queda con la fecha de fin de proceso


'Global Const Version = "5.81"
'Global Const FechaModificacion = "10/07/2014"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-25222 - Tabacal - Licencias medicas -
''          Busqueda 102: Se agrego una opcion de busqueda por proceso de liquidación a la función (bus_DiasHabLic).


'Global Const Version = "5.82"
'Global Const FechaModificacion = "16/07/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-25027 - CAPUTO - Aprobacion de vales desde MSS
''               Tipo de Busqueda 32: Vales. Se le agregó control de firmas
''

'Global Const Version = "5.83"
'Global Const FechaModificacion = "16/07/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-26398 - Santander URU - Busqueda dias trabajados año/semestre actual
''               Nuevo Tipo de Busqueda 133 - Tiempo Efectivamente Trabajado -
''                   bus_Tiempo_Trabajado()

'Global Const Version = "5.84"
'Global Const FechaModificacion = "16/07/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-25222 - Tabacal - Licencias medicas -
''          Busqueda 102: Se agrego una opcion de feriados Habiles a la función (bus_DiasHabLic).

'Global Const Version = "5.85"
'Global Const FechaModificacion = "18/07/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-23106 - HORWATH LITORAL - AMR - Modificacion Busqueda VNG
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja.
''                   Cuando la antiguedad es menor a 6 meses se le considerará 6 meses para que la escala siempre arroje resultados.
'

'Global Const Version = "5.86"
'Global Const FechaModificacion = "30/07/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-26302 - RH PRO CHILE - MODIFICACION DE FORMULA DE IMPUESTO UNICO
''               CHILE
''               Formulas : for_ImpuestoUnico(). Se agregó parametro "rapa nui" a la formula.
'

'Global Const Version = "5.87"
'Global Const FechaModificacion = "15/08/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-26398 - Santander URU - Busqueda dias trabajados año-semestre actual
''               Busqueda 133 - Tiempo Efectivamente Trabajado -
'' Ademas
''               Busqueda 126: Se agrego una validación en la busqueda de antiguedad que controla si la fecha de fin es mayor a la del periodo, se queda con la fecha de fin de periodo

'Global Const Version = "5.88"
'Global Const FechaModificacion = "08/09/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-26398 - Santander URU - Busqueda dias trabajados año-semestre actual
''               Busqueda 133 - Tiempo Efectivamente Trabajado -
'' Ademas
''               Busqueda 126: Se agrego una validación en la busqueda de antiguedad que controla si la fecha de fin es mayor a la del periodo, se queda con la fecha de fin de periodo
''Ademas
''           CAS-23106 - HORWATH LITORAL - AMR - Modificacion Busqueda VNG
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja.
''                   Se debe proporcionar la antiguedad solo al ultimo año.

'Global Const Version = "5.89"
'Global Const FechaModificacion = "19/09/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-26302 - RH PRO CHILE - MODIFICACION DE FORMULA DE IMPUESTO UNICO
''               CHILE
''               Formulas : for_ImpuestoUnico(). Cuando reliquida los periodos anteriores está tomando el acumulado mes a mes  y debería tomar solo el mes actual
'' Ademas
''           Ademas se le agregaron mas opciones de recoleccion de datos estadisticos de ejecucion del liquidador
''
'' Ademas
''           CAS-27036 - 5CA - Bug Busqueda Automatica
''               Tipo de Buqueda 82: 'Licencias Integrales. Se corrigió parametro de tipo de licencia.
'' Ademas
''           CAS-26990 - MONASTERIO - ERROR EN NETO NEGATIVO
''               Clase CNuevaCache. Se formatean los valores para que no de el resultado en notacion cientifica.
''
''           CAS-23106 - HORWATH LITORAL - AMR - Modificacion Busqueda VNG
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja. Estaba buscando mal la fase.

'Global Const Version = "5.90"
'Global Const FechaModificacion = "30/09/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-17053 - Nac Brasil - Aviso Previo de Baja - Busquedas
''               Nuevo Tipo de Busqueda 134 - Dias de Pre Aviso
''                   bus_PreAviso()

'Global Const Version = "5.91"
'Global Const FechaModificacion = "02/10/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-26789 - Santander Uruguay - Búsqueda de Antiguedad [Entrega 2]
''               Tipo de Busqueda 131: Antiguedad con Redondeo cada 6 meses
''                   Se cambió la validación de la fecha de corta de cada año analizado
''
''Ademas por el caso CAS-17053 - Nac Brasil - Aviso Previo de Baja - Busquedas
''               Tipo de Busqueda 134 - Dias de Pre Aviso.bus_PreAviso()
''                   Se corrige referencia incorrecta a campo de tabla fases


'Global Const Version = "5.92"
'Global Const FechaModificacion = "15/10/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-17053 - Nac Brasil - Aviso Previo de Baja - Busquedas
''               Tipo de Busqueda 134 - Dias de Pre Aviso.bus_PreAviso()
''                   Se agregó parametro al tipo de busqueda. Busca descuento o pago


'Global Const Version = "5.93"
'Global Const FechaModificacion = "22/10/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-17053 - Nac Brasil – Aviso Previo de Baja - Busquedas
''               Tipo de Busqueda 134 - Dias de Pre Aviso.bus_PreAviso()
''                   Se cambió definicion de campo de dias para que soporte nros largos (Cuando omiten poner fecha se toma default 01/01/1900).
''
''Ademas
''               Nuevo Tipo de Busqueda 136 - Vacaciones no Gozadas Pendientes Años Anteriores (Uruguay)
''                   bus_Vac_No_Gozadas_Pendientes_UY()
''
''               Tipo de Busqueda 133: bus_Tiempo_Trabajado
''                   Se agregaron 2 opciones nuevas a los Parametros Forma de Calculo (3) y Tipo de Fecha (7)


'Global Const Version = "5.94"
'Global Const FechaModificacion = "23/10/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-27353 - GC - Compatibilizar proceso de liquidacion R0 a R3
''               Se nivelaron todas las busquedas de BAE respecto de las modificaciones del cliente (TTI)
''                   Bus_PartesDiarios()
''                   Bus_BAE()
''                   Bus_Refrigerios()
''                   Generar_Penalidades()
''                   Generar_Sanciones()


'Global Const Version = "5.95"
'Global Const FechaModificacion = "24/10/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-27511 -  Sykes El Salvador - Retroactivo Nocturnidad
''               Nuevo Tipo de Busqueda 137 - Horas Pagas Retroactivas
''                   bus_Horas_Pagadas_Retro()


'Global Const Version = "5.96"
'Global Const FechaModificacion = "11/11/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-23106 - HORWATH LITORAL - AMR - Modificacion Busqueda VNG
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja. Se cambió el calculo de Antiguedad y la forma de proporcionar respecto del ultimo año.
''
''Ademas
''        EAM- 31/10/2014 - CAS-27511 -  Sykes El Salvador - Retroactivo Nocturnidad
''               Tipo de Busqueda 137 - Horas Pagas Retroactivas. bus_Horas_Pagadas_Retro()
''               se agrego los parentesis a la expresion porque estaba haciendo mal el calculo: NovHorasNoc = 11 * (DateDiff("d", Aux_Fecha_Desde, Aux_Fecha_Hasta) + 1)
''Ademas
''               Nuevo Tipo de Busqueda 138 - Vacaciones Vendidas
''                   bus_vac_vendidas()

'Global Const Version = "5.97"
'Global Const FechaModificacion = "01/12/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-17053 - Nac Brasil - Cálculo IRRF
''               Formulas : for_IRRF(). Formula para el calculo del impuesto en Brasil
'' Ademas
''           CAS-26302 - RH PRO CHILE - MODIFICACION DE FORMULA DE IMPUESTO UNICO
''               CHILE
''               Formulas : for_RecalcConcepto(). Calculo de adicional salud
''
''Ademas
''        EAM- CAS-27511 -  Sykes El Salvador - Retroactivo Nocturnidad
''               Tipo de Busqueda 137 - Horas Pagas Retroactivas. bus_Horas_Pagadas_Retro()
''               Se le agregaron parametros a la busqueda
'

'Global Const Version = "5.98"
'Global Const FechaModificacion = "10/12/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-27179 - AGD - MEJORA DE BUSQUEDA DE PERIODO DE GTI
''               Tipo de Busqueda 106 - Desgloce de Horas. bus_Desg_Horas()
''               Se le agregó la validacion de periodo de la empresa del empleado cuando se busca el periodo de gti

'Global Const Version = "5.99"
'Global Const FechaModificacion = "16/12/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-28440 - H&A - LIQ - Cambio legal Ganancias Dec 2354 2014
''           formula de Ganancias for_Ganancias2013. Ajustes por decreto 2354 2014.
''
''
''Ademas
''           CAS-26302 - RH PRO CHILE - MODIFICACION DE FORMULA DE IMPUESTO UNICO
''               CHILE
''               Formulas : for_RecalcConcepto(). Calculo de adicional salud
''               Cuando la diferencia de adicional salud es mayor al calculo ==> debe guardar o

'Global Const Version = "6.00"
'Global Const FechaModificacion = "30/12/2014"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-28749 - H&A - LIQ - Ganancias - Bug rango de 15000 a 25000 1er liq del año
''           formula de Ganancias for_Ganancias2013. Ajustes por decreto 2354 2014. Primer liquidacion 2015 para rango 2

'Global Const Version = "6.01"
'Global Const FechaModificacion = "05/01/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-28749 - H&A - LIQ - Ganancias - Bug rango de 15000 a 25000 1er liq del año
''           formula de Ganancias for_Ganancias2013. Correccion de mensajes de log en Rango menor a 15000

'Global Const Version = "6.02"
'Global Const FechaModificacion = "05/01/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-28749 - H&A - LIQ - Ganancias - Bug rango de 15000 a 25000 1er liq del año
''           formula de Ganancias for_Ganancias2013. Correccion de mensajes de log en Rango menor a 15000


'Global Const Version = "6.03"
'Global Const FechaModificacion = "20/01/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''               CAS-27512 - H&A - LIQ - Ganancias - Item 56 Perc.Compras Exterior mensual
''                   formula de Ganancias for_Ganancias2013. Se agregó un parametro (1140) a la formula para poder dar el beneficio de devolucion
''                   anticipada de compras en el exterior (item 56)
''Ademas
''               Formulas : for_ImpuestoUnico(). Cuando reliquida los periodos anteriores está tomando el acumulado mes a mes  y debería tomar solo el mes actual.
''                           Se ajustó el cambio del 19/09/2014 (V5.89)


'Global Const Version = "6.04"
'Global Const FechaModificacion = "29/01/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''   CAS-29187 - SYKES EL SALVADOR - Bug Lic Integradas x fecha de certificado
''           Busqueda 137: se modifico la busqueda cdo busca las Novedades /Licencias completas para que busque el tipo de hora de origen.
''                       se modifico la busqueda para que calcule las novedades retroactivas de "Citas Programas". Lo que hace es calcular el excedente de descuento.
''                            Ej: si se iforma una novedad de 5.30 horas y la busqueda esta configurada mayora a 4. calcula 1.5 que es lo que se debe descontar.
''   Busqueda 139: EAM- Se agrego una nueva busqueda de licencia con fecha de certificado (retroactivas) que controla las licencias pagas en tiempo y forma.


'Global Const Version = "6.05"
'Global Const FechaModificacion = "23/02/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29317 - H y A - LIQ - Bug en Calculo de Impuesto a las Ganancias.
''               Argentina
''                   Formula:  for_Ganancias2013. Se agregó un parametro (1140) a la formula para poder dar el beneficio de devolucion
''                               Se corrigió problema de franja > 2500 y zona 20. Para estos NO debe sumar el 20 pero si el valor del aguinaldo
''
''Ademas
''   Tipo de Busqueda 82: Licencias Integradas. Se corrigió problema de topes cuando la suma de las licencias actuales mas las ya marcadas no llegan al tope
''
''   CAS-21778 - Sykes El Salvador- QA - Bug Liquidador
''       Tipo de Busqueda 137: Horas Pagas Retroactivas. Se le agregó control de division por cero cuando no se le asigna regimen horario.
''

'Global Const Version = "6.06"
'Global Const FechaModificacion = "27/03/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29261 - Horwath litoral - AMR - Modificación Búsqueda VNG
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja.
''                   Se debe proporcionar siempre los dias segun los dias trabajados en el ultimo año.
''   Busqueda 124: EAM- se modifico la query que busca las licencias de vacaciones
''   Busqueda 128: EAM- Se modifico la busqueda de licencias por fecha de certificado. Estaba calculando mal cuando era para  febrero los topes.

'Global Const Version = "6.07"
'Global Const FechaModificacion = "06/04/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30295 - RH Pro Producto - LIQ - Ganancias - Bug liquidador 6.06 - Falta de item 56
''               Argentina
''                   Formula:  for_Ganancias2013. Item 56. Control de valores sobre ddjj cargadas.
''Ademas
''           CAS-29032 - Telefax (Santander URU) - Bug Búsqueda de antiguedad
''               Tipo de Busqueda 131(Antiguedad con Redondeo cada 6 meses). Se modificó la logica para que solo redondee el 1er año de la 1er fase


'Global Const Version = "6.08"
'Global Const FechaModificacion = "09/04/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29945 - SYKES EL SALVADOR - Error búsq. Antig Aniversario
''               Tipo de Busqueda 126(Calculo de antiguedad aniversario a fin de proceso): Se toma seimpre la ultima fase del empleado.

'Global Const Version = "6.09"
'Global Const FechaModificacion = "16/04/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29032 - Telefax (Santander URU) - Bug Búsqueda de antiguedad
''               Tipo de Busqueda 131(Antiguedad con Redondeo cada 6 meses). Se modificó la logica para que solo redondee el 1er año de la 1er fase
''                                   Se cambió el redondeo nuevamente. Ahora es: si ant >= 6 meses ==> 1 años, sino se consideran los meses y dias.

'Global Const Version = "6.10"
'Global Const FechaModificacion = "17/04/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30490 - SANTANDER URUGUAY - Error busqueda tiempo trabajado
''               Tipo de Busqueda 133(Tiempo Efectivamente Trabajado). Se corrigió la logica, solo descontaba licencias en modo debbug.

'Global Const Version = "6.11"
'Global Const FechaModificacion = "20/04/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29945 - SYKES EL SALVADOR - Error búsq. Antig Aniversario
''               Tipo de Busqueda 126(Calculo de antiguedad aniversario a fin de proceso): Se toma seimpre la ultima fase del empleado.
''               Cuando la fecha de baja es mayor a la fecha de fin de proceso estaba tomando mal la fecha de corte.

'Global Const Version = "6.12"
'Global Const FechaModificacion = "21/04/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29945 - SYKES EL SALVADOR - Error búsq. Antig Aniversario
''               Tipo de Busqueda 126(Calculo de antiguedad aniversario a fin de proceso): Se toma seimpre la ultima fase del empleado.
''               Cuando la fecha de baja prevista es mayor a la fecha de fin de proceso estaba tomando mal la fecha de corte.

'Global Const Version = "6.13"
'Global Const FechaModificacion = "23/04/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-21778 - Sykes El Salvador - QA – Busqueda Prestamos
''               Tipo de Busqueda 14(Prestamos): Se agregó opcion para cancelar siempre las cuotas (dejando su valor en 0) y sin crear cuota para otro periodo
''                                               Cuando se descuenta todo o nada y no puede descontar.


'Global Const Version = "6.14"
'Global Const FechaModificacion = "30/04/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29032 - Telefax (Santander URU) - Bug Búsqueda de antiguedad
''               Tipo de Busqueda 131(Antiguedad con Redondeo cada 6 meses). Se modificó la logica para que redondee al final de la suma de la antiguedad total.
''                                   si la cantidad de meses que no llegan al año es > 6 meses ==> se suma un año mas, sino no se consideran para la cantidad de años.

'Global Const Version = "6.15"
'Global Const FechaModificacion = "11/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29032 - Telefax (Santander URU) - Bug Búsqueda de antiguedad
''               Tipo de Busqueda 131(Antiguedad con Redondeo cada 6 meses). Se modificó la logica para que redondee a 30 la cantidad de dias para los meses de 31 dias.


'Global Const Version = "6.16"
'Global Const FechaModificacion = "11/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29325 - IMECON - Nueva Búsqueda para dias de vacaciones pendientes de goce
''               Nuevo Tipo de Busqueda 140(Saldo Vacaciones con Venta (Perú)): dias correspondientes - dias gozados - dias vendidos - dias vencidos. Genera venta de vacaciones.

'Global Const Version = "6.17"
'Global Const FechaModificacion = "12/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
'''               Argentina
'''                   Formula:  for_Ganancias2013. Cambios Legales.

'Global Const Version = "6.18"
'Global Const FechaModificacion = "12/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30516 - GEST COMPARTIDA (EDENOR) - Custom agregar funcion a liquidador
''               Se nivelaron todas las busquedas de BAE respecto de las modificaciones del cliente (GC)
''                   Generar_Sanciones()

'Global Const Version = "6.19"
'Global Const FechaModificacion = "13/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
''               Argentina
''                   Formula:  for_Ganancias2013. Cambios Legales.Controles adicionales para la determinacion de la franja.

'Global Const Version = "6.20"
'Global Const FechaModificacion = "18/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
''               Argentina
''                   Formula:  for_Ganancias2013. Cambios Legales.Controles adicionales para la determinacion de la franja.

'Global Const Version = "6.21"
'Global Const FechaModificacion = "18/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
''               Argentina
''                   Formula:  for_Ganancias2013. Cambios Legales.Controles adicionales para la determinacion de la franja.

'Global Const Version = "6.22"
'Global Const FechaModificacion = "19/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
''               Argentina
''                   Formula:  for_Ganancias2013. Cambios Legales.Controles adicionales para la determinacion de la franja.
''
''Ademas, por el caso
''           CAS-30682 - Monasterro base AMR - Bug en liquidar sin análisis detallado
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja. Se corrige problemas de variables de log


'Global Const Version = "6.23"
'Global Const FechaModificacion = "21/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30979 - RH Pro (Producto) - LIQ - Ganancias - RG 3770 nuevo cambio
''               Argentina
''                   Formula:  for_Ganancias2013. Cambios Legales.Controles adicionales para la determinacion de la franja.

'Global Const Version = "6.24"
'Global Const FechaModificacion = "24/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-30979 - RH Pro (Producto) - LIQ - Ganancias - RG 3770 nuevo cambio
''               Argentina
''                   Formula:  for_Ganancias2013. Cambios Legales.Controles adicionales para la determinacion de la franja.

'Global Const Version = "6.25"
'Global Const FechaModificacion = "27/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-31099 - RH Pro (Producto) - LIQ - Ganancias - Corrección fórmula
''               Argentina
''                   Formula:  for_Ganancias2013. Cambios Legales.Controles adicionales para la determinacion de la franja.

'Global Const Version = "6.26"
'Global Const FechaModificacion = "28/05/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-31075 - Telefax (Santander URU) - Búsqueda de antiguedad
''               Tipo de Busqueda 131(Antiguedad con Redondeo cada 6 meses). Se modificó la logica para que redondee los dias y meses de la misma forma que los años cuando se redondean los años.

'Global Const Version = "6.27"
'Global Const FechaModificacion = "01/06/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29187 - SYKES EL SALVADOR - Bug Lic Integradas x fecha de certificado
''           Busqueda 139: (Licencias Integradas por fecha de certificado con control de lic. integradas).
''                       Se corrigió problema de topes cuando la suma de las licencias actuales mas las ya marcadas no llegan al tope.

'Global Const Version = "6.28"
'Global Const FechaModificacion = "02/06/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-31205 - RH Pro (Producto) - LIQ - Ganancias - Modificación escala interna porcentajes
''               Argentina
''                   Formula:  for_Ganancias2013. Cambios Legales. Ajuste en los valores de porcentajes de incremento en las deducciones para zona patagónica.


'Global Const Version = "6.29"
'Global Const FechaModificacion = "02/06/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-29187 - SYKES EL SALVADOR - Bug Lic Integradas x fecha de certificado
''           Busqueda 139: (Licencias Integradas por fecha de certificado con control de lic. integradas).
''                       Se corrigió problema de topes cuando la suma de las licencias actuales mas las ya marcadas no llegan al tope.
''           Busqueda 128: (Licencias Parciales en horas por fecha de Certificado).
''                       Se corrigió problema de topes cuando la suma de las licencias actuales mas las ya marcadas no llegan al tope.

'Global Const Version = "6.30"
'Global Const FechaModificacion = "24/06/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-31053 - RH Pro (Producto) - NAC. PERU – EPS – Solicitud de Búsqueda para  Cálculos de EPS
''           Nuevo tipo de Busqueda 141: (EPS - Perú).
''                       Retorna cantidad de Hijos o Precio final del plan de OS elejida del empleado.

'Global Const Version = "6.31"
'Global Const FechaModificacion = "25/06/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-31674 - CDA - Bug en liquidación mensual de Junio
''               Tipo de Buqueda 7:   Acumuladores Mensuales Meses Fijos.    (sub Bus_Acum3). Estaba calculando mal el mes hasta cuando se busca semestre anterior sin mes de inicio fijo.
''               Tipo de Buqueda 107: Acum. Mensual Mes Fijo con Ajuste de aumentos.(bus_Acum3_Con_Ajuste). Estaba calculando mal el mes hasta cuando se busca semestre anterior sin mes de inicio fijo.


'Global Const Version = "6.32"
'Global Const FechaModificacion = "03/07/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-31676 - RH Pro - Modificación calculo aguinaldo IRPF
''               URUGUAY
''                   Formula:  for_irpf_simple(). Cambios Legales. Modificaciones por Ley No. 19.321 de 29/05/2015. Decreto 154/015 de 1/06/2015.
''
''
''Ademas, por caso
''           CAS-29961 - VISION - Error en busqueda de prestamos
''               Tipo de Busqueda 14(Prestamos): Se agregó opcion para retornar Cuota pura para los casos en que los intereses no estan incluidos en la misma).


'Global Const Version = "6.33"
'Global Const FechaModificacion = "31/07/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
'           CAS-29325 - IMECON - Nueva Búsqueda para dias de vacaciones pendientes de goce
'               Tipo de Busqueda 140(Saldo Vacaciones con Venta (Perú)): bus_SaldoVac_PE(). Correcion de referencia de campo

'Global Const Version = "6.34"
'Global Const FechaModificacion = "31/07/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " 'CAS-29325 - IMECON - Nueva Búsqueda para dias de vacaciones pendientes de goce [Entrega 3]
''Tipo de Busqueda 140 - Saldo vacaciones PE - Se agrego la funcinalidad de días truncos
''Tipo de Busqueda 120 - CTS - Se agrego la opcion de días truncos y se modifico los dias pendientes
'

'Global Const Version = "6.35"
'Global Const FechaModificacion = "10/08/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "  'CAS-29325 - IMECON - Nueva Búsqueda para dias de vacaciones pendientes de goce [Entrega 4]
''               Tipo de Busqueda 140 - Saldo vacaciones PE - Se modificó el calculo de los dias adquiridos pero no gozados
''               Tipo de Busqueda 120 - CTS - Tiempo Efectivamente Trabajado - bus_CTS_Tiempo_Trabajado(). Se modificó el control sobre si se incluye o no el ultimo día como trabajado
''Ademas,
''   Actualizo la cantidad de empleados en batch_proceso bprcempleados
''   Modificaciones sobre el control de firmas. Control sobre rechazadas


'Global Const Version = "6.36"
'Global Const FechaModificacion = "26/08/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " "  'CAS-32523 - SANTANDER URUGUAY - LIQ – Bug búsqueda de antigüedad
''               Tipo de Busqueda 131(Antiguedad con Redondeo cada 6 meses). Se modificó la logica para que topee la antiguedad individualmente por cada año como suma de las partes de las fases contenidas en el mismo.

'Global Const Version = "6.37"
'Global Const FechaModificacion = "01/09/2015"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-31053 - RH Pro (Producto) - NAC. PERU – EPS – Solicitud de Búsqueda para  Cálculos de EPS [Entrega 2]
''           Tipo de Busqueda 141: (EPS - Perú).
''                       Estaba calculando mal con fecha de fin de periodo.


'Global Const Version = "6.38"
'Global Const FechaModificacion = "14/09/2015"
'Global Const UltimaModificacion = " " 'LED
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''           CAS-33005 - G.COMPARTIDA - Custom en función del liquidador
''               Se nivelaron todas las busquedas de BAE respecto de las modificaciones del cliente (GC)
''                   Generar_Sanciones()

'Global Const Version = "6.39"
'Global Const FechaModificacion = "07/10/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-33430 - CIVA - Bug Venta Vacaciones
''               Se modifico el borrado de venta de vacacioens para que solo borre los registros marcados como automatico.
''               Tipo de Busqueda 138: Se agregó el campo automatico cuando se genera la venta de días de vacacioens
''               Tipo de Busqueda 119: Se agregó el campo automatico cuando se genera la venta de días de vacacioens
''       CAS-33210 - SANTANDER URUGUAY - Busqueda base de calculo paros
''               Tipo de Busqueda 132: Se agrego dos opciones mas de búsqueda 5 y 6. Intersección de Periodo de GTI con Proceso de liq anterior (5) y con mes actual (6)

'Global Const Version = "6.40"
'Global Const FechaModificacion = "21/10/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-30667 - RH Pro (Producto) - LIQ - Ganancias - Items en ajuste anual - Filtro DDJJ personal
''           for_ganncia2013 - Se agrego control si es una liquidacion final o en el ajuste anual para que no se tenga en cuenta el  item 20.

'Global Const Version = "6.41"
'Global Const FechaModificacion = "17/11/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-31676 - RH Pro - Modificación calculo aguinaldo IRPF
''       CAS-34041 - MONASTERIO (TODAS LAS BASES Y ENTORNOS) - ERROR EN COPIA SIM
''           for_irpf_simple - Se agrego en traza el porcentaje maximo alcanzado y se cambio algunas descripciones de las traza que escribia
''           Tipo de Busqueda 140: Se agregaron dos campos, venc y vacnro que corresponden a la version de GIVR4 (bus_SaldoVac_PE)

'Global Const Version = "6.42"
'Global Const FechaModificacion = "19/11/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-31676 - RH Pro - Modificación calculo aguinaldo IRPF [Entrega 3]
''       Se corrigio el query del control de versión de la version 6.41

'Global Const Version = "6.43"
'Global Const FechaModificacion = "24/11/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-33993 - NGA - Ganancias residentes en el extranjero.
''           for_ganncia2013 - Se controla para los Extranjeros que perciben el salario en Argentina para que no tenga en cuenta los item 10,11,12 y 17
''       CAS-34164 - NGA - Modificacion de item 56 y 20 ganancias
''            for_ganncia2013 - Se agrego control para que se tenga en cuenta los item 20 y 56 en los meses de diciembre si la fecha de pago es 31/12 o liquidacion final
           
'Global Const Version = "6.44"
'Global Const FechaModificacion = "26/11/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-33993 - NGA - Ganancias residentes en el extranjero.
''           for_ganncia2013 - Se agregó nuevo parametro a la formula para controla empleados Extranjeros
''       CAS-34164 - NGA - Modificacion de item 56 y 20 ganancias [Entrega 2]
''            for_ganncia2013 - Se corrigio error cuadno controlaba el item 20

'Global Const Version = "6.45"
'Global Const FechaModificacion = "01/12/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-31053 - RH Pro (Producto) - NAC. PERU - EPS - Solicitud de Búsqueda para Cálculos de EPS [Entrega 3]
''           Tipo de Busqueda 141: Se agregó dos nuevas opciones: Precio Total de todos los empleados y Cantidad de Hijos de todos los empleados
''                               Estas dos opciones son resultados globales y se guradan en cache "objCache_BusquedasGlobales"


'Global Const Version = "6.46"
'Global Const FechaModificacion = "04/12/2015"
'Global Const UltimaModificacion = " " 'FGZ - EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-33657 - NGA BASE FREDDO - Bug en sac proporcional tope 30 mensual
''           Tipo de Buqsqueda 78: SAC Proporcional tope de 30 x mes (bus_DiasSAC_Proporcional_Tope30)
''                   Corrección en calculo sobre la cantidad de dias bajo condiciones
''Ademas,
''           Tipo de Busqueda 141: Se agregó control de nulo

'Global Const Version = "6.47"
'Global Const FechaModificacion = "14/12/2015"
'Global Const UltimaModificacion = " " 'FGZ - EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-33601 - RH Pro (Producto) - Peru - Criterios de Busquedas
''           Tipo de Buqsqueda 55: bus_BaseLicencias
''                   Se agregó control para setear la fecha que busca la licencia segun la configuración de Incluye y permite buscar por mas de una licencia
''           tipo de busqueda 120: bus_CTS_Tiempo_Trabajado
''                   Se agrego el seteo de la variable si descuenta licencia que no estaba
''           tipo de busqueda 140: bus_SaldoVac_PE
''                    Se agrego la funcionalidad para que genere las indemnizaciones detallado por periodo
''           tipo de busqueda 138: bus_Vac_Vendidas
''                    Se agregó la funcionalidad de buscar las ventas de vacaciones por año del periodo. Sino busca por proceso como lo hacía originalmente

'Global Const Version = "6.48"
'Global Const FechaModificacion = "16/12/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-31053 - RH Pro (Producto) - NAC. PERU - EPS - Solicitud de Búsqueda para Cálculos de EPS [Entrega 4]
''           Tipo de Buqsqueda 141: bus_EPS
''                   Se agregó nueva opcion: monto Acumulador de todos los empleados. Estas dos opciones son resultados globales y se guradan en cache "objCache_BusquedasGlobales"


'Global Const Version = "6.49"
'Global Const FechaModificacion = "21/12/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-32751 - LA CAJA - Custom Seguros ADP
''           Tipo de Busqueda 142: bus_Seguros
''                   Se agregó nueva búsqueda que retrona cantidad de beneficiarios de Seguros segun los criterios seleccionados

'Global Const Version = "6.50"
'Global Const FechaModificacion = "29/12/2015"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-31053 - RH Pro (Producto) - NAC. PERU - EPS - Solicitud de Búsqueda para Cálculos de EPS [Entrega 4]
''           Tipo de Busqueda 141: bus_EPS
''                   Se agregó cuando aborta la funcion el seteo de la variable "bien" en true para que no de error de parámetro.

'Global Const Version = "6.51"
'Global Const FechaModificacion = "08/01/2016"
'Global Const UltimaModificacion = " " 'FGZ - EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-29467 - NGA- Citricos - Inconveniente en busqueda de escala
''           Tipo de Busqueda 1: Busquedas Internas. bus_interna().
''                   Se agregó la opcion de utilizar variable de conceptos precalculados en la liquidacion actual.
''                   bus_interna(). Se agregó la opcion de utilizar variable de conceptos precalculados en la liquidacion actual
''                   11  -1  Valor Concepto 00002    Valor_Concepto  objcache.00002
''           Tipo de Busqueda 3: bus_Grilla
''                   Se agrego en la búsqueda de los valores de la escala Order By vgrorden
''CAS-24204 - H&A - NACIONALIZACION BOLIVIA – Régimen RC-IVA

'Global Const Version = "6.52"
'Global Const FechaModificacion = "03/02/2016"
'Global Const UltimaModificacion = " " 'MDZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-34564 - MONASTERIO AMR - Bug en simular
''           Nivelación con versión del Simulador.

'Global Const Version = "6.53"
'Global Const FechaModificacion = "16/02/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-32751 - LA CAJA - Custom Seguros ADP [Entrega 2]
''           Tipo de Busqueda 142: bus_Seguros
''               Se corrigío el Query que busca seguros y se convirtio los parámetros de la búsqueda a byte

'Global Const Version = "6.54"
'Global Const FechaModificacion = "25/02/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-35783 - RH Pro (Producto) - ARG - NOM - Ganancias 2016 Decreto 394
''           FOR_GANANCIAS2013: Decreto 394/16 deroga decreto 1242/13. Se comento función que busca el maximo Bruto (BuscarBrutoAgosto2013)
''                               Se corrigio error en la búsqueda de ganancia en el item 56 estaba usando el recorset con un campo que no estaba en el query.

'Global Const Version = "6.55"
'Global Const FechaModificacion = "07/03/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-33601 - RH Pro (Producto) - Peru - Criterios de Busquedas [Entrega 2]
''           Tipo de Busqueda 140: Se agregó una nueva opcion para poder configurar los años pendientes de goce devacaciones

'Global Const Version = "6.56"
'Global Const FechaModificacion = "18/03/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       CAS-36167 - RH Pro (Producto) - NOM - Ganancias - Bug Item 56
''           FOR_GANANCIAS2013: Se corrigió la formula porque tenía en cuenta el item 56 en el cáclculo de ganancia imponible y en retenciones ya efectuadas.
''       CAS-36065 - NGA BASE CITRICOS - Bug en liquidar novedades
''           Tipo de Busqueda 9: se corrigio el query que busca las novedades de otro concepto.


'Global Const Version = "6.57"
'Global Const FechaModificacion = "10/05/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       ' EAM (Sprint 88) - RH Pro - Argentina - Cambio legal Nuevo reporte Ganancias RG 3839 AFIP
''           FOR_GANANCIAS2013: Se agregó el Item 19. Si éste está configurado se calcula, pero no se tiene en cuenta en el cálculo de la ganancia Imponible
''                                se agrego la condición (prorratea = 0) para que se tenga en cuenta tambien los item cuando es una liq. final
''                               Se agregó control de Item 20 para que se tenga en cuenta en la liquidacion final o si es fin de año.

'Global Const Version = "6.58"
'Global Const FechaModificacion = "02/06/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       ' EAM (Sprint 89) - CAS-33601 - RH Pro (Producto) - Peru - Criterios de Busquedas [Entrega 3]
'''           tipo de busqueda 120: bus_CTS_Tiempo_Trabajado
'''                   Se corrigió cuando el semestre pasa de año no estaba considerando los meses
'''           tipo de busqueda 140: bus_SaldoVac_PE
'''                    Se corrigio el armado de las fechas para buscar los períodos de vacaciones para tomar los tipos de días

'Global Const Version = "6.59"
'Global Const FechaModificacion = "13/06/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       'EAM (Sprint 90) - CAS-36566 - IBT - Error en busqueda de concepto acumulador meses fijos
''           tipo de busqueda 7: bus_Acum3
''                   Se agregó control a la busqueda cuando es semestral y tiene mes de inicio configurado.
''                   Ahora chequea en el semestre que cae y en funcion de eso setea el mes de inicio.

'Global Const Version = "6.60"
'Global Const FechaModificacion = "16/06/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       'EAM (Sprint 90) - 1320 - Error en tipo de búsqueda 138 - Vacaciones vendidas
''           tipo de busqueda 138: bus_Vac_Vendidas
''                   Controla si no encuentra período de vacaciones y se corrigio el log

'Global Const Version = "6.61"
'Global Const FechaModificacion = "28/06/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       'EAM (Sprint 91) - 1941 - CAS-33601 - RH Pro (Producto) - Peru - Criterios de Busquedas [Entrega 4]
''           tipo de busqueda 120: bus_CTS_Tiempo_Trabajado
''                   Se agregó a la opcion "Proporcion 30 Dias" para que tope los días de los meses a 30 días ya que originalmente funcionaba diferente.

'Global Const Version = "6.62"
'Global Const FechaModificacion = "01/07/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
'       'EAM (Sprint 91) - 764 - CAS-37441 - GRUPO ROCIO -  ADECUACIONES – Bono por Asistencia
'           tipo de busqueda 130: bus_Acumulado_Diario
'                   Se agregó la funcionalidad para que busque las horas del acumulado diario por legajo o por documento

                                
'Global Const Version = "6.63"
'Global Const FechaModificacion = "06/07/2016"
'Global Const UltimaModificacion = " " 'MDF
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       MDF - Srpint 91 pbi 182 - Seguridad - Liquidador - Log - Se colocan los logs en la carpeta del usuario que ejecuta el proceso

                                
'Global Const Version = "6.64"
'Global Const FechaModificacion = "19/07/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       'EAM (Sprint 93) - 2646 - CAS-33601 - RH Pro (Producto) - Peru - Criterios de Búsquedas  - Corrección CTS
''           tipo de busqueda 120: ´Se volvio hacer la busqueda de la version 6.61 prque se perdieron las modificaciones

'Global Const Version = "6.65"
'Global Const FechaModificacion = "08/09/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
'       'EAM (Sprint 96) - 3531 - CAS-38678 - MONASTERIO APEX CHILE - Mejora licencias integradas
'           tipo de busqueda 82: Se corrigió la búsqueda cuando marca las licenias procesadas. Se agregó la marca al final para que no se vuelvan a considerar.

'Global Const Version = "6.66"
'Global Const FechaModificacion = "12/09/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       'EAM (Sprint 96) - 3386 - CAS-38734 - MONASTERIO BASE APEX CHILE - Búsqueda novedad de concepto
''           tipo de busqueda 9: Se corrigió la búsqueda cuando tiene mas de una novedad por estructura y el empleado no tiene alguna novedad.

'Global Const Version = "6.67"
'Global Const FechaModificacion = "07/10/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
'       'PBI 3409 - CAS-37159 - RHPro Consulting CH - Error en recálculo impuesto unico
'           Se modifico la formula de chile for_RecalcConcepto para que no haga el recálculo cuando tiene licencias médicas liquidadas en el período de recalculo.
'           Se agrego una función que busca las novedades individual, estructura y global. Se modificó para que no las busque mas de una vez (se gurarda en un recordset).
'           Se agregó una nueva funcion novgegiNew que usa los recordset de las novedades para que haga query a las bases

'Global Const Version = "6.68"
'Global Const FechaModificacion = "09/11/2016"
'Global Const UltimaModificacion = " " 'LED
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
'       (Sprint 101) - 5149 - ITRP 73054 - Raffo - MODULO SIMULACION
'           Nivelacion de version con proceso de simulacion.

'Global Const Version = "6.69"
'Global Const FechaModificacion = "11/11/2016"
'Global Const UltimaModificacion = " " 'LED
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       (Sprint 101) - 5093 - Liquidador - Creación de archivo de log
''           Creacion del archivo de log, si falla la conexion.

'Global Const Version = "6.70"
'Global Const FechaModificacion = "15/12/2016"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       (Sprint 103) - 5784 - ITRP - 93331 - GC - Edenor - Calculo de BAE
''           Se Agregaron las funciones enviadas por Edenor para el cálculo de Edenor
''       (Sprint 103) - 5554 - CAS- 37159 - Santander CL - Error en reliq SIS + error cesantia emp reliq
''           Se corrigió error cuando se hace el recalculo y no utilizaban el parametro de licencias médicas "for_RecalcConcepto()"


'Global Const Version = "6.71"
'Global Const FechaModificacion = "28/12/2016"
'Global Const UltimaModificacion = " " 'FGZ
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
'       (Sprint 104) PBI 6006 - ITRP - 4532815 - NGA - Farmacity - Liquidador - msg de error en log
'              Los msg de error se deben mostrar siempre

'Global Const Version = "6.72"
'Global Const FechaModificacion = "18/01/2017"
'Global Const UltimaModificacion = " " 'LED
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       (Sprint 106) PBI 6248 - Performance - Global - Payroll Process
''              Se quito la creacion del archivo Liq_Querys_OpenR y escrituras en el archivo mismo.

'Global Const Version = "6.73"
'Global Const FechaModificacion = "07/02/2017"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       (Sprint 106) PBI 6171 - ITRP - 4593941 - Cambio legal Ganancias 2017
''             for_Ganancias2017: Se realizo una nueva formula de gananacia para argentina.
''                   Para zona Patagónica se incrementan un 22% el valor de las escalas del punto 1 informado en el parámetro 1008 (Item 10,11,16,17,31)
''       (Sprint 106) PBI 6245 - ITRP - 4631111 - Uruguay -Cambio legal - IRPF 2017
''             for_irpf: Se modifico la forma en que se obtiene la taza de deducciones a aplicar. ahora se entra en escala con los bcp del empleado utilizado en la escala de renta
''       ITRP - 4692247 - Farmacity - Error en Busqueda Acum meses fijos
''               Se corrigio la busqueda Acum meses fijos en el cálculo de máximo y mínimo con proceso actual no estaba calculando bien los montos y retonrnaba como máximo y mínimo un valor incorrecto.

'Global Const Version = "6.74"
'Global Const FechaModificacion = "10/02/2017"
'Global Const UltimaModificacion = " " 'EAM
'Global Const UltimaModificacion1 = " "
'Global Const UltimaModificacion2 = " " '
''       (Sprint 107) PBI 4081 - ITRP - 70415 - GC - Liquidacion - Calculo de Mopre
''             Se modifico el tope de Mopre tenindo en cuenta la cantidad de días que imputan en el cálculo
''       (Sprint 107) PBI 6735 - ITRP - 4779473 - MEDICUS - 79473 Sistema no procesa a partir del cambio en liquidador 6.73
''             Se modifico el order by haciendo referencia a la tabla que ordena porque en oracle da error.
''       (Sprint 107) PBI 6744 - ITRP - 4606019 Error en recálculo de imp unico- RH PRO CHILE [SB]
''               Se corrige control cuando busca el acumulador de licencias medicas y es nulo

Global Const Version = "6.75"
Global Const FechaModificacion = "21/02/2017"
Global Const UltimaModificacion = " " 'EAM
Global Const UltimaModificacion1 = " "
Global Const UltimaModificacion2 = " " '
'       (Sprint 108) PBI 6626 - ITRP - 4747498 - RH Pro - Cambio Legal Ganancias - Alquileres
'               ITRP - 4747498 - RH Pro - Cambio Legal Ganancias - Alquileres

''--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------PENDIENTE DE LIBERAR



Public Function ValidarV(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que determina si el proceso esta en condiciones de ejecutarse.
' Autor      : FGZ
' Fecha      : 05/08/2009
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True

Select Case TipoProceso
Case 3: 'Liquidador
    If Version >= "4.17" Then
        'Tabla nueva
        
        'CREATE TABLE [dbo].[novretro](
        '      [nretronro] [int] IDENTITY(1,1) NOT NULL,
        '      [ternro] [int] NOT NULL,
        '      [concnro] [int] NOT NULL,
        '      [nretromonto] [decimal](19, 4) NULL,
        '      [nretrocant] [int] NULL,
        '      [pliqnro] [int] NULL,
        '      [simpronro] [int] NULL,
        '      [pronro] [int] NULL,
        '      [pronropago] [int] NULL
        ') ON [PRIMARY]
        
        Texto = "Revisar que exista tabla novretro y su estructura sea correcta."
        
        StrSql = "Select nretronro,ternro,concnro,nretromonto,nretrocant,pliqnro,simpronro,pronro,pronropago FROM novretro WHERE ternro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If
    
    If Version >= "4.24" Then
        'Tabla nueva
        
        'CREATE TABLE [dbo].[emp_fr_comp](
        '    [frannro] [int] IDENTITY(1,1) NOT NULL,
        '    [ternro] [int] NOT NULL,
        '    [fecha] [date] NOT BULL,
        '    [unidad] [int] NOT NULL,
        '    [Cantidad] Not [decimal](19, 4),
        '    [comentario] [varchar](200) NULL,
        '    liq smallint NOT NULL default 0,
        '    pronro int NULL
        ') ON [PRIMARY]
        'GO
        
        Texto = "Revisar que exista tabla emp_fr_comp y su estructura sea correcta."
        
        StrSql = "Select ternro FROM emp_fr_comp WHERE ternro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If

    If Version >= "5.05" Then
        'campos nuevos
        '   convenios.divisorhora
        Texto = "Revisar campo divisorhora en la tabla convenios."
        
        StrSql = "Select convenios.divisorhora FROM convenios WHERE estrnro = 1"
        OpenRecordset StrSql, rs
        
        '   embargo.reghorario
        Texto = "Revisar campo reghorario en la tabla embargo."
        
        StrSql = "Select embargo.reghorario FROM embargo WHERE embnro = 1"
        OpenRecordset StrSql, rs
        
        
        'Tabla regimenhorario
        'CREATE TABLE [dbo].[regimenHorario](
        '    -- Llave
        '    [reginro] [int] IDENTITY NOT NULL,
        '    [estrnro] int NOT NULL,
        '    [HorasDia] int NULL,
        '    [HorasMes] int NULL
        ')
        Texto = "Revisar que exista tabla regimenHorario y su estructura sea correcta."
        
        StrSql = "Select estrnro FROM regimenHorario WHERE estrnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If
    
    If Version >= "5.11" Then
        'campos nuevos
        '   embargo.pronro

        Texto = "Revisar campo pronro en la tabla embargo."
        
        StrSql = "Select embargo.pronro FROM embargo WHERE embnro = 1"
        OpenRecordset StrSql, rs
        
        '   embargo.embestant
        Texto = "Revisar campo embestant en la tabla embargo."
        
        StrSql = "Select embargo.embestant FROM embargo WHERE embnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If

    If Version >= "5.13" Then
        'Tabla nueva
        
        'CREATE TABLE Sim_curva (
        '    curnro int PRIMARY KEY IDENTITY(1,1) NOT NULL,
        '    curdesc varchar(20) NOT NULL,
        '    curmes1 dec(19,4) NULL DEFAULT 0,
        '    curmes2 dec(19,4) NULL DEFAULT 0,
        '    curmes3 dec(19,4) NULL DEFAULT 0,
        '    curmes4 dec(19,4) NULL DEFAULT 0,
        '    curmes5 dec(19,4) NULL DEFAULT 0,
        '    curmes6 dec(19,4) NULL DEFAULT 0,
        '    curmes7 dec(19,4) NULL DEFAULT 0,
        '    curmes8 dec(19,4) NULL DEFAULT 0,
        '    curmes9 dec(19,4) NULL DEFAULT 0,
        '    curmes10 dec(19,4) NULL DEFAULT 0,
        '    curmes11 dec(19,4) NULL DEFAULT 0,
        '    curmes12 dec(19,4) NULL DEFAULT 0,
        ')
        
        Texto = "Revisar que exista tabla sim_curva y su estructura sea correcta."
        
        StrSql = "Select curnro FROM sim_curva WHERE curnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If

    If Version >= "5.21" Then
        'Tabla nueva
        
        'CREATE TABLE [dbo].[lic_pagas](
        '[ternro] [int] NOT NULL,
        '[fecha] [datetime] NOT NULL,
        '[liq] [smallint] NOT NULL default 0,
        '[emp_licnro] [int] NOT NULL,
        '[concnro] [int] NULL,
        '[simpronro] [int] NULL,
        '[liqpronro] [int] NULL
        ') ON [PRIMARY]
        'GO
        Texto = "Revisar que exista tabla lic_pagas y su estructura sea correcta."
        
        StrSql = "Select ternro, fecha,liq,emp_licnro,simpronro,liqpronro FROM lic_pagas WHERE ternro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If


    If Version >= "5.23" Then
        'Tabla nueva
        
        'CREATE TABLE [dbo].[liq_comision](
        '    [ternro] [int] NOT NULL,
        '    [fecha] [datetime] NOT NULL,
        '    [concnro] [int] NULL,
        '    [tpanro] [int] NULL,
        '    [thnro] [int] NULL,
        '    [mpt] [decimal](19, 4) NULL,
        '    [tht] [decimal](19, 4) NULL,
        '    [th] [decimal](19, 4) NULL,
        '    [pronro] [int] NULL
        ') ON [PRIMARY]
        'GO
        Texto = "Revisar que exista tabla liq_comision y su estructura sea correcta."
        
        StrSql = "Select ternro,fecha,concnro,tpanro,thnro,mpt,tht,th,pronro FROM liq_comision WHERE ternro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If


    If Version >= "5.33" Then
        'campos nuevos
        '   alter table gastos alter column  gaspagado int not null
        '   alter table gastos add pronro int null

        
        'Tabla gastos completa
        'CREATE TABLE [dbo].[gastos](
        '    [gasnro] [int] IDENTITY(1,1) NOT NULL,
        '    [gasdesabr] [varchar](100) NULL,
        '    [proyecnro] [int] NULL,
        '    [monnro] [int] NULL,
        '    [gasvalor] [decimal](19, 4) NULL,
        '    [ternro] [int] NULL,
        '    [provnro] [int] NULL,
        '    [gasfechaida] [datetime] NULL,
        '    [gashoraida] [varchar](4) NULL,
        '    [gasfechavuelta] [datetime] NULL,
        '    [gashoravuelta] [varchar](4) NULL,
        '    [gasrevisadopor] [varchar](50) NULL,
        '    [gaspagacliente] [smallint] NULL,
        '    [gaspagado] [int] NOT NULL,
        '    [tipgasnro] [int] NOT NULL,
        '    [pronro] [int] NULL
        ') ON [PRIMARY]
        'GO
        
        Texto = "Revisar campo pronro en la tabla gastos."
        
        StrSql = "Select gastos.pronro FROM gastos WHERE gasnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If



    If Version >= "5.35" Then
        'tablas nuevas liq_emputil y liq_emputil_det
       
        'Tablas completas
        'CREATE TABLE [dbo].[liq_emputil](
        '      [utilnro] [int] IDENTITY(1,1) NOT NULL,
        '      [estrnro] [int] NOT NULL,
        '      [anio] [int] NOT NULL,
        '      [mes] [int] NOT NULL,
        '      [empcantempleados] [int] NOT NULL,
        '      [empremempleados] [decimal](19, 4) NOT NULL,
        '      [empdiastrabempleados] [int] NOT NULL,
        '      [emprenta] [decimal](19, 4) NULL,
        '      [empperdidas] [decimal](19, 4) NULL,
        '      [basecalculoutil] [decimal](19, 4) NOT NULL,
        '      [empporcpart] [decimal](19, 4) NOT NULL,
        '      [cargaFliaMonto] [decimal](19, 4) NULL,
        '      [cargaFliaCant] [int] NULL,
        '      [Bpronro] Not [Int]
        ') ON [PRIMARY]
        'GO
        '
        'CREATE TABLE [dbo].[liq_emputil_det](
        '      [utilnro] [int] NOT NULL,
        '      [utildetnro] [int] IDENTITY(1,1) NOT NULL,
        '      [ternro] [int] NOT NULL,
        '      [terrem] [decimal](19, 4) NOT NULL,
        '      [terdiastrab] Not [Int]
        ') ON [PRIMARY]
        'GO
        
        
        Texto = "Revisar tabla liq_emputil"
        StrSql = "select utilnro,estrnro,anio,mes,empcantempleados,empremempleados,empdiastrabempleados,emprenta,empperdidas,basecalculoutil,empporcpart,cargaFliaMonto,cargaFliaCant,bpronro from liq_emputil WHERE utilnro = 1"
        OpenRecordset StrSql, rs
        
                
        Texto = "Revisar tabla liq_emputil_det"
        StrSql = "select utilnro,utildetnro,ternro,terrem,terdiastrab FROM liq_emputil_det WHERE utilnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If


    If Version >= "5.36" Then
        'tabla "nueva"
        'vacvendidosnro,ternro,empleg,aprobado,fechapago, pronro, iduser,fechacarga,cantvacvendidos
       
        'Tabla completa
        'CREATE TABLE [dbo].[vacvendidos](
        '    [vacvendidosnro] [int] IDENTITY(1,1) NOT NULL,
        '    [ternro] [int] NOT NULL,
        '    [empleg] [int] NOT NULL,
        '    [aprobado] [smallint] NOT NULL,
        '    [fechapago] [datetime] NOT NULL,
        '    [pronro] [int] NULL,
        '    [iduser] [varchar](40) NULL,
        '    [fechacarga] [datetime] NOT NULL,
        '    [cantvacvendidos] [decimal](18, 4) NOT NULL
        ') ON [PRIMARY]
        'GO

        
        Texto = "Revisar tabla vacvendidos"
        StrSql = "select vacvendidosnro,ternro,empleg,aprobado,fechapago, pronro, iduser,fechacarga,cantvacvendidos from vacvendidos WHERE vacvendidosnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If


    If Version >= "5.37" Then
        'nuevos campos en tabla liq_emputil_det
       
        'Tabla completa
        'CREATE TABLE [dbo].[liq_emputil_det](
        '      [utilnro] [int] NOT NULL,
        '      [utildetnro] [int] IDENTITY(1,1) NOT NULL,
        '      [ternro] [int] NOT NULL,
        '      [terrem] [decimal](19, 4) NOT NULL,
        '      [terdiastrab] Not [Int],
        '      [famrem] [decimal](19, 2) NULL,
        '      [famdias] [int] NULL
        ') ON [PRIMARY]
        'GO
        
        Texto = "Revisar campos famrem y famdias de la tabla liq_emputil_det"
        StrSql = "select famrem, famdias FROM liq_emputil_det WHERE utilnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If

    If Version >= "5.40" Then
        'campos nuevos
        '   ALTER TABLE Gastos ADD gasretro smallint not null default(0)
        '   ALTER TABLE Gastos ADD pliqdesde int null
        '   ALTER TABLE Gastos ADD pliqhasta int null
        
        'Tabla gastos completa
        'CREATE TABLE [dbo].[gastos](
        '    [gasnro] [int] IDENTITY(1,1) NOT NULL,
        '    [gasdesabr] [varchar](100) NULL,
        '    [proyecnro] [int] NULL,
        '    [monnro] [int] NULL,
        '    [gasvalor] [decimal](19, 4) NULL,
        '    [ternro] [int] NULL,
        '    [provnro] [int] NULL,
        '    [gasfechaida] [datetime] NULL,
        '    [gashoraida] [varchar](4) NULL,
        '    [gasfechavuelta] [datetime] NULL,
        '    [gashoravuelta] [varchar](4) NULL,
        '    [gasrevisadopor] [varchar](50) NULL,
        '    [gaspagacliente] [smallint] NULL,
        '    [gaspagado] [int] NOT NULL,
        '    [tipgasnro] [int] NOT NULL,
        '    [pronro] [int] NULL,
        '    [gasretro] [smallint] NOT NULL default(0),
        '    [pliqdesde] [int] NULL,
        '    [pliqhasta] [int] NULL
        ') ON [PRIMARY]
        'GO
        
        'FGZ - 26/07/2013 - estaba mal el comentario, los campos agregados no son esos
        'Texto = "Revisar campo pronro en la tabla gastos."
        Texto = "Revisar campo gasretro, pliqdesde, pliqhasta en la tabla gastos."
        
        StrSql = "Select gasretro, pliqdesde, pliqhasta FROM gastos WHERE gasnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If


    If Version >= "5.51" Then
        'campos nuevos
        '   ALTER TABLE batch_tipproc ADD estadistica [smallint] NULL default (0)
        
        'Tabla nueva
        'CREATE TABLE [dbo].[His_batch_proceso_est](
        '    [bpronro] [int] NOT NULL,
        '    [bpronroori] [int] NULL,
        '    [version] [varchar](10) NULL,
        '    [debug] [smallint] NULL,
        '    [andet] [smallint] NULL,
        '    [bdlocal] [smallint] NULL,
        '    [cantlectbd] [int] NULL,
        '    [cantemp] [int] NULL,
        '    [cantconc] [int] NULL,
        '    [cantacu] [int] NULL,
        '    [cantbusq] [int] NULL,
        '    [cantbusqint] [int] NULL,
        '    [cantbusqnovg] [int] NULL,
        '    [cantbusqnove] [int] NULL,
        '    [cantbusqnovi] [int] NULL,
        '    [cantconcaju] [int] NULL,
        '    [segundos] [int] NULL,
        '    [promemp] [decimal](19, 4) NULL
        ') ON [PRIMARY]
        'GO
      
        Texto = "Revisar campo estadistica en la tabla batch_tipproc."
        StrSql = "Select estadistica FROM batch_tipproc WHERE btprcnro = 3"
        OpenRecordset StrSql, rs
        
        
        Texto = "Revisar que exista y tenga permisos la tabla His_batch_proceso_est."
        StrSql = "Select * FROM His_batch_proceso_est WHERE bpronro = 1"
        OpenRecordset StrSql, rs
        
        
        V = True
    End If


    If Version >= "5.72" Then
        'Tabla nueva
        'CREATE TABLE [dbo].[nov_dist](
        '    [nedistnro] [int] IDENTITY(1,1) NOT NULL,
        '    [novnro] [int] NOT NULL,                -- FK (novemp o novaju)
        '    [auto] [smallint] NOT NULL default 0,
        '    [tiponov] [int] NOT NULL default 1,     -- {1 novemp, 2 novaju}
        '    [concnro] [int] NOT NULL,               -- FK {concepto}
        '    [tpanro] [int] NULL,                    -- FK {parametro del concepto} -- no se si es FK porque si es novaju ==> va 0
        '    [Masinro] Not [Int], --FK(mod_asiento)
        '    [tenro] [int] NOT NULL,                 -- FK (estructura) ---- no se si es FK porque necesitamos ponerle 0 cuando no tiene distr
        '    [Estrnro] Not [Int], --FK(estructura)
        '    [tenro2] [int] NULL,                    -- FK (estructura)
        '    [estrnro2] [int] NULL,                  -- FK (estructura)
        '    [tenro3] [int] NULL,                    -- FK (estructura)
        '    [estrnro3] [int] NULL                   -- FK (estructura)
        ') ON [PRIMARY]
        'GO


        'CREATE TABLE [dbo].[concepto_dist](
        '    [Ternro] Not [Int], --FK(Tercero)
        '    [ConcNro] Not [Int], --FK(Concepto)
        '    [pronro] Not [Int], --FK(Proceso)
        '    [Masinro] Not [Int], --FK(mod_asiento)
        '    [tenro] [int] NULL,         -- FK (estructura)
        '    [estrnro] [int] NULL,       -- FK (estructura)
        '    [tenro2] [int] NULL,        -- FK (estructura)
        '    [estrnro2] [int] NULL,      -- FK (estructura)
        '    [tenro3] [int] NULL,        -- FK (estructura)
        '    [estrnro3] [int] NULL,      -- FK (estructura)
        '    [porcentaje] [decimal](5, 2) NOT NULL Default(100),
        '    [Monto] Not [decimal](19, 4)
        ') ON [PRIMARY]
        'GO
        
        Texto = "Revisar que exista y tenga permisos la tabla nov_dist "
        StrSql = "Select * FROM nov_dist WHERE nedistnro = 1"
        OpenRecordset StrSql, rs
        
        Texto = "Revisar que exista y tenga permisos la tabla concepto_dist "
        StrSql = "Select * FROM concepto_dist WHERE Ternro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If


    If Version >= "5.72" Then
        'campos nuevos
        
        'Tablas nueva
        
        'Escala de comisiones
        'CREATE TABLE [dbo].[escala_comision](
        '    [esccomnro] [int] IDENTITY(1,1) NOT NULL,
        '    [esccomdesabr] [varchar](100) NOT NULL,     -- Descripcion Abreviada
        '    [esccomdesext] [varchar](500) NULL,         -- Descripcion Extendida
        '    [version] [varchar](10) NULL,               -- Version
        '    [activa] [smallint] NOT NULL default(-1),   -- Activa (True / False)
        '    [fecdesde] [datetime] NOT NULL,             -- Fecha desde de vigencia
        '    [fechasta] [datetime] NULL                  -- Fecha hasta de vigencia
        ') ON [PRIMARY]
        'GO
        
        'productos(conceptos) asociados a la escala de comisiones
        'CREATE TABLE [dbo].[escala_comision_conc](
        '    [esccomnro] [int] NOT NULL,                         -- FK a tabla escala_comision
        '    [concnro] [int] NOT NULL                            -- Concepto (Producto). FK a tabla concepto
        ') ON [PRIMARY]
        'GO
        
        'Lineas de productos(Estructuras) asociados a la escala de comisiones
        'CREATE TABLE [dbo].[escala_comision_estr](
        '    [esccomnro] [int] NOT NULL,                         -- FK a tabla escala_comision
        '    [tenro] [int] NOT NULL,                             -- Tipo de estructura. FK a tabla tipo_estructura
        '    [estrnro] [int] NOT NULL                            -- Estructura. FK a tabla estructura
        ') ON [PRIMARY]
        'GO
        
        'detalle de comisiones por linea(estructura) y producto(concepto)
        'CREATE TABLE [dbo].[escala_comision_det](
        '    [esccomnro] [int] NOT NULL,                         -- FK a tabla escala_comision
        '    [esccomdetnro] [int] IDENTITY(1,1) NOT NULL,        -- identidad del detalle
        '    [tenro] [int] NOT NULL,                             -- Tipo de estructura. FK a tabla tipo_estructura
        '    [estrnro] [int] NOT NULL,                           -- Estructura. FK a tabla estructura
        '    [concnro] [int] NOT NULL,                           -- Concepto (Producto). FK a tabla concepto
        '    [pordesde] [decimal](19, 4) NOT NULL,               -- Porcentaje desde. Rango
        '    [porhasta] [decimal](19, 4) NOT NULL,               -- Porcentaje Hasta. Rango
        '    [comision] [decimal](19, 4) NOT NULL,               -- Valor de Comision
        '    [comision2] [decimal](19, 4) NOT NULL default(0),   -- Valor de Comision 2
        '    [comision3] [decimal](19, 4) NOT NULL default(0),   -- Valor de Comision 3
        '    [comision4] [decimal](19, 4) NOT NULL default(0)    -- Valor de Comision 4
        ') ON [PRIMARY]
        'GO
      
        Texto = "Revisar que exista y tenga permisos la tabla escala_comision."
        StrSql = "Select * FROM escala_comision WHERE esccomnro = 1"
        OpenRecordset StrSql, rs
        
        Texto = "Revisar que exista y tenga permisos la tabla escala_comision_conc."
        StrSql = "Select * FROM escala_comision_conc WHERE esccomnro = 1"
        OpenRecordset StrSql, rs
        
        Texto = "Revisar que exista y tenga permisos la tabla escala_comision_estr."
        StrSql = "Select * FROM escala_comision_estr WHERE esccomnro = 1"
        OpenRecordset StrSql, rs
        
        Texto = "Revisar que exista y tenga permisos la tabla escala_comision_det."
        StrSql = "Select * FROM escala_comision_det WHERE esccomnro = 1"
        OpenRecordset StrSql, rs
        
        
        V = True
    End If


    If Version >= "5.77" Then
        'campos nuevos
        
        'Tablas nueva
        'Cabecera de Paros
        'CREATE TABLE parcab(
        '    parnro int IDENTITY (1, 1) NOT NULL,
        '    pardesabr varchar(50) NOT NULL,
        '    parfecdesde datetime NOT NULL,
        '    parfechasta datetime NULL,
        '    pardiacomp smallint NOT NULL,
        '    parhordesde varchar(4),
        '    parhorhasta VarChar(4)
        ') ON [PRIMARY]
        'GO
        
        'Detalle del Paro
        'CREATE TABLE pardet(
        '    parnro int NOT NULL,
        '    ternro int NOT NULL,
        '    detmonto decimal(19, 4),
        '    detcanthor decimal(19, 4),
        '    pronro int
        ') ON [PRIMARY]
        'GO
        
        'Sectores Relacionados al Paro
        'CREATE TABLE parrel(
        '    parnro int,
        '    secnro int
        ') ON [PRIMARY]
        'GO
        
      
        Texto = "Revisar que exista y tenga permisos la tabla parcab"
        StrSql = "Select * FROM parcab WHERE parnro = 1"
        OpenRecordset StrSql, rs
        
        Texto = "Revisar que exista y tenga permisos la tabla pardet."
        StrSql = "Select * FROM pardet WHERE parnro = 1"
        OpenRecordset StrSql, rs
        
        Texto = "Revisar que exista y tenga permisos la tabla parrel."
        StrSql = "Select * FROM parrel WHERE parnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If

    If Version >= "5.90" Then
        'Tablas nueva
        
        'CREATE TABLE fases_preaviso(
        'fasnro int NOT NULL,
        'aviso  int NOT NULL default 0,
        'descuenta int NOT NULL default 0,
        'fecha_preaviso Not DateTime
        ')
        Texto = "Revisar que exista y tenga permisos la tabla fases_preaviso"
        StrSql = "Select * FROM fases_preaviso WHERE fasnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If

    If Version >= "5.95" Then
        'Campo nuevo
        
        'alter table gti_novedad add fechaprocesamiento datetime
        Texto = "Revisar campo fechaprocesamiento en la tabla gti_novedad."
        
        StrSql = "Select fechaprocesamiento FROM gti_novedad WHERE gnovnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If


    If Version >= "6.13" Then
        'Campo nuevo
        
        'ALTER TABLE [pre_cuota] ADD [cuocancelado] [decimal](19, 4) NULL
        Texto = "Revisar campo cuocancelado en la tabla pre_cuota."
        
        StrSql = "Select cuocancelado FROM pre_cuota WHERE prenro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If
    
    If Version >= "6.39" Then
        'Campo nuevo de venta de vacaciones
                
        Texto = "Revisar campo cuocancelado en la tabla vacvendidos."
        
        StrSql = "Select automatico FROM vacvendidos where vacvendidosnro=1"
        OpenRecordset StrSql, rs
        
        V = True
    End If
    
    If Version >= "6.41" Then
        'Campo Nuevo
        'venc,vacnro
        'ALTER TABLE [vacvendidos] ADD [venc] [int] NOT  NULL
        'ALTER TABLE [vacvendidos] ADD [vacnro] [int] NOT  NULL
                
        Texto = "Revisar tabla vacvendidos campos venc,vacnro"
        StrSql = "select venc,vacnro from vacvendidos WHERE vacvendidosnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If


    ''*****************************************************************
    ''Hay que agregar este control cuando se libere el caso 22808
    'If Version >= "5.78" Then
    '    'campos nuevos
    '    'ALTER TABLE novaju ADD nadist smallint default 0
    '    'ALTER TABLE novemp ADD nedist smallint default 0
    '    'ALTER TABLE novemp ADD gpanro [int] NULL
    '
    '    'ALTER TABLE sim_novaju ADD nadist smallint default 0
    '    'ALTER TABLE sim_novemp ADD nedist smallint default 0
    '    'ALTER TABLE sim_novemp ADD gpanro [int] NULL
    '
    '    Texto = "Revisar campo nadist en la tabla novaju."
    '    StrSql = "Select nadist FROM novaju WHERE nanro = 1"
    '    OpenRecordset StrSql, rs
    '
    '    Texto = "Revisar campos nadist y gpanro en la tabla novemp."
    '    StrSql = "Select nadist,gpanro FROM novemp WHERE nenro = 1"
    '    OpenRecordset StrSql, rs
    '
    '    Texto = "Revisar campo nadist en la tabla sim_novaju."
    '    StrSql = "Select nadist FROM sim_novaju WHERE nanro = 1"
    '    OpenRecordset StrSql, rs
    '
    '    Texto = "Revisar campos nadist y gpanro en la tabla sim_novemp."
    '    StrSql = "Select nadist,gpanro FROM sim_novemp WHERE nenro = 1"
    '    OpenRecordset StrSql, rs
    '
    '
    '    V = True
    'End If
    ''*****************************************************************
    
'Cambio en busqueda de embargos bus_embargos
Case Else:
    Texto = "version correcta"
    V = True
End Select



    ValidarV = V
Exit Function

ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function

