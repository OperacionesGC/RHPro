Attribute VB_Name = "mdlGTI_PRC_30"
Option Explicit

'Version: 2.01. Integracion con las politicas customizadas de Halliburton

'Const Version = 2.02    'Politca 60v2. Nulos en Partes de asignacion horaria.
'Const FechaVersion = "07/10/2005"

'Const Version = 2.03    'Justificaciones de Licencias autorizadas solamente.
'Const FechaVersion = "17/11/2005"
''Procedimientos modificados:
'        'Generar_Dia_Justificacion
'        'Generar_Justif_Parcial
'        'Politica 480

'Const Version = 2.04    'Politica 60. Validacion de dia anterior no nulo
'Const FechaVersion = "05/12/2005"

'Const Version = 2.05    'Generar Horario normal. Modificacion cuando arma el desgloce fijo y se pasa de dia
'Const FechaVersion = "06/12/2005"

'Const Version = 2.06    'mas detalle de log
'Const FechaVersion = "07/12/2005"

'Const Version = 2.07    'politica 14
'Const FechaVersion = "13/12/2005"

'Const Version = 2.08    'politica 70
'Const FechaVersion = "14/12/2005"

'Const Version = 2.09    'Controla que el periodo de gti no este cerrado
'Const FechaVersion = "19/12/2005"

'Const Version = 2.11     'Justificaciones parciales
'Const FechaVersion = "26/12/2005"

'Const Version = 2.12     'Cambiar_Horas_hasta. Habia un error cuando el tipo de desgloce era fijo / relativo. cargaba la hora hasta pero no cargaba la fecha hasta
'Const FechaVersion = "27/12/2005"

'Const Version = 2.13    'Generar Horario normal. Modificacion cuando arma el desgloce relativo y se pasa de dia y el turno es nocturno
'Const FechaVersion = "10/01/2006"
'cuando la FP es relativa y va de posterior a la salida y anterior a la entrada
'tiene problemas con los turnos nocturnos ya que la salida pasa de dia ==>
'hay que hacer el control desde afuera y pasar de dia al calculo de la fecha hasta

'ademas

'cuando el turno es nocturno y se pasa de dia cuando busca el desgloce no debe buscar el desgloce
'del siguiente dia para las registraciones que vienen desde el dia anterior
'Ej. Turno de 22 a 06
'     registra de 22 a 06
'supongamos que el dia es lunes y ademas los desgloces para los dias lunes a viernes son iguales pero sabado y domingo son francos
' cuando genere el horario normal va a partir las registraciones y hará de 22 a 22 y de 00 a 06
' cuando pase a la parte de 00 a 06 no debe buscar el desgloce del dia siguiente sino que debe quedarse con el desgloce del dia que se esta procesando
' porque la ejecucion podria fallar dependiendo del dia que se procese. Si por casualidad el dia siguiente tiene el mismo desgloce que el dia que se esta procesando
' no va a haber problemas pero ...
' si el siguiente dia es no laborable ==> no va a generar las horas como corresponde Ejemplo tipico cuando proceso un viernes

'Const Version = 2.14    'Generar Horario normal. Modificacion cuando arma el desgloce relativo y se pasa de dia y el turno es nocturno.
'                        'Generar Horario Libre.  Nunca buscar el desgloce del dia siguiente cuando pasa de dia la registracion.
'Const FechaVersion = "11/01/2006"


'Const Version = 2.15
'Const FechaVersion = "04/04/2006"
''Modificaciones para Halliburton:
'    'Politica 50. Nueva version 5. Tope de horas al dia de la registracion.
'    'Politica 220. Nueva politica y con version 1. Halliburton y estandar.
'    'ConvertirHorasNormales2. Nuevo procedimiento = a ConvertirHorasNormales con una condicion menos
'    '                         If HorasNormales >= 8 And NroTh = 3 Then
'    'Politica 700. Nueva Politica y con version 1. Halliburton y estandar.
'    'Politica 450. Se pasó la politica a configurable con un solo parametro y ademas
'    'se hizo versiones 1(estandar) y 2(Custom Halliburton).

'Const Version = 2.16
'Const FechaVersion = "05/04/2006"
''Modificaciones:
'    'sub PRC30 - en Informix hay que poner el from en el delete. DELETE FROM gti_proc_emp


'Const Version = 2.17
'Const FechaVersion = "11/04/2006"
''Modificaciones:
'    'Politca 70 - Se agregó el subn = 4


'Const Version = 2.18
'Const FechaVersion = "12/04/2006"
''Modificaciones:
'    'Generar_Horario_Libre y Generar_Horario_Normal, Generar_Horario_Libre_C1 y Generar_Horario_Normal_C1
'        'si el desgloce es por cantidad de horas hay que guardar todos los pares porque acumula antes de guardar
'        'y se pisan los nro de las registraciones a marcar como procesadas

'Const Version = 2.19
'Const FechaVersion = "18/04/2006"
''Modificaciones:
'    'Justificaciones parciales fijas, quedaban algunos casos que no andaban

'Const Version = 2.2
'Const FechaVersion = "19/04/2006"
''Modificaciones:
'    'Generar_Horario_Libre y Generar_Horario_Libre_C1
'        'Justificaciones parciales fijas, turnos libres
'        'no se estaban teniendo en cuenta las horas justificadas en la suma de las horas del dia
'    'Politica 480
'        'No mostraba bien las horas desde hasta cuando el turno es libre
        
'Const Version = 2.21
'Const FechaVersion = "27/04/2006"
''Modificaciones:
'    'Politica 450
'        'Habia un error en el insert. nroreg,nroreg
'    'Politica 1000
'        'nueva version para halliburton que reemplazaria a la anterior
        
'Const Version = 2.22
'Const FechaVersion = "28/04/2006"
''Modificaciones:
'    'Politica 1000
'        'nueva version para halliburton que reemplazaria a la anterior
        
'Const Version = 2.23
'Const FechaVersion = "04/05/2006"
''Modificaciones:
'    'Politica 1000
'        'Agregados de log
'    'Politica 1001
'        'Agregados de log
        
'Const Version = 2.24
'Const FechaVersion = "04/05/2006"
''Modificaciones:
'    'Politica 1000
'        'Agregados de log + habia un bucle infinito
        
'Const Version = 2.25
'Const FechaVersion = "09/05/2006"
''Modificaciones:
'    'Politica 1000, 220 y 540
'        'Agregados de logs
        
'Const Version = 2.26
'Const FechaVersion = "12/05/2006"
''Modificaciones:
'        'Agregados de logs en general cuando inserta en gti_hor_cumplido
        
        
'Const Version = 2.27
'Const FechaVersion = "22/05/2006"
''Modificaciones: FGZ
''           Politica 450: no estaba buscando bien el desgloce de horas porque no estaba asociando el turno
'        'Antes
'        '        StrSql = "SELECT * FROM gti_desgdia  WHERE fpgonro= " & Nro_fpgo & _
'        '        " AND desgdtipo = " & diatipo & " AND desgdnrodia = " & Weekdia & _
'        '        " ORDER BY desgdnrodia ASC, desgdtipo ASC"
'
'        'Ahora
'        '        StrSql = "SELECT * FROM gti_desgdia  WHERE fpgonro= " & Nro_fpgo & _
'        '        " AND subturnro = " & Nro_Subturno & " AND desgdtipo = " & diatipo & " AND desgdnrodia = " & Weekdia & _
'        '        " ORDER BY desgdnrodia ASC, desgdtipo ASC"
'
'        'Ademas
'        'Habia un problema si no es feriado ni libre (es decir: esFeriado = False y Dia_Libre = False)
'        'Diatipo termina valiendo 2 (es decir Feriado), y te paga cualquier configuracion de HC que tengas como fijo o fijo sin registracion de un Feriado en un dia Libre.

'Const Version = 2.28
'Const FechaVersion = "09/06/2006"
''Modificaciones: FGZ
''           Politica 1001 (Generacionde horario Libre): problemas en turnos trasnoche, no desgloza bien la FP cuando hay un feriado de pormedio
''------------------------------------------------------------------------------------------------
''Caso Nro    Turno   Tipo Dia Inicial    Tipo Dia Final  Cambia FP   Turno   Politica que ejecuta
''------------------------------------------------------------------------------------------------
''1          NORMAL  LABORABLE           LABORABLE           NO      NORMAL      1000
''2          NORMAL  LABORABLE           FRANCO              NO      NORMAL      1000
''3          NORMAL  LABORABLE           FERIADO             SI      NORMAL      1000
''4          LIBRE   LABORABLE           LABORABLE           NO      LIBRE       1001
''5          LIBRE   LABORABLE           FRANCO              NO      LIBRE       1001
''6          LIBRE   LABORABLE           FERIADO             SI      LIBRE       1001
''7          LIBRE   FRANCO              LABORABLE           NO      LIBRE       1001
''8          LIBRE   FRANCO              FRANCO              SI      LIBRE       1001
''9          LIBRE   FRANCO              FERIADO             SI      LIBRE       1001
''10         LIBRE   FERIADO             LABORABLE           SI      LIBRE       1001
''11         LIBRE   FERIADO             FRANCO              SI      LIBRE       1001
''12         LIBRE   FERIADO             FERIADO             SI      LIBRE       1001
       
        
'Const Version = 2.29
'Const FechaVersion = "09/06/2006"
''Modificaciones: FGZ
''    'generacion de horario normal: si tiene una sola registracion llama a la politica 40
''                                   si tiene una sola registracion = general el horario normal
        
        
'Const Version = 2.3
'Const FechaVersion = "12/06/2006"
''Modificaciones: FGZ
''    'generacion de horario normal: el siguiente caso no andaba
'''------------------------------------------------------------------------------------------------
'''Caso Nro    Turno   Tipo Dia Inicial    Tipo Dia Final  Cambia FP   Turno   Politica que ejecuta
'''------------------------------------------------------------------------------------------------
''3          NORMAL  LABORABLE           FERIADO             SI      NORMAL      1000
''9          LIBRE   FRANCO              FERIADO             SI      LIBRE       1001
''10         LIBRE   FERIADO             LABORABLE           SI      LIBRE       1001
''11         LIBRE   FERIADO             FRANCO              SI      LIBRE       1001


'Const Version = 2.31
'Const FechaVersion = "13/06/2006"
''Modificaciones: FGZ
''    politica 40: se activó el codigo de la politica
''                y se habilitó tambien para los turnos normales porque solo estaba disponible para turnos libres

'Const Version = 2.32
'Const FechaVersion = "14/06/2006"
'Modificaciones: FGZ
''    'generacion de horario libre: el siguiente caso no andaba
'''------------------------------------------------------------------------------------------------
'''Caso Nro    Turno   Tipo Dia Inicial    Tipo Dia Final  Cambia FP   Turno   Politica que ejecuta
'''------------------------------------------------------------------------------------------------
''11         LIBRE   FERIADO             FRANCO              SI      LIBRE       1001



'Const Version = 2.33
'Const FechaVersion = "28/07/2006"
''Modificaciones: FGZ
'''    'Politica 70 - Cuando busca las registraciones no tiene en cuenta el estado, ahora si


'Const Version = 2.34
'Const FechaVersion = "01/08/2006"
''Modificaciones: FGZ
''           Politica 1001 (Generacionde horario Libre): Version

'Const Version = 2.35
'Const FechaVersion = "11/08/2006"
''Modificaciones: FGZ
''           Politica 450: Version 2 (Halliburton). Agregado de logs

'Const Version = 2.36
'Const FechaVersion = "22/08/2006"
''Modificaciones: FGZ
''           Politica 400: La rehice porque estaba mal.
''           Politica 410: Modificacion, ahora tiene encuenta que si hay una justificacion fija ya esta generada y no debe generar la AP.
''           Politica 480: Modificacion, No actualizaba las justificaciones cuando la justif era > que la anormalidad.
''           Sub Buscar_Justif: No se estaban inicializando los nro de justificaciones y cuando cambiaba de dia para un mismo legajo arrastraba esos nro y hacia macanas

'Const Version = 2.37
'Const FechaVersion = "25/08/2006"
''Modificaciones: FGZ
''           Politica 220: Modificaciones varias.

'Const Version = 2.38
'Const FechaVersion = "28/08/2006"
''Modificaciones: FGZ
''           Politica 220: Modificaciones, le saque la restriccion por el tipo de hora.
''           Politica 110: Le agregué un control al estandar para que chequee que el par de reg faltante no este cubierta por las registraciones hechas

'Const Version = 2.39
'Const FechaVersion = "28/08/2006"
''Modificaciones: FGZ
''           Politica 220: Modificaciones, Habia un bug en un sql y tomaba registraciones de otros legajos

'Const Version = 2.4
'Const FechaVersion = "29/08/2006"
''Modificaciones: FGZ
''           Politica 220: Modificaciones, Habia un bug en un sql y tomaba registraciones de otros legajos

'Const Version = 2.41
'Const FechaVersion = "31/08/2006"
''Modificaciones: FGZ
''           Las variables que llevaban el nro de desgloce de FP estaba definido integer y cuando el nro de desgloce supera los 32.6xx da error de desbordamiento.
''           Procedimientos:
''           Cambiar_Horas_Desde
''           Cambiar_Horas_Desde2
''           Cambiar_Horas_Hasta
''           Cambiar_Horas_Hasta2
''           Cambiar_Horas
''           Politica 1000: Modificaciones en procedimientos internos,
''                   Generar_Horario_Normal
''                   Generar_Horario_Normal_C1
''                   Generar_Horario_Normal_C3
''           Politica 1001: Modificaciones en procedimientos internos,
''                   Generar_Horario_Libre
''                   Generar_Horario_Libre_C1
''                   Generar_Horario_Libre_3

'Const Version = 2.42
'Const FechaVersion = "06/09/2006"
''Modificaciones: FGZ
''           En el main cuando lee el campo bprcparam hay ocaciones en que el asp no pasa los parametros, es decir, queda vacion ==> da un error asociando el campo depurar.
''           Detalles de log cuando pasa al historico
''           Tiempos en los log de pol 220

'Const Version = 2.43
'Const FechaVersion = "06/09/2006"
''Modificaciones: FGZ
''           Sub ConvertirHorasNormales2

'Const Version = 2.45
'Const FechaVersion = "07/09/2006"
''Modificaciones: FGZ
''           Sub Pol 220

'Const Version = 2.46
'Const FechaVersion = "20/09/2006"
''Modificaciones: FGZ
''           Sub Pol 700 - Se agregó las conversiones para el coeficiente 106

'Const Version = 2.47
'Const FechaVersion = "22/09/2006"
''Modificaciones: FGZ
''           Sub Pol 1000 - Version Halliburton, habia error cuando fraccionaba las horas y en el desglose venia nulo

'Const Version = 2.48
'Const FechaVersion = "29/09/2006"
''Modificaciones: FGZ
''           Sub Pol 410 - se podia producir un ciclo infinito

'Const Version = 2.49
'Const FechaVersion = "02/10/2006"
''Modificaciones: FGZ
''           Sub Pol 480     - Inserto siempre
''           Sub Pol 200V1   - Estaba armando mal la consulta cuando la salida del embudo pasaba de dia
''           sub generar_justificacion_Parcial - Si tenia solo justif en la salida no generaba ninguna

'Const Version = 2.5
'Const FechaVersion = "10/10/2006"
''Modificaciones: FGZ
''           Sub Pol 20     - Habia un error cuando insertaba en la traza


'Const Version = 2.51
'Const FechaVersion = "18/10/2006"
''Modificaciones: FGZ
''           Sub ConvertirHorasNormales2 (parte de la pol 220 - Custom Halliburton)
''               le agregué un chequeo para ver si las hs adicionales ya se generaron


'Const Version = 2.52
'Const FechaVersion = "08/11/2006"
''Modificaciones: FGZ
''           Politica 1000: Version estandar: Se corrigió el sub Generar_Horario_Normal cuando busca los desgloces por ambas
''           Politica 1001: Version estandar: Se corrigió el sub Generar_Horario_Libre cuando busca los desgloces por ambas

'Const Version = 2.53
'Const FechaVersion = "03/11/2006"
''Modificaciones: FGZ
''           Politica 14: Version 3: version para AGD, Turnos rotativos flexibles
''


'Const Version = 2.54
'Const FechaVersion = "21/11/2006"
''Modificaciones: FGZ
''           Politica 14: Version 3: version para AGD, Turnos rotativos Flexibles.
''                                   Problema de los dias sabados
''                                   Tema del procesamiento on-line que procesa cualquier cosa y luego
''                                       cuando se dispara el proceso para el legajo solo anda bien

'Const Version = 2.55
'Const FechaVersion = "27/11/2006"
''Modificaciones: FGZ
''           Politica 35(Sin Registraciones en dias feriados y no Francos): Nueva Politica.
''                           Versiones:
''                            1: Genera Tipo de Hora configurada en el turno (Generalmente Ausencia) para el empleado sin registraciones.
''                                Cantidad de Horas   ==> segun diferencia entre E/S teoricas del dia.
''                                Anormalidad         ==> (8) Ausencia
''                                Obs                 ==> El parametro de la politica que se configura es tipo de hora interna
''
''                            2: Genera Tipo de Hora configurada en el turno (Generalmente Ausencia) para el empleado sin registraciones.
''                                Cantidad de Horas   ==> segun horas obligatorias del dia.
''                                Anormalidad         ==> (8) Ausencia
''                                Obs                 ==> El parametro de la politica que se configura es tipo de hora interna




'Const Version = 2.56
'Const FechaVersion = "05/12/2006"
''Modificaciones: FGZ
''    Blanqueo de las Registraciones Marcadas en la Fecha
''    Pol 14 Version 2 : Cambié la consulta que busca la registraciones
''    Pol 85           : Modificacion cuando cuenta la cantidad de registraciones
''    Pol 110 Version 7: Nueva Version, muy similar a la v2 pero genera anormalidad 7 y thnro de AP
''    Pol 190 Version 3: Nueva Version, para politica de Llegadas tarde
''    Pol 200 Version 3: Nueva Version, para politica de Salidas temprano
''                      Problema: no registra las LLT y ST cuando entra/sale despues/antes de la mitad de la franja del turno
''                      Sol:      Analiza en toda la franja en lugar de solo la 1era mitad de la franja
''    Procedimiento Insertar_GTI_Proc_Emp

'Const Version = 2.57
'Const FechaVersion = "07/12/2006"
''Modificaciones: FGZ
''    Pol 220 Version 1: Modificacion en Conversion de Hs Normales 100 % para Halliburton
''                       Generar_Horario_Normal_C1

'Const Version = 2.58
'Const FechaVersion = "15/12/2006"
''Modificaciones: FGZ
''    Pol 51 Nueva Politica: se encarga de controlar la ventana que arma la politica 50
''    Pol 50 Todas las Versiones: al final de la politica se le agregó una llamada a la politica 51
''    Pol 85 : Correcion, habia problemas con los turnos nocturnos en cuanto a la modificacion hecha en la version 2.56

'Const Version = 2.59
'Const FechaVersion = "19/12/2006"
''Modificaciones: FGZ
''    Pol 220 Version 1: Modificacion en Conversion de Hs Normales 100 % para Halliburton

'Const Version = 2.6
'Const FechaVersion = "19/12/2006"
''Modificaciones: FGZ
''    Pol 440 Version 8: Nueva version
'                        '1) Si el dia es laborable ==> la cantidad de horas la saca de la hs obligatorias del dia
'                        '2) Si el dia es Franco    ==> la cantidad de horas la saca de la hs obligatorias del turno
'                        '3) Si el dia es feriado   ==> la cantidad de horas las resuleve dependiendo del tipo de dia original, es decir,
'                            'a) Si el dia es laborable ==> la cantidad de horas la saca de la hs obligatorias del dia
'                            'b) Si el dia es Franco    ==> la cantidad de horas la saca de la hs obligatorias del turno

'Const Version = 2.6
'Const FechaVersion = "21/12/2006"
''Modificaciones: FGZ
''    Pol 220 Version 1: Modificacion en Conversion de Hs Normales 100 % para Halliburton

'Const Version = 2.61
'Const FechaVersion = "21/12/2006"
''Modificaciones: FGZ
''    Pol 51 Version 1: le agregué la validacion de la hora cuando corta la ventana. Tenia impacto sobre la politica 70

'Const Version = 2.62
'Const FechaVersion = "05/01/2007"
''Modificaciones: FGZ
''    Definicion de va Global Nro_Dia_Original As Long: Estaba en modulo PRC30 y la pasé al modulo Politicas

'Const Version = 2.63
'Const FechaVersion = "12/01/2007"
''Modificaciones: FGZ
''    Pol 450: Cuando el dia era feriado y franco andaba mal para la version estandar
         
'Const Version = 2.64
'Const FechaVersion = "26/01/2007"
''Modificaciones: FGZ
'''    Pol 220 Version 1: Modificacion en Conversion de Hs Normales 100 % para Halliburton
         
'Const Version = 2.65
'Const FechaVersion = "08/03/2007"
''Modificaciones: FGZ
'''    Pol 220 Version 1: Modificacion en Conversion de Hs Normales 100 % para Halliburton
         
'Const Version = 2.66
'Const FechaVersion = "21/03/2007"
''Modificaciones: FGZ
'''    Pol 50 Nueva Version 6: Especial para Turnos Nocturnos. Desde las 0000 del dia hasta -0400 + tamVent horas
         
'Const Version = 2.67
'Const FechaVersion = "26/03/2007"
''Modificaciones: FGZ
'''    Agregados de logs con fines estadisticos y en Pol 85
         
         
'Const Version = 2.68
'Const FechaVersion = "03/04/2007"
''Modificaciones: FGZ
''    Pol 110 Version 8: Nueva Version, muy similar a la v4 pero genera anormalidad 3 y thnro de FRO

'Const Version = 2.69
'Const FechaVersion = "03/04/2007"
''Modificaciones: FGZ
''    Pol 85           : Modificacion cuando cuenta la cantidad de registraciones


'Const Version = 2.7
'Const FechaVersion = "03/04/2007"
''Modificaciones: FGZ
''           Politica 1000: Nueva Version 4: version integrada de estandar + Nocturna sub Generar_Horario_Normal_C4
''           Politica 1000: Nueva Version 5: version Absoluta (para Nocturnos) sub Generar_Horario_Normal_C5

'Const Version = "2.71"
'Const FechaVersion = "18/04/2007"
''Modificaciones: FGZ
''           Politica 220: Modificacion: le agregué el control por EOF

'Const Version = "2.72"
'Const FechaVersion = "20/04/2007"
''Modificaciones: FGZ
'''           Politica 1001: Modificacion Version 1 y 2: subs Generar_Horario_Libre y Generar_Horario_Libre_C1
''            Politica 14  : Agregado de logs e inicializacion de variables de tipo de turno: movil, flexible.

'Const Version = "2.73"
'Const FechaVersion = "23/04/2007"
''Modificaciones: FGZ
''            Politica 60  : cambio en version estandar cuando el horario es movil, estaba cambiando la fecha de procesamiento.

'Const Version = "2.74"
'Const FechaVersion = "25/04/2007"
''Modificaciones: FGZ
''            Politica 50  : Version 4. Faltaba asignar fecha desde y fecha hasta.

'Const Version = "2.75"
'Const FechaVersion = "26/04/2007"
''Modificaciones: FGZ
''            Sub BlanquearVariables: faltaba esta inicializacion
''                                   UsaFeriadoConControl = False
''                                   Pasa_de_Dia = False

'Const Version = "2.76"
'Const FechaVersion = "30/04/2007"
''Modificaciones: FGZ
''                   Pol 190 y 200. no debe generar este tipo de anormalidades cuando el dia es feriado o franco, a excepcion de
''                    los turnos Moviles, Fexibles o Rotativos flexibles (Politica 14 V1, V2 y V3 respectivamente)


'Const Version = 2.77
'Const FechaVersion = "09/05/2007"
''Modificaciones: FGZ
''    Pol 110 Versiones 7 y 8: Modificacion, se le agregó el control de si la registracion es de entrada y/o salida
'''    Pol 85           : Modificacion cuando cuenta la cantidad de registraciones

'Const Version = 2.78
'Const FechaVersion = "10/05/2007"
''Modificaciones: FGZ
''    Pol 40 : Modificacion, Marca la registracion por la cual se generan las horas.

'Const Version = 2.79
'Const FechaVersion = "17/05/2007"
''Modificaciones: FGZ
''    dll bsucar_dia : Modificacion, se sacó las redefiniciones de ciertas variables ...que estaban definidas como globales
''                                   Dim Nro_Dia As Long
''             Sub BlanquearVariables: faltaba esta inicializacion
''                                    p_turcomp = False
''                                    Nro_Justif = 0
''                                    justif_turno = False
''                                    Tiene_Justif = False
''                                    P_Asignacion = False
''                                    Nro_Dia_Original = 0

'Const Version = "2.80"
'Const FechaVersion = "18/05/2007"
''Modificaciones: FGZ
''      Pol 85 : Modificacion, Problema turnos partidos con 1 solo par de reg. que abarcan ambas franjas teoricas.

'Const Version = "2.81"
'Const FechaVersion = "28/05/2007"
''Modificaciones: FGZ
''      Pol 80 : Modificacion en v4, Si encuentra otra reg ==> borra anormalidades de reg impares posiblemente generadas.

'Const Version = "2.82"
'Const FechaVersion = "29/05/2007"
''Modificaciones: FGZ
''      Pol 85 : Modificacion, otros Problemas en turnos partidos.

'Const Version = "2.83"
'Const FechaVersion = "06/06/2007"
''Modificaciones: FGZ
''      Pol 85 : Modificacion, otros Problemas en turnos Nocturnos.

'Const Version = "2.84"
'Const FechaVersion = "06/06/2007"
''Modificaciones: FGZ
'''    Pol 440 Version 9: Nueva version
''                        '1) Si el dia es laborable ==> la cantidad de horas la saca de la hs obligatorias del dia
''                        '2) Si el dia es Franco    ==> no justifica
''                        '3) Si el dia es Feriado   ==> no justifica
'============================================================================================
'Const Version = "3.00"
'Const FechaVersion = "06/06/2007"
''Modificaciones: FGZ
''      Mejoras generales de performance.
''       Sin gti_traza

'Const Version = "3.01"
'Const FechaVersion = "08/06/2007"
''Modificaciones: FGZ
'''    Prc30:    Cuando tiene justificacion y la pol 470 esta activa en version 1. ==> genera Just y termina
''                 ahora en ese caso se agregó la Imputacion de horas sin control de prescencia Politica(450)

'Const Version = "3.02"
'Const FechaVersion = "11/06/2007"
''Modificaciones: FGZ
'''    Politica 14: Version 4: version para AGD, Turnos flexibles con subturnos de distintos tamaños

'Const Version = "3.03"
'Const FechaVersion = "27/06/2007"
''Modificaciones: FGZ
'''    Modulo de clase    :    BuscarTurno: faltaban inicializar algunas variables

'Const Version = "3.04"
'Const FechaVersion = "04/07/2007"
''Modificaciones: FGZ
''''    Politica 80: Version 5: Trilenium, toma las puntas teniendo en cuenta E/S si hay una sola reg
''''    Politica 80: Version 6: Trilenium, toma las puntas teniendo en cuenta E/S si hay una sola reg, y genera la registracion Teorica faltante
''''    Politica 14: Version 1: Trilenium, no seteaba la fecha hasta cuando el turno es nocturno
''''    Politica 20: , faltaba un campo en el select y daba error cada vez que queria depurar

'Const Version = "3.05"
'Const FechaVersion = "20/07/2007"
''Modificaciones: FGZ
'''    Politica 61 Nueva Politica: se encarga de controlar la ventana que arma la politica 60
''                                AGD, la cantidad de registraciones son impares y mayor que la cantidad de registraciones obligatorias
''                             ==> descarta la ultima registracion y se acorta la ventana
''     Politica 200 Version V2:  No genera ST cuando sale el dia anterior (en caso de ser turno Nocturno)
'''    Politica 50: Version 1: se corrigió un select por falta de un campo
''     Politica 190 Version V2:  No genera LLT cuando entra el dia posterior a la entrada (en caso de ser turno Nocturno)

'Const Version = "3.06"
'Const FechaVersion = "30/07/2007"
'Modificaciones: FGZ
'    Politica 1001:Modificaciones en procedimientos internos para Versiones estandar(1) y Nocturno(3)
'                   Generar_Horario_Libre: se le agregó un control para que no genere Ausencias cuando el dias es feriado y se lo considera franco (POL 25)y el turno es libre.
'                   Generar_Horario_Libre_3: se le agregó un control para que no genere Ausencias cuando el dias es feriado y se lo considera franco (POL 25)y el turno es libre.

'Const Version = "3.07"
'Const FechaVersion = "17/09/2007"
''Modificaciones: Diego Rosso
''    Se modifico la funcion Feriado para que busque en todas las estructuras asignadas en la pol. de alcance para GTI ya que buscaba solo en la primera. CAS-04896

'Const Version = "3.08"
'Const FechaVersion = "21/09/2007"
'Modificaciones: FGZ
'    Politica 61:se agregó control de si reloj marca e/s

'Const Version = "3.09"
'Const FechaVersion = "30/10/2007"
'Modificaciones: FGZ - G. Bauer
'    se modifico para que traigas bien la entrada y salida cuando utiliza ventanas.

'Const Version = "3.10"
'Const FechaVersion = "12/11/2007"
''Modificaciones:
''Diego Rosso
''PARA HORARIO FLEXIBLE: Cuando ModificaHT es falso graba el horario que le corresponderia en el dia
''                       y cuando es verdadero graba el numero de dia calculado.

'Const Version = "3.11"
'Const FechaVersion = "11/01/2008"
'Modificaciones: FGZ
'   Politica 14: version 5: Es igual que la version 4 pero resulve un prblema de horario extra anterior al HT
' Ademas se redefinió el sub Horario_Teorico y se lo puso en el modulo mdlTurno
'   Politica 14: version 6: Es igual que la version 5 pero resulve un prblema de horario extra anterior al HT
'                       y con un solo par de registracion se cubren 2 subturnos

'Const Version = "3.12"
'Const FechaVersion = "04/04/2008"
''Modificaciones: Diego Rosso - Se agrego el parametro version a la politica.
''                              Se creo la version 2--> Se queda con la registracion de mayor hora del intervalo y borra las restantes.
''                              Se creo la version 3--> igual a la 2 pero con la diferencia que va recalculando el intervalo(tolerancia) cada vez que elimina una registracion.
''                                 Version 1--> Se queda con la primera registracion que encuentra en el intervalo y borra las restantes.

'Const Version = "3.13"
'Const FechaVersion = "14/03/2008"
''Modificaciones: FGZ
''   Politica 15: Se le agregaó mas detalle de log
''   Politica 14: version 6: Es igual que la version 5 pero resulve un prblema de horario extra anterior al HT
''                       y con un solo par de registracion se cubren 2 subturnos


'Const Version = "3.14"
'Const FechaVersion = "08/04/2008"
''Modificaciones: FGZ
''   Politica 80: todas las versiones.
''           No estaba marcando como leidas las registraciones una vez que las "leia"

'Const Version = "3.15"
'Const FechaVersion = "16/04/2008"
''Modificaciones: FGZ
''   Modulo politicas: sub Cargar_DetallePoliticas.
''           Se cambió en el where el <> '' por IS NOT NULL

'Const Version = "3.16"
'Const FechaVersion = "17/04/2008"
''Modificaciones: FGZ
''   politica 15:
''           Mejora para que tome diferencias negativas u ajustes en el calculo de diferencia

'Const Version = "3.17"
'Const FechaVersion = "13/05/2008"
''Modificaciones: FGZ
''   DLL Buscar Turno:
''           sub Buscar_Turno: solo tiene en cuenta justif por licencias aprobadas u otro tipo de just
''   modulo Buscar Turno:
''           sub Buscar_Turno_nuevo: solo tiene en cuenta justif por licencias aprobadas u otro tipo de just

'Const Version = "3.18"
'Const FechaVersion = "27/05/2008"
''Modificaciones: FGZ
''   Se hicieron modificaciones para que solo se tomen registraciones que no tienen la marca de llamada.
''       Esas registraciones son procesadas solamente por la politica 510 "Horas de Llamada"
'
''    Politica 510 Nueva Politica: se encarga de pagar las horas de llamada
''           Version 1:  llama a Politica510_APQ
''           Version 2:  llama a Politica510_APS
''           Version 3:  llama a Politica510_BB
''           Version 4:  llama a Politica510_BB2
'
''   Ademas se hicieron versiones nuevas de la politica 80
''           version 7: Lama a Politica80vBB custom para CARGIL
''           version 8: Lama a Politica80vAPQ custom para CARGIL
''           version 9: Lama a Politica80vAPS custom para CARGIL
''           version 10: Lama a Politica80vBB2 custom para CARGIL
''   Otras Politicas modificadas
''           1000 (en todas sus versiones)
''           1001 (en todas sus versiones)
''           85 (para que recorra las registraciones de llamada)
''           190 (en todas las versiones)
''           200 (en todas las versiones)
''   Procedimientos generales modificados
''           sub Buscar_Turno:
''           sub Buscar_Turno_nuevo:

'Const Version = "3.19"
'Const FechaVersion = "03/06/2008"
''Modificaciones: FGZ
''    Politica 50 Nueva version
''           Version 7:  Igual version 1 pero si la 1er reg es un domingo despues de las 20 hs ==> no se considerara ninguna registracion

'Const Version = "3.20"
'Const FechaVersion = "07/07/2008"
''Modificaciones: FGZ - Circuito de firmas para partes de turno y partes de asignacion horaria
'''   Procedimientos generales modificados
'''           sub Buscar_Turno:
'''           sub Buscar_Dia:


'Const Version = "3.21"
'Const FechaVersion = "07/07/2008"
''Modificaciones: FGZ - Se agregó
'''    Politica 499 Nueva Politica: Feriados Nacionales

'Const Version = "3.22"
'Const FechaVersion = "30/07/2008"
''Modificaciones: FGZ - Se modificaron las poliagregó
''Modificaciones: FGZ
''           Politica 1000: en todas las versiones. Modificacion en la validacion de la fraccion minima cuando paga por cantidad de horas.
''           Politica 1001: en todas las versiones. Modificacion en la validacion de la fraccion minima cuando paga por cantidad de horas.
''           Politica 15: Posible cambio de turno. Si el turno es libre no se debe ejecutar esta politica.
'

'Const Version = "3.23"
'Const FechaVersion = "17/08/2008"
''Modificaciones: FGZ
''    Politica 400 Politica de Almuerzos
''           Version 1:  Estandar Vieja
''           Version 2:  Andreani
''           Version 3:  Estandar Nueva
''    Politica80vAPQ


'Const Version = "3.24"
'Const FechaVersion = "25/08/2008"
'Modificaciones: CAT
'   Politica499
'   PRC_30: llamada a politica 85 en dias no laborables
'   Politica 110: V4 Custom Expofrut. Estaba generando AP que no existina

'Const Version = "3.25"
'Const FechaVersion = "28/08/2008"
''Modificaciones: CAT
''   politica 80
''           version Politica80vBB custom para CARGIL
''           version Politica80vBB2 custom para CARGIL
''           version Politica80vAPQ custom para CARGIL
''           version Politica80vAPS custom para CARGIL

'Const Version = "3.26"
'Const FechaVersion = "16/09/2008"
''Modificaciones: FAF
''   Politica510_APQ - Se modifico el orden en que se toman las registraciones. Se agrego la fecha


'Const Version = "3.27"
'Const FechaVersion = "30/09/2008"
''Modificaciones: FGZ
''    Politica80v3: tenia unos problemas y abortaba
''
''    Politica80: Versiones de CARGIL
''       cuando estas buscando las registraciones, en el caso de que haya una llamada de entrada
''       y sea del dia siguiente de procesamiento, no busca mas porque asume que las va a encontrar
''       en el procesamiento del dia siguiente
''       Politica80vBB
''       Politica80vBB2
''       Politica80vAPQ
''       Politica80vAPS
''       sub BuscarRegistracionLLamada: tenia un problema cuando no se crean los TTempWFDia (Franco)
''           entonces el insert de TTempWFTurno daba error y quedaba el proceso en Incompleto.

 
'Const Version = "3.28"
'Const FechaVersion = "10/10/2008"
''Modificaciones: FGZ y CAT
''    Politica80: Versiones de CARGIL
''       corrección que hice sobre las políticas de llamada:
''              Cuando la ultima registración encontrada es una Entrada de Llamada, busco la siguiente registración y si es una salida de llamada la considero.
''        Procedimientos
''            Politica80vBB
''            Politica80vBB2
''            Politica80vAPQ
''            Politica80vAPS
''
''            Y un procedimiento nuevo: BuscarSalidaLlamada


'Const Version = "3.29"
'Const FechaVersion = "17/10/2008"
''Modificaciones: CAT
''    Politica1000: Versiones de CARGIL - Sub Generar_Horario_Normal_C7()
''       El problema se daba cuando una persona en un turno noctuno entraba en un dia Laborable y salia en un dia Feriado.
''            Ej: suponiento que el dia 13/10 es Feriado y una persona trabaja del 12/10 a las 22:00  al 13/10 a las 06:00.
''            El sistema pagaba la fraccion de 22:00 a 24:00 como dia Lorable y la fraccion de 00:00 a 06:00 la pagaba como Feriado.
''            Esto para Cargill es Incorrecto ya que todo se paga en funcion de la forma de pago de la entrada.


'Const Version = "3.30"
'Const FechaVersion = "23/10/2008"
''Modificaciones: FGZ
''    Politica1001:
''       Nueva Versiones para Libres - Sub Generar_Horario_Libre_W()
''           Es igual que la estandar pero con unas correcciones en la generacion de FP por cantidad de horas y Franja (AMBAS)
''       Modificacion en version para Nocturnos - Generar_Horario_Libre_3
''           Tenia problemas cuando pasaba de un dia franco a Laborabl


'Const Version = "3.30"
'Const FechaVersion = "23/10/2008"
''Modificaciones: FGZ
'''           Politica 400: Le cambié el codigo de la anormalidad que genera de 7 a 12 porque la 7 ya se estaba utilizando para Ausencia Parcial.

'Const Version = "3.31"
'Const FechaVersion = "27/10/2008"
''Modificaciones: FGZ
''    Politica 400: Solo se dispara cuando tiene registraciones
''    Politica1001:
''       Modificacion en todas las versiones - Generar_Horario_Libre_x
''           Se le agregó fecha hasta en el insert cuando la cantidad de hs son menores a las obligatorias.
''           Esto traia problemas en la politica 480


'Const Version = "3.32"
'Const FechaVersion = "03/11/2008"
''Modificaciones: FGZ
''    Politica1000:
''       Modificacion en todas las versiones - Generar_Horario_Normal_x
''           redefiní los 2 procedimientos porque no estaban teniendo en cuenta la durecion de la hora.
''    Politica1001:
''       Modificacion en todas las versiones - Generar_Horario_Libre_x
''           redefiní los 2 procedimientos porque no estaban teniendo en cuenta la durecion de la hora.
''   DLL FehasHoras
''           subs Duracion_Hora y Convertir_A_Hora


'Const Version = "3.33"
'Const FechaVersion = "07/11/2008"
''Modificaciones: FGZ
''    Politica70:
''       Nueva Version 5 - Genera Anormalidad y ademas termina el procesamiento
''    Politica470:
''       Nueva Version 3 - Genero la Justificacion y en caso de que el empleado registre, no se genera nada.
''           NO se generan las horas sin control de presencia (Politica 450) ni la politica de Feriados (Politica 499)
''    Politica 91 Nueva Politica: Chequea inconsistencias del tipo Licencias o Novedad de dia completo y tiene registraciones.
''    Politica 92 Nueva Politica: Chequea inconsistencias del tipo Licencias o Novedad Parcial y tiene registraciones solapadas.


'Const Version = "3.34"
'Const FechaVersion = "12/11/2008"
''Modificaciones: FGZ
''    Politica410:
''       Unica Version: se le agregaron controles por eof
''        Cambiar_Horas_Desde2 y Cambiar_Horas_Hasta2
''        Cambiar_Horas_Desde3 y Cambiar_Horas_Hasta3
''   se modificó el sub Convertir_A_Hora_cDuracion


'Const Version = "3.35"
'Const FechaVersion = "04/12/2008"
''Modificaciones: FGZ
''   se modificó el sub Redondeo_Horas_Tipo2.
''   se modificó la definicion de parametros formales de subs Convertir_A_Hora_cDuracion2 y Convertir_A_Hora_cDuracion.
''   Politica1000: Nueva Version de CARGIL - Sub Generar_Horario_Normal_C8().

'Const Version = "3.36"
'Const FechaVersion = "19/12/2008"
''Modificaciones: FGZ
''   Politica40: pasó a ser configurable y se agregaron 4 versiones

'Const Version = "3.37"
'Const FechaVersion = "19/12/2008"
''Modificaciones: FGZ
''   Politica1000: Se modificó la version v4


'Const Version = "3.38"
'Const FechaVersion = "21/01/2009"
''Modificaciones: FGZ
''   Encriptacion de string de conexion
''    Politica 400 Politica de Almuerzos
''           Version 2:  Andreani
''           Version 3:  Estandar Nueva


'Const Version = "3.39"
'Const FechaVersion = "30/01/2009"
''Modificaciones: FGZ
''    funcion Feriado: si el empleado valida alguna de las estructuras asociadas al feriado ==> TRUE sino FALSE
''    Politica 190 y 200: Aplicacion en Dias Feriados laborables (antes no se analizaban)


'Const Version = "3.40"
'Const FechaVersion = "16/02/2009"
''Modificaciones: FGZ
''   Encriptacion de string de conexion


'Const Version = "3.41"
'Const FechaVersion = "26/02/2009"
''Modificaciones: FGZ
''    Politica1001:
''       Modificacion en todas las versiones - Generar_Horario_Libre_x
''           Cambié una consulta sobre gti_turno que se hacia de mas dado que el dato ya se tenia,
''           Ademas esta consulta podia retornar vacio y con esto un error posterior porque no se controlaba EOF.


'Const Version = "3.42"
'Const FechaVersion = "04/03/2009"
''Modificaciones: FGZ
''    Politica1001:
''       Modificacion en todas las versiones - Generar_Horario_Libre_x
''           cuando revisa si se llegó a la canidad de hs obligatoria hacia referencia a un campo que no tenia y deba EOF
''               mas exactamente el campo de duracion de la hora


'Const Version = "3.43"
'Const FechaVersion = "06/03/2009"
''Modificaciones: FGZ
''    Politica1000:
''       Modificacion en sub Cambiar_Horas_Desde que afecta a todos las versiones de la politica 1000 y 1001.
''           el problema estaba cuando el desgloce de hs es fijo/variable y algun desgloce pasaba de dia


'Const Version = "3.44"
'Const FechaVersion = "07/04/2009"
''Modificaciones: FGZ
''    Sub esFeriado:
''       Se agregaron mensajes de log

'Const Version = "3.45"
'Const FechaVersion = "08/04/2009"
''Modificaciones: FGZ
''    Politica1001:
''       Nueva version estandar. Se corrigieron los problemas en las FP relativas para los turnos nocturnos.


'Const Version = "3.46"
'Const FechaVersion = "13/04/2009"
''Modificaciones: FGZ
''    Politica 50 Nueva version
''           Version 8:  Igual version 1 pero si la 1er reg es un domingo despues de las 18 hs ==> no se considerara ninguna registracion
''           Version 9:  Igual version 1 pero si la 1er reg es un domingo despues de las 17 hs ==> no se considerara ninguna registracion



'Const Version = "3.47"
'Const FechaVersion = "02/06/2009"
''Modificaciones: FGZ
''    Politica 600 Nueva Politica: Control de Hs Nocturnas.
''    Politica 402 Nueva Politica: Vales Comedor.
''    Politica 21 Nueva Politica: Parte las registraciones que pasan de un dia para otro.
''               Crea las registraciones en los limites del dia cuando entra en un dia y sale en otro
'

'Const Version = "3.48"
'Const FechaVersion = "24/06/2009"
''Modificaciones: FGZ
''    Politica 190 : Llegadas tarde.
''           Version 3: Version integrada.
''               Estaba teniendo problemas cuando habia mas registrasciones de las obligatorias


'Const Version = "3.49"
'Const FechaVersion = "07/08/2009"
''Modificaciones: FGZ
''    Se trata de un cambio muy grande y de mucho impacto
''       Descripcion
''           Se agregó un campo en la configuracion del subturno para determinar
''           para que dia se pagan las horas cuando las registraciones pasan de un dia para el otro
''
''               gti_subturno.subtgen
''
''           Ademas se agregaron 2 campos en la tabla de registraciones y horario cumplido
''               gti_registraciones.fechagen date
''               gti_horcumplido.horfecgen
''           y en sus respectivos historicos
''
''       Scripts
''        -- Oracle ----------------------------------------------------------------
''        -- se le agregó el campo subtgen a la tabla gti_subturno
''        ALTER TABLE gti_subturno ADD subtgen INT DEFAULT 1 null;
''        commit;
''
''        ALTER TABLE gti_registracion ADD fechagen date null;
''        commit;
''        ALTER TABLE gti_hisreg ADD fechagen date null;
''        commit;
''
''        ALTER TABLE gti_horcumplido ADD horfecgen date null;
''        commit;
''        ALTER TABLE gti_hishc ADD horfecgen date null;
''        commit;
''
''        -- SQL Server ------------------------------------------------------------
''        ALTER TABLE [dbo].[gti_subturno] ADD [subtgen] [int] NULL
''        GO
''        ALTER TABLE [dbo].[gti_registracion] ADD [fechagen] [datetime] NULL
''        GO
''        ALTER TABLE [dbo].[gti_hisreg] ADD [fechagen] [datetime] NULL
''        GO
''        ALTER TABLE [dbo].[gti_horcumplido] ADD [horfecgen] [datetime] NULL
''        GO
''        ALTER TABLE [dbo].[gti_hishc] ADD [horfecgen] [datetime] NULL
''        GO
''
''    Politicas afectadas:
''       Politica 15 : Posible cambio de turno.
''       Politica 30 y 35: Ausencias.
''       Politica 40 : Una sola registracion.
''       Politica 70 : Registraciones impares.
''       Politica 91 : Inconsistencias.
''       Politica 92 : Inconsistencias.
''       Politica 110 : Falta de registraciones obligatorias.
''       Politica 190 : Llegadas tarde.
''       Politica 200 : Salidas Temprano.
''       Politica 400 : Almuerzo variable.
''       Politica 402 : Vales Comedor.
''       Politica 450 : Horas sin control de presencia.
''       Politica 480 : Generacion de justificaciones parciales variables.
''       Politica 510 : Horas de llamada.
''       Politica 540 : Tope mensual de horas.
''       Politica 1000 : Generacion de Horario Normal.
''       Politica 1001 : Generacion de Horario Libre.
''
''   Tambien se agregó un procedimiento general en el modulo global para hacer chequeos de version
''   Este sub se debe invocar luego de crear el archivo de log y antes de comenzar con la logica especifica de cada proceso
''   Si el proceso no pasase el chequeo de version el proceso finalizará en estado "Error de Version"


'Const Version = "3.50"
'Const FechaVersion = "20/08/2009"
''Modificaciones: FGZ
''    Politica 80 Version 3: No ampliaba la ventana porque no encontraba la registraciones en L



'Const Version = "3.51"
'Const FechaVersion = "27/08/2009"
''Modificaciones: FGZ
''    Politica 1001 todas las versiones: Se agregó que escribieran que registraciones estan involucradas en cada hora generada
''                                       que no lo estaba grabando para las FP por cantidad y Ambas


'Const Version = "3.52"
'Const FechaVersion = "14/09/2009"
''Modificaciones: FGZ
''    Politica 1001 todas las versiones: Habia problemas en las FP por cantidad y Ambas

'Const Version = "3.53"
'Const FechaVersion = "23/09/2009"
''Modificaciones: FGZ
''    Politica 1001: sub Generar_Horario_Libre_Detallado: Habia problemas en las FP por cantidad y Ambas

'Const Version = "3.54"
'Const FechaVersion = "01/10/2009"
''Modificaciones: FGZ
''    Politica 1001: sub Generar_Horario_Libre_STD: Habia un problema en el insert con las horas desde y hasta

'Const Version = "3.55"
'Const FechaVersion = "05/10/2009"
''Modificaciones: FGZ
''    Politica 1001: sub Generar_Horario_Libre_Detallado: Cambio en la manera que redondea


'Const Version = "3.56"
'Const FechaVersion = "15/10/2009"
''Modificaciones: FGZ
''    Politica 499: Feriados Nacionales: Se agregaron 3 parametros (version, lista de tipos de horas y lista de tipos de licencias)


''--------------------------------------------------
'Const Version = "4.00"
'Const FechaVersion = "19/11/2009"
''Modificaciones: FGZ
''    Cambios Importantes
''        ALTER table gti_horcumplido add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_hishc add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_acumdiario add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_hisad add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_achdiario add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_his_achdiario add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_det add horas varchar(10) null default ‘0:00’
'
''   Se agregó 1 campo en varias tablas para agregar la funcionalidad de que el resultado
''       se pueda expresar en distintas unidades (Valor decimal o en horas y minutos)
''   -----
''   OBS:
''   -----
''       El proceso de generacion de novedades de GTI permanece sin cambios, es decir,
''           el alcance de las modificaciones afectan hasta el proceso de Acumulado Parcial.
''       La cuenta corriente de horas se sigue manejando en valores decimales solamente.


'Const Version = "4.01"
'Const FechaVersion = "04/12/2009"
''Modificaciones: FGZ
''    Politica 430 Nueva Politica: Horas de Descanzo.

'Const Version = "4.02"
'Const FechaVersion = "14/01/2010"
''Modificaciones: FGZ
''    Politica 430: Horas de Descanzo. Se corrigieron unos problemas


'Const Version = "4.03"
'Const FechaVersion = "20/01/2010"
''Modificaciones: FGZ
''    Politica 201 Nueva Politica: Impuntualidad.

'Const Version = "4.04"
'Const FechaVersion = "01/03/2010"
''Modificaciones: FGZ
''    Politica 201 - Impuntualidad: Se generó una nueva version de la politica.
''                           Version 1: genera siempre cantidad 1 cuando hay anormalidad.
''                           Version 2: genera la cantidad de minutos de impuntualidad cuando hay anormalidad.

'Const Version = "4.05"
'Const FechaVersion = "03/03/2010"
''Modificaciones: FGZ
''    Procedimientos generales modificados
''                           sub Buscar_Turno_nuevo del modulo mdlturno: esta funcin solo se usa para buscar turnos en dias que no es el que se está procesando

'Const Version = "4.06"
'Const FechaVersion = "04/03/2010"
''Modificaciones: FGZ
''    Politica 402 Vales Comedor. Se agregaron mensajes de log.


'Const Version = "4.07"
'Const FechaVersion = "30/03/2010"
''Modificaciones: FGZ
''    Politica 21 : Parte las registraciones que pasan de dia.
''           Todas las Version .
''               Revisa si los relojes distinguen entre E y S
''               sino lo hacen ==> se supone que la 1er reg es E y la siguiente es S


'Const Version = "4.08"
'Const FechaVersion = "07/04/2010"
''Modificaciones: FGZ
''    Politica 200 : Salida Temprano.
''           Todas las Versiones.
''                       Estaba aplicando el redondeo de la ST aun cuando debia generar AP, se arregló para que use el redondeo que corresponde


'Const Version = "4.09"
'Const FechaVersion = "16/04/2010"
''Modificaciones: FGZ
''    Politica 402 Vales Comedor.
''               Se le agregó la opcion de controlar que haya trabajado en rango corrido
''               Esta actualizacion lleva actualizacion de BD y asp
''                -- SQL
''                ALTER TABLE gti_vales ADD valida_hcorr smallint NOT NULL default (0)
''                GO
''                ALTER TABLE gti_vales ADD hc1_ent varchar(4)
''                GO
''                ALTER TABLE gti_vales ADD hc1_sal varchar(4)
''                GO
''                ALTER TABLE gti_vales ADD hc2_ent varchar(4)
''                GO
''                ALTER TABLE gti_vales ADD hc2_sal varchar(4)
''                GO
''                ALTER TABLE gti_vales ADD hc3_ent varchar(4)
''                GO
''                ALTER TABLE gti_vales ADD hc3_sal varchar(4)
''                GO
''
''                -- Oracle
''                ALTER TABLE gti_vales
''                ADD (
''                valida_hcorr INT default 0 NOT NULL,
''                hc1_ent VARCHAR2(4),
''                hc1_sal VARCHAR2(4),
''                hc2_ent VARCHAR2(4),
''                hc2_sal VARCHAR2(4),
''                hc3_ent VARCHAR3(4),
''                hc3_sal VARCHAR3(4));

'Const Version = "4.10"
'Const FechaVersion = "06/05/2010"
''Modificaciones: FGZ
''    Politica 1000 Horario Normal.
''           Habia problemas con las FP que son por cantidad de horas Y franja horaria (ambas) y ademas el turno es nocturno.
''       Se modificaron todas las versiones
''           Generar_Horario_Normal
''           Generar_Horario_Normal_C1
''           Generar_Horario_Normal_C3
''           Generar_Horario_Normal_C4
''           Generar_Horario_Normal_C5
''           Generar_Horario_Normal_C6
''           Generar_Horario_Normal_C7
''           Generar_Horario_Normal_C8


'Const Version = "4.11"
'Const FechaVersion = "10/05/2010"
''Modificaciones: EGO - Liz Oviedo
''           Justificaciones parciales. NO estaba justificando bien cuando las justificaciones eran por mas de 1 dia.
''               Cuando justificaba totalmente la anormalidad no borraba la anormalidad.
''       sub Ventana
''           La justificacion si era por varios dias no funcionaba bien.
''       sub Revisar_Justif
''           No ingresaba bien en el rango que correspondia


'Const Version = "4.12"
'Const FechaVersion = "11/05/2010"
''Modificaciones: FGZ
''    Politica 402 Vales Comedor.
''               Se corrigió bug en control de rango corrido.
''    Politica 21 Parte las registraciones que pasan de dia.
''               Se agregó control sobre las registraciones en estado X (que las tiene que descartar)


''----------------------------
'Const Version = "5.00"
'Const FechaVersion = "15/06/2010"
''Modificaciones: FGZ
''    Control por entradas fuera de termino.
''           Antes
''               cuando se queria procesar algo en una fecha que caia en un periodo cerrado no se procesaba.
''           Ahora
''               Se puede reprocear un periodo cerrado solo cuando se aprueba una entrada fuera de termino.
''               Para ese reprocesamiento solo se tendrá en cuenta todas las entradas fuera de termino aprobadas.


'Const Version = "5.01"
'Const FechaVersion = "08/07/2010"
''Modificaciones: FGZ
''    Politica 201 Puntualidad.
''               Se cambió el orden de ejecucion de las politicas 190, 200 y 201


'Const Version = "5.02"
'Const FechaVersion = "05/08/2010"
''Modificaciones: FGZ
''    Politica 111 Nueva Politica: Control de Secuencia de Registraciones. Nueva Politica.
''    Politica 14: Version 7: version para Monresa, En principio es igual que la version 1
''                 pero la cantidad de horas obligatorias las saca del dia de orden correspondiente
''

'Const Version = "5.03"
'Const FechaVersion = "20/08/2010"
''Modificaciones: FGZ
''    Politica 111: Control de Secuencia de Registraciones. Nueva version (2) Secuencia obligatoria.


'Const Version = "5.04"
'Const FechaVersion = "14/09/2010"
''Modificaciones: FGZ
'''   politica 80
''           version Politica80vBB custom para CARGIL
''           version Politica80vBB2 custom para CARGIL
''           version Politica80vAPQ custom para CARGIL
''           version Politica80vAPS custom para CARGIL
''         Estaban buscando mal las registraciones(traian registraciones que no correspondian al dia procesado).


'Const Version = "5.05"
'Const FechaVersion = "23/09/2010"
''Modificaciones: FGZ
''    Politica 14: Version 7: version para Monresa. Modificacion para cuando no encuentra registraciones en el dia.
''                 Version 8: nueva version para Monresa. Igual que la 7 pero solo busca registraciones para el dia procesado y con la marca propia de Monresa.
''    SUB PRC30. llamada a sub Ventana para cuando el turno es de Horario_Movil (politica 14 Activa)
''    politica 60:
''           version 1 y 2. Armaba mal las ventana cuando el turno es de Horario_Movil (politica 14 Activa).
''    Politica 1000 Horario Normal.
''           Habia combinaciones de turnos norturnos con FP relativas y ademas combiando con politica 14 que no se podian resolver.
''       Se modificaron todas las versiones
''           Generar_Horario_Normal_C6

'Const Version = "5.06"
'Const FechaVersion = "24/09/2010"
''Modificaciones: FGZ
''    politica 80:
''           version 11. Nueva version. Procesa las puntas pero mantiene el resto para control (politica 111)
''    Politica 111 Control de secuencia de registraciones.
''           Se hizo una modificacion para que funcione en conjunto con la v11 de la politica 80.
''    Politica 400 Politica de Almuerzos
''           Version 4:  Custom Monresa

'Const Version = "5.07"
'Const FechaVersion = "06/10/2010"
''Modificaciones: FGZ
''    Politica 14: Version 8: Nueva version para Monresa.
''    Politica 14: Version 9: Nueva version para Monresa.
''    Politica 14: Version 10: Nueva version para Monresa.
'
''    politica 20:
''           version 4. Nueva version. Se modificó el procedimiento porque estaba marcando mal la punta.
'
''    politica 80:
''           version 11. Nueva version. Se modificó el procedimiento porque estaba marcando mal la punta.
'
''    Politica 111: Control de secuencia de registraciones.
''           Nueva version 3: Un descanso obligatorio y n opcionales.
''           Nueva version 4: 2 descansos obligatorios.
''           Nueva version 5: 2 descansos obligatorios. Genera todas las anormalidades de las reg faltantes.
''          Los dias francos o feriados no hay control de pausas obligatorias. Solo se exige entrada y salida al turno, el resto es opcional.
'
''    Politica 400 Politica de Almuerzos
''           Se eliminó la Version 4:  Custom Monresa
''               Se reemplazó por  la politica 401
'
''    Nueva Politica 401 Politica de Pausas
''           Version 1:  Custom Monresa
'
'
''    Politica 598: Control de Hs Nocturnas 3.
''           Tiene 3 versiones parecedias a la politica 600.
'
''    Politica 597: Control de Hs Nocturnas 4.
''           Tiene 3 versiones parecedias a la politica 600.
'
''    Politica 599: Control de Hs Nocturnas 2.
''           Tiene 3 versiones parecedias a la politica 600.
'
''    Politica 600: Control de Hs Nocturnas.
''           Nueva version 2: no parte las horas solapadas sino que genera las horas aprobadas.
''           Nueva version 3: genera 1 tipo de hora si no se solapan y genera otro tipo de hora si se solapadan.
'
''    Politica1001:
''       Nueva version 6. Nocturnos flexibles, es decir, que usan la politica 14.


'Const Version = "5.08"
'Const FechaVersion = "08/10/2010"
''Modificaciones: FGZ
''    Politica 110: Falta de Registracion Obligatoria.
''       La politica pasó a ser configurable (ahora va a tener 5 parametros) Por una cuesation de compatibilidad
''           si la misma no se le configuran los parametros funciona igual que antes.
''    Politica 85: v2.
''       Se agregó esta version para que no redefina si es entrada o salida segun orden de registracion
''    politica 80:
''           version 1. Si tiene la marca del tipo de entrada la utiliza y sino deduce E o S segun orden.


'Const Version = "5.09"
'Const FechaVersion = "12/10/2010"
''Modificaciones: FGZ
''    Nueva Politica 131 Tolerancia Descuento para LLegadas Tarde
''    Nueva Politica 141 Tolerancia Descuento para Salidas Temprano
''    Politica 190: LLegadas Tarde.
''       Se modificaron todas las versiones de la politica para que tenga en cuenta la tolerancia seteada en la Politica 131.
''    Politica 200: Salidas Temprano.
''       Se modificaron todas las versiones de la politica para que tenga en cuenta la tolerancia seteada en la Politica 141.

'Const Version = "5.10"
'Const FechaVersion = "20/10/2010"
''Modificaciones: FGZ
''    Politica 495: Paros Gremiales - Descuentos. Nueva politica.
''    Politica 14: Version 10: Modificacion por cuando la fecha de salida es menor que la de entrada se deduce que pasa de dia pero daba error.
''    Politica 401 Politica de Pausas
''           Version 1:  Custom Monresa. No estaba grabando las horas desde y hasta cuando insertaba exceso de pausa

'Const Version = "5.11"
'Const FechaVersion = "22/10/2010"
''Modificaciones: FGZ
''    Politica 495: Paros Gremiales - Descuentos. Modificacion para tipo de Paro Monresa.
''    Nueva Politica 490 Descuentos Autorizados: Genera hs de descuento por novedades parciales fijas cargadas.
''    Politica 190: LLegadas Tarde.
''       habian quedado mal las fechas desde la ultima modificacion por la politica 131.
''    Politica 200: Salidas Temprano.
''       habian quedado mal las fechas desde la ultima modificacion por la politica 141.


'Const Version = "5.12"
'Const FechaVersion = "22/10/2010"
''Modificaciones: FGZ
''    Politica 490 Descuentos Autorizados: ajustes varios.


'Const Version = "5.13"
'Const FechaVersion = "28/10/2010"
''Modificaciones: FGZ
''    Politica 495 Paros Gremiales: ajustes cuando elimina las FRO.



'Const Version = "5.14"
'Const FechaVersion = "03/11/2010"
''Modificaciones: FGZ
''    Politica 60: Ventana Normal.
''       Nueva version 10: Especial para turnos rotativos. La idea es que cuando la rotacion es nocturna arme la ventana
''                       desde la ventana con entrada en el dia anterior. (Se hizo para CCU pero se puede aplicar en cualquier cliente)
''    Politica 111 Control de secuencia de registraciones:.
''    Politica 401 Politica de Pausas
''           Version 1:  Custom Monresa. .
'
'Const Version = "5.15"
'Const FechaVersion = "11/11/2010"
''Modificaciones: FGZ
''    Politica 14: Version 10: Modificacion cuando no hay registraciones ni procesamientos en los dias anteriores.


'Const Version = "5.16"
'Const FechaVersion = "21/12/2010"
''Modificaciones: FGZ
''   Sub PRC30
''           Se agregó al final un seccion para eliminar las registraciones automaticas generadas por la politica 21
''    Politica 21
''           se cambió la logica para dejar configurable si las registraciones generadas automaticamente seran permanentes o temporales.
''    politica 20:
''           version 4. Se modificó el procedimiento porque no estaba controlando bien todas las registracioens.


'Const Version = "5.17"
'Const FechaVersion = "04/01/2011"
''Modificaciones: FGZ
''    politica 1001:
''           version 6. Se modificó el procedimiento porque estaba insertando una justificacion incorrecta cuando faltan hs obligatorias.
''                       ademas estaba haciendo macanas con la fecha de procesamiento p_fecha
''    politica 1000:
''           version 6. 'estaba haciendo macanas con la fecha de procesamiento p_fecha.

'Const Version = "5.18"
'Const FechaVersion = "28/01/2011"
''Modificaciones: FGZ
''    politica 401:
''           Se modificó la hora a partir de la cual se considera nocturno (antes 2000 y ahora 1730).


'Const Version = "5.19"
'Const FechaVersion = "09/03/2011"
''Modificaciones: FGZ
''    politica 1001:
''           Cuando la cantidad de hs trabajadas es menor que la cantidad obligatoria, estaba generando el codigo de anormalidad 3 (Falta una registracion).
''               ahora genera el codigo de anormalidad 7 (ausencia parcial).


'Const Version = "5.20"
'Const FechaVersion = "22/03/2011"
''Modificaciones: FGZ
''           Politica 1000: Horario Normal
''               Todas las versiones
''               No estaba controlando que no se inserten registros con valor 0 hs luego del redondeo cuando se paga por cantidad de hs
''           Politica 1001: Horario Libre
''               Todas las versiones
''               No estaba controlando que no se inserten registros con valor 0 hs luego del redondeo cuando se paga por cantidad de hs
'
'
'
''    Nueva Politica 4: Notificaciones Horarias
''           Busca las notificaciones horarias (custom Sykes) y las pasa a un parte de asignacion horaria.
''           Debe eliminar los partes creados por una notificacion antes de procesar y luego
''               debe buscar notificion horaria y pasarlo a parte de asignacion horaria
''

'Const Version = "5.21"
'Const FechaVersion = "26/05/2011"
''Modificaciones: FGZ
''    Politica 400: Politica de Almuerzos
''           Version 4: Nueva version custom para Sykes

'Const Version = "5.22"
'Const FechaVersion = "06/06/2011"
''Modificaciones: FGZ
''    problema en sub de control de versiones

'Const Version = "5.23"
'Const FechaVersion = "21/06/2011"
''Modificaciones: FGZ
''   Politica 400 (Almuerzos Variables): Se modificó la version Custom 4 (Sykes - Costa Rica).
''       Se modifica el contenido da la tabla empleg de la tabla wc_lunch , en lugar de venir el legajo envian el nro de tarjeta
''       Ademas se agregó un campo a la tabla wc_lunch para registrar el tipo de tarjeta de gti
''    Se agregaó el control de firmas a las novedades horarias
''       Se modifico:
''           Buscar_Turno
''           Buscar_Turno_nuevo
''           Politica 480
''           Politica 490


'Const Version = "5.24"
'Const FechaVersion = "23/06/2011"
''Modificaciones: FGZ
''   Politica 91 (Inconsistencias Lic completa y Registraciones): No se estaban considerando las horas desde y hasta del armado de la ventana del turno

'Const Version = "5.25"
'Const FechaVersion = "30/06/2011"
''Modificaciones: FGZ
''    Politica 400: Politica de Almuerzos
''           Version 4: custom para Sykes. Cambios en las hs a generar

'Const Version = "5.26"
'Const FechaVersion = "25/07/2011"
''Modificaciones: FGZ
''    Politica 14: Turnos de Horario Variable
''           Version 11: Nueva version custom para Sykes. Toma la cantidad de hs del turno del parte de asignacion horaria
''    Politica 1000: Generacion de horario Normal
''           Version 9: Nueva version custom para Sykes.

'Const Version = "5.27"
'Const FechaVersion = "01/09/2011"
''Modificaciones: FGZ
''    Politica 1001: Generacion de horario Libre
''           Version 2: Generar_Horario_Libre_C1()
''           Version 3: Generar_Horario_Libre_3()
''           Version 5: Generar_Horario_Libre()
''       Problema en la generacion de hs de anormalidad por cantidad de hs cuando no llega al minimo

'Const Version = "5.28"
'Const FechaVersion = "06/09/2011"
''Modificaciones: FGZ
''    Politica 410: Ausencia Parcial
''       Problema en el manejo de borrado cuando la misma registracion es entrada y salida (problema que no deberia ocurrir pero...)

'Const Version = "5.29"
'Const FechaVersion = "12/09/2011"
''Modificaciones: FGZ
''    sub Buscar_turno_Nuevo
''       Habia un recordset mal referenciado
''   Politica 91: inconsistencias del tipo Licencias o Novedad de dia completo y tiene registraciones
''       Problema en turnos en los cuales el HT se resuelve por registraciones (Politica 14).
''       Se le agregó control sobre la definicion de la ventana de analisis.


'Const Version = "5.30"
'Const FechaVersion = "14/09/2011"
''Modificaciones: FGZ
''    Politica 130: Tolerancia general de llegada tarde.
''    Politica 190: Llegadas Taredes
''    Politica 140: Tolerancia general de salida temprano.
''    Politica 200: Salidas Temprano

'Const Version = "5.31"
'Const FechaVersion = "28/09/2011"
''Modificaciones: FGZ
''    Politica 30: Ausencias.
''           Se le agregó un parametro configurable para configurar el codigo de anormalidad que genera
''    Politica 14: Turnos de Horario Variable
''           Version 12: Nueva version custom para Sykes. Toma la cantidad de hs del turno del parte de asignacion horaria.
''                       si no hay registraciones entonces se queda con el horario teorico asignado por definicion o por el parte
''   Politica 85: Orden de Recorrida de Registraciones a Evaluar.
''               Problemas cuando se trata de un turno partido puedo tener problemas si la justificaciones estan en la primera franja

'Const Version = "5.32"
'Const FechaVersion = "04/10/2011"
''Modificaciones: FGZ
''    Politica 480: Justificaciones parciales variables.
''           Habia un problema cuando no estaba habilitado el control de firmas
''    Politica 14: Turnos de Horario Variable
''           Todas las Version: cuando se asigna un parte de cambio de turno tal que toca dia libre o el turno es libre daba un error de desborde


'Const Version = "5.33"
'Const FechaVersion = "25/10/2011"
''Modificaciones: FGZ
''    Politica 400: Politica de Almuerzos
''           antes solo se ejecutaba si habia registraciones pero surgió la necesidad de hacer estos controles por cada version.
''           Ahora para la version 4 revisa los almuerzos independientemente de si registra o no. El resto de las versiones quedan igual.
''               Ademas se le agregó un parametro opcional a la politica para poder configurar la lista de codigos de anormalidades
''               tal que si existen hs generadas con los codigos de anormalidades ==> NO se generan hs de lunch


'Const Version = "5.34"
'Const FechaVersion = "20/12/2011"
''Modificaciones: Margiotta, Emanuel
''    Se agrego en la política 80 version 2 (Proceamiento de las puntas), para que marque como leídas las registraciones dentro de las puntas.
'     Esto se hizo par Multivoice y se agrego al estandar porque multivoice tiene una custom 3.46c

'Const Version = "5.35"
'Const FechaVersion = "22/12/2011"
''Modificaciones: Margiotta, Emanuel
''   Se corrigio en la función Ventana_Movil cuando la justificación era para un turno nocturno que cambiaba de día y caía al final del turno
''   la sql estaba mal. Esto se toco para Monresa.

'Const Version = "5.36"
'Const FechaVersion = "26/12/2011"
''Modificaciones: Margiotta, Emanuel
''   Se mejoro la performance para la modificacion e la version 5.34

'Const Version = "5.37"
'Const FechaVersion = "19/01/2012"
''Modificaciones: FGZ
''    Politica 21 Politica: Registraciones Automaticas.
''               Crea las registraciones automaticas segun horario teorico. Se utiliza para turnos que no registran
''                   pero que se necesita que se generen todas las hs definidas en la forma de pago

'Const Version = "5.38"
'Const FechaVersion = "30/01/2012"
''Modificaciones: FGZ
''    Politica 21 Politica: Registraciones Automaticas.
''               Ajustes de validacion.

'Const Version = "5.39"
'Const FechaVersion = "09/02/2012"
''Modificaciones: EAM
''    Politica 4  : HJI realizo unas modificaciones para Syke para determinar si es un turno nocturno
''    Politica 19 : EAM- Se corrigio un bug que no estaba guardando bien los campos fehordesde,fechorhasta
''    Politica 21 : EAM- Se agrego una validación que si tiene una justificacion de día completo no inserte las registraciones (inserta segun el HT o el parte).


'Const Version = "5.39.2"
'Const FechaVersion = "09/02/2012"
'Modificaciones: Se publican las modificaciones segun el detatte siguiente:
'                   Fecha de Modificación: "29/12/2011"
''                  Modificaciones: Margiotta, Emanuel
''                                 Se hicieron mejoras varias para reducir los tiempos de procesamientos.


'Const Version = "5.40"
'Const FechaVersion = "04/09/2012"
''Modificaciones: CDM - CAS 16825
''               Se corrigio la sentecia que borraba la tabla temporal en la politica 410 ya que rompia en Oracle. Ahora se usa la funcion BorrarTempTable.
''               Politicas 400 se modifican las sentecias de insercion en la tabla gti_novedad ya que el campo gnovestado no acepta cadenas vacias ni nulas.


'Const Version = "5.41"
'Const FechaVersion = "19/09/2012"
''Modificaciones: FGZ
''Version no integrada de EAM 11/04/2012
''    Politica 60  : EAM- Se agrego la version 11 para monresa. Busca parte de asignacin horaria y sino se basa en el turno y verifica si el horario es nocturno y arma la ventana.
''                   Se Toco la función Cambiar_Horas_Desde y Cambiar_Horas_Hasta del Módulo de política para que verifique si tiene seteada la variable
''                   TurnoNocturno_HaciaAtras (Esta se setea de la politca 14 version 13) y busca el parte en el dia siguiente ya que los turnos nocturnos van hacia atras.
''Version no integrada de FGZ 18/05/2012
'''    Saqué la definicion la de la variable del modeulo del PRC30 y la pasé a global porque el procedimiento donde se usa está en un modulo compartido y está dando error de compilacion todos los demas procesos de gti.
'''               Global TurnoNocturno_HaciaAtras As Boolean  'EAM- Esto se setea de la version 13 de la politica 14 y se usa para insertar las registraciones en tabla temporales segun el HT
'
'
''                   Pol 200V2 hay un problema con las justificaciones de Ausencia Parcial
''                   Pol 200_Configurable. hay un problema con las justificaciones de Ausencia Parcial


'Const Version = "5.42"
'Const FechaVersion = "10/12/2012"
''Modificaciones: EAM (CAS 17794)
'' Se corrigieron errores en sintaxis en las sqls de la función GenerarJustificacionesAutomaticas


'Const Version = "5.43"
'Const FechaVersion = "02/01/2013"
''Modificaciones: EAM (CAS-17528)
''   Politica 4: Se modifico la política para que soporte distintas versiones y se agrego la versión 2 para America TV.
''               Si no se configura ninguna por version, toma por default la version 1 que es la de syke.

'Const Version = "5.44"
'Const FechaVersion = "01/02/2013"
''Modificaciones: EAM (CAS-16360)
''   Politica 720: Se agregó una nueva politica para el calculo de horas insalubres


'Const Version = "5.45"
'Const FechaVersion = "10/05/2013"
''Modificaciones: FGZ (CAS-19053)
''                   Pol 200_Configurable. hay un problema con las justificaciones de Ausencia Parcial cuando el teorico pasa de dia.
''                   Politicas 400. Habia un problema de redondeo en el manejo de la justificacion que hacia que se sonsumieran justificaciones que quedaban con resto.


'Const Version = "5.46"
'Const FechaVersion = "28/05/2013"
'Modificaciones: FGZ (CAS-19053)
'                   Pol 200. Se le agregó control de NULO a las tolerancias porque cuando NO se activa la Pol 140 (tolerancias) Dá error de tipos la POlitica 200 configurable.
'                   Politicas 400. Politica400V3. SUB GenerarJustificacionesAutomaticas(). Cuando hay mas de una jsustificacion acumulable estaba actualizando los registros y las hs desde hasta quedaban desprolijas.
'                               Ahora se insertan por separado con sus correspondientes paras de hs desde hasta
'                               Ademas se le hicieron controles prar que no justifique sobre Salidas temprano (la politica no estaba preparada para ese tipo de justificacion y estaba generando inconsistencias)

'Const Version = "5.47"
'Const FechaVersion = "05/11/2013"
''Modificaciones: CM (CAS-11908)
''           Se comenta el las querys de las politicas 50 y 51 el control de las registraciones de llamada para que traiga
''           Todas las registraciones independientemente de si tienen la marca o no
''
''   Ademas (FGZ)
''           Se agregó una version 5 a la pol 400 para justificar sobre Salidas temprano pero no está debidamente testeada por lo cual queda inactiva de momento
''           Pol 450 (Hs sin control de presencia). Se Agregó la version 3 que genera hs segun desglose por cantidad y/o franja horaria
''           Pol 1000 (Generacion de Horario Normal). V1 (Estandar) y V9 (Sykes) Se agregó control para que no tenga en cuenta los desgloses de tipo 2 y 3 (fijos y fijos sin registracion)



'   ** version no liberada aun ***************************
'   ** version no liberada aun ***************************
'   ** version no liberada aun ***************************
'Const Version = "5.48"
'Const FechaVersion = "26/12/2013"
''Modificaciones: FGZ
''           Pol 450 (Hs sin control de presencia). Version 3. Se corrigió problema de efecto colateral por variables globales.
''
''       Ademas
''           Correccion de problemas de borrado cuando se procesan mas de 900 empleados en simultaneo
''
''           Pol 1001 Se agregó la versión 7. Cambia la forma de pago para el dia de la regintración de entrada.
''           Cuando se configura para insertar las registraciones a las 00:00 entonces paga asi:
''               Ej: Si entra el domingo a las 22:00 y sale el lunes a las 06:00 la forma de pago queda 2 horas como dia domingo y 4 como dia lunes, además si es feriado pagará como feriado
''
''   Ademas
''           Politica190_Configurable. Se corrigió problemas cuando estan las 3 toleancias configuradas y solo activa la primera
''                                     y Ademas no estaba levantando el tipo de redondeo de la configuracion de horas del turno.
''           Politica200_Configurable. Se corrigió problemas cuando estan las 3 toleancias configuradas y solo activa la primera
''                                     y Ademas no estaba levantando el tipo de redondeo de la configuracion de horas del turno.




'Const Version = "5.49"
'Const FechaVersion = "21/02/2014"
''Modificaciones: EAM    CAS-23930 - SYKES  -  Nueva version para politica 400
''           Política 400: se agrego una nueva version 6 para Sykes SV, la cual descuenta las horas lunch a un tipo de hora configurada en la politica y que tiene mayor peso de horas.
''

'Const Version = "5.50"
'Const FechaVersion = "28/03/2014"
'Modificaciones: EAM    5CA  -
'           Se modifico el delete del PRC30 que borra las registraciones por la política 21. se cambio regfecha por fechagen en el where.
'       FGZ- Política 21: SYKES  -se agrego 2 parametros mas por si se quiere generar tipo de hora y anormalidad cuando se crean registraciones teoricas

'Const Version = "5.51"
'Const FechaVersion = "21/04/2014"
''Modificaciones: 21/04/2014 - fernandez, Matias - CAS - 21289 - AMR - Bug en Posible Cambio de Turno- cuando la politica 5 esta desactivada, la politica 15 toma la primera registracion como entrada.


'Const Version = "5.52"
'Const FechaVersion = "22/04/2014"
''Modificaciones: 22/04/2014 - FGZ - CAS-23481 - TELEFAX - Monresa - Custom GTI.
''    Politica 401 Politica de Pausas
''           Version 3:  Custom Monresa. Se agregó version para turno de Embotelladores con control de pausas individual por par de pausas


'Const Version = "5.53"
'Const FechaVersion = "04/06/2014"
''Modificaciones: CAS-25591 - SYKES EL SALVADOR - DESCUENTO DE HORAS LUNCH
''    Politica 400 Politica de Lunch
''           Version 6:  Syke SV. Se modifico la política para que genere hs lunch según la cantidad que viene informada en los movimientos horarios
''
''Ademas
''
''   Politica 21(Chequea si hay Licencias o Novedad parcial fija y ajusta registraciones automaticas en consecuencia).
''       solo las registraciones automaticas


'Const Version = "5.54"
'Const FechaVersion = "09/06/2014"
''Modificaciones: CAS-25133 - FARMOGRAFICA - error en corrimiento de turnos
''   Se modificaron los procedimientos de Buscar_Dia() tanto en modulo de clase como modulo mdlturno sub Buscar_Dia_Nuevo().
''   Cuando se configuraban turnos rotativos con subturnos con mas de una iteracion estaba calculando mal el dia y obtenia un horario teorico incorrecto.
''Ademas
''
''   Politica 21(Cuando genera registraciones teoricas chequea si existen registraciones, si existiesen ==> no crea registraciones).
'

'Const Version = "5.55"
'Const FechaVersion = "17/06/2014"
'Modificaciones: FGZ
'       CAS-26045 - Sykes- GTI - Error en lectura de movimientos horarios
'           Politica 400 Politica de Lunch
'               Version 6:  Syke SV. Se modifico la política para que genere hs lunch según la cantidad que viene informada en los movimientos horarios


'Const Version = "5.56"
'Const FechaVersion = "27/06/2014"
'Modificacion: Fernandez, Matias
'CAS-26049 - 5CA - Error en procesamiento masivo - Se creo la funcion borrar_todas, en dicha funcion se explica el motivo

'Const Version = "5.57"
'Const FechaVersion = "30/06/2014"
'Modificacion: Fernandez, Matias
'CAS-26049 - 5CA - Error en procesamiento masivo -  se corrigio el condicional mal puesto en la funcion borrar_todas


'Const Version = "5.58"
'Const FechaVersion = "18/07/2014"
''Modificacion: Fernandez, Matias
''CAS-26396 - TRILENIUM - ERROR EN DIA FERIADO -  se volvio a la funcionalidad de la version 5.55, la funcion borrar_todas queda sin uso
' '- se soluciona el problema planteado en el CAS-26049 - 5CA - Error en procesamiento masivo agregando un condicinal mas en la politica 21 (EESS - EES) .


'Const Version = "5.59"
'Const FechaVersion = "30/07/2014"
''Modificaciones: Fernandez, Matias
''       CAS-26396 - TRILENIUM - ERROR EN DIA FERIADO- se arreglo cuando escribia en el log en la politica 190 configurable


'Const Version = "5.60"
'Const FechaVersion = "12/08/2014"
''Modificaciones: FGZ
''       CAS-21778 - Sykes El Salvador - Bug PRC030 Duplicidad de novedades
''           Politica 21: Registraciones Automaticas
''               Control de justificaciones parciales fijas
''           Justificaciones Parciales Fijas. sub generar_justificacion_Parcial()
''               Cuando las justificacion parciales fijas pasan de día estaba insertando mal la fecha.
''           Politica 85: Orden de Recorrida de Registraciones a Evaluar
''               se agregaron 3 niveles mas de registraciones obligatorias
''           Politica 402: Vales Comedor
''               se agregaron 3 niveles mas de registraciones obligatorias
''           Politica 499: Feriados Nacionales
''


'Const Version = "5.61"
'Const FechaVersion = "15/09/2014"
''Modificaciones: FGZ
''       CAS-21778 - Sykes El Salvador - Bug PRC030 Duplicidad de novedades
''           Justificaciones Parciales Fijas. sub generar_justificacion_Parcial()
''               Cuando las justificacion parciales fijas pasan de día estaba tomando justificaciones de otro día.
''           Buscar_Justif: Cuando se trataba de turnos nocturnos y tenia justificaciones en ambos dias, no estaba tomando las justificaciones del dia de salida.


'Const Version = "5.62"
'Const FechaVersion = "07/10/2014"
''Modificaciones: FGZ
''       CAS-21778 - Sykes El Salvador- QA - Bug GTI Horas Nocturnas
''           Pol 1000 (Generacion de Horario Normal). V1 (Estandar) y V9 (Sykes) Se agregó control para desgloses fijos cuando el par de registraciones está entero en el dia de trasnoche)
''Ademas
''       CAS-21778 - Sykes El Salvador- QA - Bug GTI  - PRC30 _ Novedades Horarias
''           Justificaciones Parciales Fijas.
''           Buscar_Justif: no estaba estableciendo bien la primer y ultima jsutificacion del dia.
'


'Const Version = "5.63"
'Const FechaVersion = "17/10/2014"
''Modificaciones: FGZ
''       CAS-27466 - ASM - Mejora Politica 21
''           Politica 21: Registraciones Automaticas
''               Se agregó mejora para que funcione para registraciones que provienen de relojes que no distinguen E/S. La politica 5 DEBE estar incativa.



'Const Version = "5.64"
'Const FechaVersion = "04/11/2014"
'''Modificaciones: Fernandez, Matias
'''       CAS-27680- Coop Seguros - Bug en generacion de ausencias parciales
'''       Politica200_Configurable, salida temprana a pesar de que hay una registracion de entrada luego de la salida teorica


'Const Version = "5.65"
'Const FechaVersion = "01/12/2014"
''Modificaciones: FGZ & LM
''       CAS-28148 - UP - GTI error en turnos 24 hs
''           Se mejoró el procedimientos de calculo de horario teorico para que informe bien las fechas de entrada y salida para turnos de mas de 24 hs
''Ademas
''       CAS-27466 - ASM - Mejora Politica 21
''           Politica 21: Registraciones Automaticas
''               Se agregó mejora para que controle que las registraciones permanentes que se pudieron haber generado para dia posterior todavía existan y esten sin marcar. Para que no replique.



'Const Version = "5.66"
'Const FechaVersion = "06/04/2015"
'Modificaciones: Fernandez, Matias - CAS-30209 - MONASTERIO BASE 2- Bug en procesar novedades -
                 'Error en buscar el ultimo dia habil trabajado, posible ciclo infinito.
                   

'Const Version = "5.67"
'Const FechaVersion = "28/04/2015"
'Modificaciones: Fernandez, Matias -  CAS-30307-AMR - Bug en generacion de hs extras - No se cambia el dia en el desglose,
                                      'version 10 amr
                                      
                                      
                                      
'Const Version = "5.68"
'Const FechaVersion = "08/05/2015"
'Modificaciones: Fernandez, Matias -  CAS-30307-AMR - Bug en generacion de hs extras - No se cambia el dia en el desglose,
                                      'puse la llamada a las politicas 200 y 201 (habian desaparecido, junto a la politica 400v7 y 400v8)
                                      

'Const Version = "5.69"
'Const FechaVersion = "12/05/2015"
''Modificaciones: Fernandez, Matias -  CAS-30359 - 5CA - Bug en anormalidad de licencia
'                                      'Cuando la licencia es por dia completo, y el turno abarca mas de un dia,
'                                      'las horas que genera la licencia, se reparten

'Const Version = "5.70"
'Const FechaVersion = "12/06/2015"
'Modificaciones: FGZ -  CAS-28248 - ASM - Error en turnos nocturnos
'           Politica 1008: Horario Libre
'               Se agregó nueva version 8 para turnos nocturnos
'Ademas
'       CAS-27961 - TATA - Error en tablero de GTI
'           Politica 205: Forma de Pago en Feriados
'               Se transformó la politica a configurable. Se le agregó un parametró de opcion tal que
'                   En feriados Laborable se pueda elegir que FP generar.
'                       1- Solo FP Feriado
'                       2- Solo FP Laborable
'                       3- FP Feriado y Laborable
'           Politica 25: Indica si el Feriado es Franco
'               Se transformó la politica a configurable. Se le agregó un parametró de opcion tal que se pueda elegir si se sontrola Anormalidades o no
'                       1- SI Controla anormalidades
'                       2- NO Controla anormalidades

'Admas
'       CAS-27961 - TATA - Error en tablero de GTI
'           Se mejoró el procedimiento de buscar_justif para que no levante justificaciones parciales variables
'           Politica 400 - Se agregó detalle de log
                                 

'Const Version = "5.71"
'Const FechaVersion = "18/06/2015"
''Modificaciones: Fernandez, Matias -CAS-31461 - G.COMPARTIDA - Error en PRC30 - se declaran variables de tipo double en modulo fechashoras
''                                      (licencias por 45 dias producian overflow)

'Const Version = "5.72"
'Const FechaVersion = "23/06/2015"
'Modificaciones: Margiotta, Emanuel - CAS-28352 - Salto Grande - Custom GTI - Mejora en tableros de GTI -
'               Se saco el if para que inserte siempre el horario teorico y se dejo solo la propiedad manual si tiene movilidad


'Const Version = "5.73"
'Const FechaVersion = "07/07/2015"
'Modificacion: MDF - CAS-31852 - 5CA - Error en novedades horarias - Se contempla que haya mas de una 2 justificaciones parciales
'en un dia para insertarla en gti_horariocumplido

'Const Version = "5.74"
'Const FechaVersion = "06/08/2015"
'Modificacion: Carmen Quintero - CAS-31268 - IBT - CUSTOM PARTE DE CAMBIO DE TURNO - Se contempla si existen parte de cambio de turno con distribucion
'en un dia para insertarla en gti_horariocumplido

'Const Version = "5.75"
'Const FechaVersion = "24/09/2015"
'Modificacion : Fernandez, Matias- CAS-33202 - ILE - Error en generacion de hs extras - version 12 de politica 60, contempla diferentes aperturas de ventanas
                                                                                      'dependiendo del subturno

         
         
'Const Version = "5.76"
'Const FechaVersion = "30/09/2015"
'Modificacion : Fernandez, Matias- CAS-33202 - ILE - Error en generacion de hs extras - version 12 de politica 60, correccion e la llamada
                                                                                     ' a resta y suma horas cuando hay mas de una ventana
         
         
'Const Version = "5.77"
'Const FechaVersion = "14/10/2015"
'Modificacion : Fernandez, Matias-CAS-33296- TABACAL - error novedades/licencias para turnos nocturnos- correccion en dia que se generan las  justificaciones


Const Version = "5.78"
Const FechaVersion = "12/01/2016"
'Modificacion : Fernandez, Matias-CAS-35000 - G.Compartida - Bug en horas de ausencia- Correccion en generacion de ausencia en politica 30 v3


'VERSION NO LIBERADA -------------------------


'----------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------


Global IdUser As String
Global CEmpleadosAProc As Integer
Global CDiasAProc As Integer
Global IncPorc As Single
Global IncPorcEmpleado As Single
Global HuboErrores As Boolean
Global EmpleadoSinError As Boolean
Global Progreso As Single
Global ProgresoEmpleado As Single
Global fec_proc As Integer ' 1 - Política Primer Reg.
                           ' 2 - Política Reg. del Turno
                           ' 3 - Política Ultima Reg.
Global Usa_Conv As Boolean
Global objBTurno As New BuscarTurno
Global objBDia As New BuscarDia
Global objFeriado As New Feriado
Global diatipo As Byte
Global ok As Boolean
Global esFeriado As Boolean
Global hora_desde As String
Global fecha_desde As Date
Global fecha_hasta As Date
Global Hora_desde_aux As String
Global hora_hasta As String
Global Hora_Hasta_aux As String
Global No_Trabaja_just As Boolean
Global nro_jus_ent As Long
Global nro_jus_sal As Long
Global Total_horas As Single
Global Tdias As Integer
Global Thoras As Integer
Global Tmin As Integer
Global Cod_justificacion1 As Long
Global Cod_justificacion2 As Long

Global Horas_Oblig As Single
'Global Existe_Reg As Boolean
'Global Existe_Reg_LLamada As Boolean
Global Forma_embudo  As Boolean

Global tiene_turno As Boolean
Global Nro_Turno As Long
Global Tipo_Turno As Integer

Global Tiene_Justif As Boolean
Global Justif_Completa As Boolean
Global Nro_Justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global Nro_Grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global Trabaja     As Boolean ' Indica si trabaja para ese dia
Global Orden_Dia As Integer
Global Nro_Dia As Integer
Global Nro_Subturno As Integer
Global Dia_Libre As Boolean
Global Dias_trabajados As Integer
Global Dias_laborables As Integer

Global Aux_Tipohora As Integer
Global aux_TipoDia As Integer

Global Hora_Tol As String
Global Fecha_Tol As Date
Global hora_toldto As String
Global fecha_toldto As Date

'Global fe1 As Date
'Global fe2 As Date
'Global fe3 As Date
'Global fs1 As Date
'Global fs2 As Date
'Global fs3 As Date

Global fv1 As Date
Global fv2 As Date
Global fv3 As Date
Global fv4 As Date
Global fv5 As Date
Global fv6 As Date
Global fv7 As Date

Global v1 As String
Global v2 As String
Global v3 As String
Global v4 As String
Global v5 As String
Global v6 As String
Global v7 As String

Global tol As String

Global Cant_emb As Integer
Global toltemp As String
Global toldto As String
Global acumula As Boolean
Global acumula_dto As Boolean
Global acumula_temp As Boolean
Global convenio As Long

Global tdias_oblig As Single

Global procesada_en_partes As Boolean 'MDF-- no se usa
Global lista_procesadas As String 'MDF
Global lista_justificacionesparciales As String 'MDF

'FGZ - 18/05/2012 ------
'Global TurnoNocturno_HaciaAtras As Boolean  'EAM- Esto se setea de la version 13 de la politica 14 y se usa para insertar las registraciones en tabla temporales segun el HT



'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Horario cumplido.
' Autor      : Maller José D.
' Fecha      : 18/10/02
' Ultima Mod.: FGZ
' Descripcion: 19/08/2005
' ---------------------------------------------------------------------------------------------
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Fecha As Date
Dim Ternro As Long
Dim Legajo As Long

Dim pos1 As Byte
Dim pos2 As Byte
Dim strcmdLine As String
Dim tinicio
Dim i As Long

'Dim objConn As New ADODB.Connection
Dim Nombre_Arch As String
Dim PeriodoCerrado As Boolean

Dim myrs As New ADODB.Recordset
Dim rsEmpleado As New ADODB.Recordset
Dim rs_batch_proceso As New ADODB.Recordset
Dim rs_His_batch_proceso As New ADODB.Recordset
Dim rs_Per As New ADODB.Recordset

Dim ListaPar

Dim PID As String
Dim ArrParametros
'EAM- 29/12/2011
ReDim arrEmpProcesadosOK(0)
Dim listEmpleadoOk As String
Dim nroMaxEmpBorrar As Long
Dim topeListaBorrar As Long
Dim cantProgreso As Double
Dim totalProgreso As Double
Dim horaUpdateBD As String
lista_procesadas = "0" '-MDF-----------------------

    strcmdLine = Command()
    ArrParametros = Split(strcmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProceso = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProceso = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strcmdLine) Then
                NroProceso = strcmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    depurar = False
    HuboErrores = False
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    Call CargarNombresTablasTemporales

    'FGZ - 26/03/2007
    Cantidad_de_OpenRecordset = 0
    Cantidad_Call_Politicas = 0

'    'FGZ - 01/06/2007 ----
'    Cantidad_Feriados = 0
'    Cantidad_Turnos = 0
'    Cantidad_Dias = 0
'    Cantidad_Empl_Dias_Proc = 0
'    'FGZ - 01/06/2007 ----


    ' seteo del nombre del archivo de log
    Nombre_Arch = PathFLog & "PRC30" & "-" & NroProceso & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'EAM- Levanta los datos configurado en el SrvDefaults sino le pone 5 minutos (v6.00)
    Call SetarDefaultsReducido
    horaUpdateBD = Time

    tinicio = Now
   
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error Resume Next
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
   
    On Error GoTo ME_Main:

    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    'FGZ - 05/08/2009 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 1, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Final
    End If
    'FGZ - 05/08/2009 --------- Control de versiones ------
    
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    StrSql = "SELECT iduser,bprcfecdesde,bprcfechasta,bprcparam FROM batch_proceso WHERE bpronro = " & NroProceso
    'objRs.Open StrSql, objConn
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Parametros: "
        Flog.writeline Espacios(Tabulador * 2) & "Usuario: " & objRs!IdUser
        IdUser = objRs!IdUser
        Flog.writeline Espacios(Tabulador * 2) & "Desde: " & objRs!bprcfecdesde
        FechaDesde = objRs!bprcfecdesde
        Flog.writeline Espacios(Tabulador * 2) & "Hasta: " & objRs!bprcfechasta
        FechaHasta = objRs!bprcfechasta
        'FGZ - 01/09/2006
        Flog.writeline Espacios(Tabulador * 2) & "bprcparam: " & objRs!bprcparam
        
        'EAM- Escribe en el log el tiempo de espera (v6.00)
        Flog.writeline Espacios(Tabulador * 2) & "Tiempo de espera: " & (TiempoDeEsperaNoResponde - 1)

        If Not EsNulo(objRs!bprcparam) Then
            If InStr(1, objRs!bprcparam, ".") <> 0 Then
                ListaPar = Split(objRs!bprcparam, ".", -1)
                depurar = IIf(IsNumeric(ListaPar(0)), CBool(ListaPar(0)), False)
                'FGZ - 18/05/2010 - se le agregó un nuevo parametro ----
                If UBound(ListaPar) > 1 Then
                    If Not EsNulo(ListaPar(2)) Then
                        ReprocesarFT = IIf(IsNumeric(ListaPar(2)), CBool(ListaPar(2)), False)
                    Else
                        ReprocesarFT = False
                    End If
                Else
                    ReprocesarFT = False
                End If
                'FGZ - 18/05/2010 - se le agregó un nuevo parametro ----
            
                'EAM- Si tiene congigurado el progreso del proceso lo levanta sino es 1 por default (v6.00)
                If UBound(ListaPar) = 3 Then
                    If ListaPar(2) <> "" Then
                        cantProgreso = ListaPar(2)
                    Else
                        cantProgreso = 1
                    End If
                Else
                    cantProgreso = 1
                End If

                
            Else
                depurar = False
                ReprocesarFT = False
            End If
        Else
            depurar = False
            ReprocesarFT = False
            cantProgreso = 1
        End If
        Flog.writeline Espacios(Tabulador * 2) & "Log detallado: " & depurar
        Flog.writeline Espacios(Tabulador * 2) & "Reprocesar Periodo Cerrado: " & ReprocesarFT
    Else
        Exit Sub
    End If
    OpenConnection strconexion, CnTraza
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    'EAM- Escribe en el log el progreso configurado (v6.00)
    Flog.writeline Espacios(Tabulador * 2) & "Progreso: " & cantProgreso
    
    
    StrSql = "SELECT empleado.ternro, empleado.empleg FROM batch_empleado "
    StrSql = StrSql & " INNER JOIN empleado ON batch_empleado.ternro = empleado.Ternro "
    StrSql = StrSql & " WHERE batch_empleado.bpronro = " & NroProceso
    If myrs.State = adStateOpen Then myrs.Close
    OpenRecordset StrSql, myrs
    
    CEmpleadosAProc = myrs.RecordCount
    If CEmpleadosAProc = 0 Then CEmpleadosAProc = 1
    CDiasAProc = DateDiff("d", FechaDesde, FechaHasta) + 1
    IncPorc = ((100 / CEmpleadosAProc) * (100 / CDiasAProc)) / 100
    IncPorcEmpleado = (100 / CDiasAProc)
    
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 2) & "Empleados a procesar: " & CEmpleadosAProc
    Flog.writeline Espacios(Tabulador * 2) & "Fechas a procesar. Desde: " & FechaDesde & " Hasta: " & FechaHasta
    Progreso = 0
    
    'FGZ - Mejoras ----------
    Call Inicializar_Globales
    'FGZ - Mejoras ----------
    
    'EAM- Inicializa el total del progreso (v6.00)
    totalProgreso = cantProgreso

    
    Do While Not myrs.EOF
        Ternro = myrs!Ternro
        Empleado.Ternro = Ternro
        Empleado.Legajo = myrs!empleg
        'Flog.writeline
        'Flog.writeline Espacios(Tabulador * 1) & "Inicio Empleado:" & Ternro & " " & Fecha
        
        Fecha = FechaDesde
        
        ' para ese empleado voy desde la fecha desde hasta la fecha hasta
        ProgresoEmpleado = 0
        EmpleadoSinError = True
        
        'FGZ - Mejoras ----------
        Call Cargar_PoliticasIndividuales
        'FGZ - Mejoras ----------
        'Call borrar_todas(FechaDesde, FechaHasta) 'mdf borra todas las reg de la pol 21 para todos los dias, de una
        Do While Fecha <= FechaHasta And EmpleadoSinError

            'Reviso que el periodo de GTI al que pertenece la fecha no se encuentre cerrado
            StrSql = "SELECT * FROM gti_per "
            StrSql = StrSql & " WHERE pgtidesde <= " & ConvFecha(Fecha)
            StrSql = StrSql & " AND pgtihasta >= " & ConvFecha(Fecha)
            If rs_Per.State = adStateOpen Then rs_Per.Close
            OpenRecordset StrSql, rs_Per
            PeriodoCerrado = False
            Do While Not rs_Per.EOF
                If Not rs_Per!pgtiestado Then
                    PeriodoCerrado = True
                    Flog.writeline Espacios(Tabulador * 1) & "Periodo cerrado"
                End If
                rs_Per.MoveNext
            Loop
            
            If depurar Then
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------------------------------"
                Flog.writeline Espacios(Tabulador * 1) & "Inicio Empleado:" & Empleado.Legajo & " " & Fecha
            End If
            If Not PeriodoCerrado Or (PeriodoCerrado And ReprocesarFT) Then
                If (PeriodoCerrado And ReprocesarFT) Then
                    'If depurar Then
                        Flog.writeline Espacios(Tabulador * 2) & "Se reprocesará teniendo en cuenta todas las entrada fueras de termino aprobadas."
                    'End If
                End If
                
                'FGZ - 18/04/2006
                Call LimpiarJustificaciones
                Call PRC_30(Fecha, Ternro, True, depurar)
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 1) & "Periodo cerrado. Empleado:" & Ternro & " " & Fecha & " no se procesa"
                End If
            End If
            'EAM- Se comento la linea para mejorar la performance v(v6.00)
'            If EmpleadoSinError Then
'                'Actualizo tambien el porcentaje del empleado
'                If depurar Then
'                    Flog.writeline Espacios(Tabulador * 2) & "---> Actualizo tambien el porcentaje del empleado " & Now
'                End If
'                ProgresoEmpleado = ProgresoEmpleado + IncPorcEmpleado
'                StrSql = "UPDATE batch_empleado SET progreso = " & ProgresoEmpleado & " WHERE bpronro = " & NroProceso & " AND ternro = " & Ternro
'                objConnProgreso.Execute StrSql, , adExecuteNoRecords
'                If depurar Then
'                    Flog.writeline Espacios(Tabulador * 2) & "---> Porcentaje actualizado " & Now
'                End If
'            End If
            If depurar Then
                Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------------------------------"
                Flog.writeline
            End If
            Progreso = Progreso + IncPorc
            
            
            'EAM- Si el progreso es el configurado o esta por cumplirse el horario de espera del appserver actualiza en bach_proceso (v6.00)
            If (Progreso > totalProgreso) Or (DateDiff("n", Format(horaUpdateBD, "HH:mm:ss"), Format(Time, "HH:mm:ss")) >= (TiempoDeEsperaNoResponde - 1)) Then
                        
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 2) & "---> Actualizo progreso del proceso " & Now & " progreso: " & Progreso
                End If
            
                StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
                objConnProgreso.Execute StrSql, , adExecuteNoRecords
            
                totalProgreso = totalProgreso + cantProgreso
                horaUpdateBD = Time
            End If
            
            'Debug.Print Progreso
            Fecha = DateAdd("d", 1, Fecha)
        Loop
        

        'EAM- Si el empleado se proceso bien, lo agrega a una lista para luego borrarlo de batch_empleado
        If EmpleadoSinError Then
            ReDim Preserve arrEmpProcesadosOK(UBound(arrEmpProcesadosOK) + 1)
            arrEmpProcesadosOK(UBound(arrEmpProcesadosOK)) = Ternro
        End If

        'EAM- Esto ahora lo hace al final, no por cada empleado (v6.00)
'        'si el empleado se proceso por completo entonces lo borro de batch_empleados
'        StrSql = "SELECT progreso FROM batch_empleado WHERE bpronro = " & NroProceso & " AND ternro = " & Ternro
'        If rsEmpleado.State = adStateOpen Then rsEmpleado.Close
'        'rsEmpleado.Open StrSql, objConn
'        OpenRecordset StrSql, rsEmpleado
'        If Not rsEmpleado.EOF Then
'            If rsEmpleado!Progreso = 100 Then
'                If depurar Then
'                    Flog.writeline Espacios(Tabulador * 1) & "---> Empleado terminado, borro del batch" & Now
'                End If
'                StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso & " And Ternro = " & Ternro
'                objConn.Execute StrSql, , adExecuteNoRecords
'                If depurar Then
'                    Flog.writeline Espacios(Tabulador * 1) & "---> Empleado borrado del batch " & Now
'                End If
'            End If
'        End If
                                
        'siguiente empleado
        myrs.MoveNext
    Loop
    
    
    'EAM- Elimina todos los empleados procesados correctamente (v6.00)
    listEmpleadoOk = 0
    topeListaBorrar = 900
    nroMaxEmpBorrar = topeListaBorrar
    
    If (UBound(arrEmpProcesadosOK) > topeListaBorrar) Then
    
        For i = 1 To UBound(arrEmpProcesadosOK)
            listEmpleadoOk = listEmpleadoOk & "," & arrEmpProcesadosOK(i)
            
            'EAM- Borro los empleados armados en la lista
            If (nroMaxEmpBorrar <= i) Then
                StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso & " And Ternro IN (" & listEmpleadoOk & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                            
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 1) & "---> " & Now & " Se borraron los Empleados " & listEmpleadoOk
                End If
                listEmpleadoOk = "0"
                nroMaxEmpBorrar = nroMaxEmpBorrar + topeListaBorrar
            End If
        Next
    Else
        For i = 1 To UBound(arrEmpProcesadosOK)
            listEmpleadoOk = listEmpleadoOk & "," & arrEmpProcesadosOK(i)
        Next
        
        StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso & " And Ternro IN (" & listEmpleadoOk & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
                    
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "---> " & Now & " Se borraron los Empleados " & listEmpleadoOk
        End If
        listEmpleadoOk = "0"
    End If

    
    
    If objRs.State = adStateOpen Then objRs.Close
    
    'Actualizo el Btach_Proceso
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' -----------------------------------------------------------------------------------
    'FGZ - 22/09/2003
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso"
        Flog.writeline
    End If
    
    If Not HuboErrores Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "---> Proceso teminado, paso al historico ... " & Now
        End If
    
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        'rs_batch_proceso.Open StrSql, objConn
        OpenRecordset StrSql, rs_batch_proceso

        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_batch_proceso!bpronro & "," & rs_batch_proceso!btprcnro & "," & _
        ConvFecha(rs_batch_proceso!bprcfecha) & ",'" & rs_batch_proceso!IdUser & "'"
        
        If Not IsNull(rs_batch_proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprchora & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_batch_proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfechasta)
        End If
        If Not IsNull(rs_batch_proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_batch_proceso!bprcprogreso
        End If
        If Not IsNull(rs_batch_proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfecfin)
        End If
        If Not IsNull(rs_batch_proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_batch_proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_batch_proceso!empnro
        End If
        If Not IsNull(rs_batch_proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_batch_proceso!bprcPid
        End If
        If Not IsNull(rs_batch_proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_batch_proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_batch_proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_batch_proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_batch_proceso!bprcUrgente
        End If
        If Not IsNull(rs_batch_proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_batch_proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_batch_proceso!bprcHoraFinEj & "'"
        End If
        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        'rs_His_batch_proceso.Open StrSql, objConn
        OpenRecordset StrSql, rs_His_batch_proceso
        If Not rs_His_batch_proceso.EOF Then
            'Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "---> Historico Actualizado " & Now
        End If
        If rs_batch_proceso.State = adStateOpen Then rs_batch_proceso.Close
        If rs_His_batch_proceso.State = adStateOpen Then rs_His_batch_proceso.Close
    End If
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------
            
    'Cierro y libero todo
    If TransactionRunning Then MyRollbackTrans
    
    If objConn.State = adStateOpen Then objConn.Close
    If objConnProgreso.State = adStateOpen Then objConnProgreso.Close
    If CnTraza.State = adStateOpen Then CnTraza.Close

    If myrs.State = adStateOpen Then myrs.Close
    If rsEmpleado.State = adStateOpen Then rsEmpleado.Close
    If rs_batch_proceso.State = adStateOpen Then rs_batch_proceso.Close
    If rs_His_batch_proceso.State = adStateOpen Then rs_His_batch_proceso.Close
    If rs_Per.State = adStateOpen Then rs_Per.Close
    
    Set myrs = Nothing
    Set rsEmpleado = Nothing
    Set rs_batch_proceso = Nothing
    Set rs_His_batch_proceso = Nothing
    Set rs_Per = Nothing

Final:
    Flog.writeline Espacios(Tabulador * 0) & "Fin :" & Now
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.writeline "Cantidad de Lecturas en BD          : " & Cantidad_de_OpenRecordset
    Flog.writeline "Cantidad de llamadas a politicas    : " & Cantidad_Call_Politicas
'    Flog.writeline "Cantidad de llamadas a EsFeriado    : " & Cantidad_Feriados
'    Flog.writeline "Cantidad de llamadas a BuscarTurno  : " & Cantidad_Turnos
'    Flog.writeline "Cantidad de llamadas a BuscarDia    : " & Cantidad_Dias
'    Flog.writeline
'    Flog.writeline "Cantidad de dias procesados         : " & Cantidad_Empl_Dias_Proc
    Flog.writeline Espacios(Tabulador * 0) & "---------------------------------------------------------------------------------"
    Flog.Close
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
        StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error General', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
        objConnProgreso.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
End Sub



Public Sub PRC_30(Fecha As Date, Ternro As Long, P_Reproceso As Boolean, ByVal depurar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de calculo de Horario cumplido.
' Autor      : Maller José D.
' Fecha      : 18/10/02
' Ultima Mod.: FGZ - 30/08/2005 - integracion customizaciones
' Ultima Mod.: FGZ - 04/04/2006 - se agregó la llamada a la pol 700. Call Politica(700)
' Ultima Mod.: FGZ - 06/11/2008 - se agregó las llamadas a las pol 91 y 92
' Ultima Mod.: FGZ - 18/01/2010 - se agregó las llamadas a la pol 201 - Impuntualidad
' ---------------------------------------------------------------------------------------------
Dim CantH_Acum As Integer
Dim rs As New ADODB.Recordset
Dim rs_gti_Proc_Emp As New ADODB.Recordset

    On Error GoTo ce
Comienzo:
    
    UltimaRegInsertadaWFTurno = "N"
    'depurar = False
    Justif_Completa = False
    ok = True
    p_fecha = Fecha
    fecha_proceso = Fecha
    'FGZ - 07/11/2008 - agregué estas inicializaciones ---
    Genero_Sin_Control_Presencia = True
    Sigo_Generando = True
    Continua_Procesando = True
    'FGZ - 07/11/2008 - agregué estas inicializaciones ---
    
    'Debug.Print "   Legajo " & Empleado.Legajo & "(" & Empleado.Ternro & ") -->  " & p_fecha
    Call BlanquearVariables
    Call CreateTempTable(TTempWFDia)
    Call CreateTempTable(TTempWFDiaLaboral)
    Call CreateTempTable(TTempWFEmbudo)
    Call CreateTempTable(TTempWFTurno)
    Call CreateTempTable(TTempWFInputFT)
    
    'FGZ - Mejoras ----------
    Call Cargar_PoliticasEstructuras(p_fecha)
    'FGZ - Mejoras ----------
    
    Set objFechasHoras.Conexion = objConn
    
    'Chequeo en el Histórico de Estructura
    StrSql = " SELECT his_estructura.estrnro FROM his_estructura "
    StrSql = StrSql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro "
    StrSql = StrSql & " WHERE (tanro = " & lngAlcanGrupo & ") AND (ternro = " & Empleado.Ternro & ") AND "
    StrSql = StrSql & " (htetdesde <= " & ConvFecha(p_fecha) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then 'Si no lo encontré en el histórico
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "El Grupo de GTI no tiene alcance. SQL " & StrSql
            GeneraTraza Empleado.Ternro, p_fecha, "No posee turno de trabajo asociado"
        End If
        Exit Sub
    End If
    
'FGZ - 01/06/2007 - Mejoras ------
'    If depurar Then
'        Flog.writeline Espacios(Tabulador * 1) & "Call Procedimiento FechaProceso"
'    End If
'    Call FechaProceso
'    If depurar Then
'        Flog.writeline Espacios(Tabulador * 1) & "Termino Procedimiento FechaProceso"
'    End If
'FGZ - 01/06/2007 - Mejoras ------

    If P_Reproceso Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "Reproceso. Depuracion ..."
        End If
        'Elimino los registros de la tabla de horarios cumplidos
        MyBeginTrans
            'FGZ - 29/07/2009 - cambié el desmarcado ---------------
            'StrSql = "DELETE FROM gti_horcumplido WHERE ternro = " & Empleado.Ternro & " AND horfecrep = " & ConvFecha(fecha_proceso) & " AND hormanual = 0"
            StrSql = "DELETE FROM gti_horcumplido WHERE ternro = " & Empleado.Ternro & " AND horfecgen = " & ConvFecha(fecha_proceso) & " AND hormanual = 0"
            'FGZ - 29/07/2009 - cambié el desmarcado ---------------
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        
        'Blanqueo de las Registraciones Marcadas parcialmente
        'FGZ - 07/12/2010 --------------------------------------
        MyBeginTrans
            StrSql = "UPDATE gti_registracion SET hornro = 0 "
            'StrSql = StrSql & " , fechagen = " & FechaNula
            StrSql = StrSql & " WHERE ternro = " & Empleado.Ternro
            StrSql = StrSql & " AND hornro = -1 "
            StrSql = StrSql & " AND ( fechagen = " & ConvFecha(fecha_proceso)
            StrSql = StrSql & " OR (regestado = 'L' AND regfecha = " & ConvFecha(fecha_proceso) & "))"
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        'FGZ - 07/12/2010 --------------------------------------
        
        
               'FGZ - 17/01/2012 --------------------------------------------------------------------------
        'cambié el orden. Primero borro y de lo que quedá inicializo -------------------------------
        'FGZ - 18/03/2009 - Borro las Registraciones permamenetes que se generaron por la politica 21
        MyBeginTrans
            StrSql = "DELETE gti_registracion "
            StrSql = StrSql & " WHERE ternro = " & Empleado.Ternro
            'FGZ - 29/07/2009 - cambié el desmarcado ---------------
            'StrSql = StrSql & " AND (fechaproc = " & ConvFecha(fecha_proceso)
            StrSql = StrSql & " AND (fechagen = " & ConvFecha(fecha_proceso)
            'FGZ - 29/07/2009 - cambié el desmarcado ---------------
            'StrSql = StrSql & " OR ((regestado = 'I' OR regestado = 'L') AND regfecha = " & ConvFecha(fecha_proceso) & "))"
            StrSql = StrSql & " OR (regestado = 'L' AND fechagen = " & ConvFecha(fecha_proceso) & "))"
            StrSql = StrSql & " AND hornro = 21"
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        'FGZ - 18/03/2009 - Borro las Registraciones que se generaron por la politica 21
        
        'FGZ - 10/12/2010 - Borro las Registraciones temporales que se generaron por la politica 21
        MyBeginTrans
            StrSql = "DELETE gti_registracion "
            StrSql = StrSql & " WHERE ternro = " & Empleado.Ternro
            StrSql = StrSql & " AND (fechagen = " & ConvFecha(fecha_proceso)
            StrSql = StrSql & " OR ((regestado = 'I' OR regestado = 'L') AND regfecha = " & ConvFecha(fecha_proceso) & "))"
            'StrSql = StrSql & " OR (regestado = 'L' AND regfecha = " & ConvFecha(fecha_proceso) & "))"
            StrSql = StrSql & " AND hornro = -121"
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        'FGZ - 10/12/2010 - Borro las Registraciones temporales que se generaron por la politica 21
        'FGZ - 17/01/2012 --------------------------------------------------------------------------
        
       
        'Blanqueo de las Registraciones Marcadas en la Fecha
        MyBeginTrans
            'FGZ - 28/11/2006
            'StrSql = "UPDATE gti_registracion SET regestado = 'I', fechaproc = " & FechaNula & " WHERE ternro = " & Empleado.Ternro & " AND fechaproc = " & ConvFecha(fecha_proceso)
            StrSql = "UPDATE gti_registracion SET regestado = 'I', fechaproc = " & FechaNula
            'FGZ - 29/07/2009 - cambié el desmarcado ---------------
            StrSql = StrSql & " , fechagen = " & FechaNula
            'FGZ - 29/07/2009 - cambié el desmarcado ---------------
            StrSql = StrSql & " WHERE ternro = " & Empleado.Ternro
            'FGZ - 29/07/2009 - cambié el desmarcado ---------------
            'StrSql = StrSql & " AND ( fechaproc = " & ConvFecha(fecha_proceso)
            StrSql = StrSql & " AND ( fechagen = " & ConvFecha(fecha_proceso)
            'FGZ - 29/07/2009 - cambié el desmarcado ---------------
            StrSql = StrSql & " OR (regestado = 'L' AND regfecha = " & ConvFecha(fecha_proceso) & "))"
            ''FGZ - 19/05/2010 ------------ Control FT -------------
            '   creo que aca no corresponde. Testear
            'StrSql = StrSql & " AND (gti_registracion.ft = 0 OR (gti_registracion.ft = -1 AND gti_registracion.ftap = -1))"
            ''FGZ - 19/05/2010 ------------ Control FT -------------
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        
        
        'limpio
        MyBeginTrans
            StrSql = "DELETE FROM gti_proc_emp "
            StrSql = StrSql & " WHERE ternro =" & Ternro
            StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        
        'Borro los partes de asignacion horaria que se cargaron por alguna notificacion horarias (custom Sykes)
        MyBeginTrans
                       
            'cabeceras
            StrSql = "SELECT distinct gti_cabparte.gcpnro FROM gti_detturtemp "
            StrSql = StrSql & " INNER JOIN gti_cabparte ON gti_cabparte.gcpnro = gti_detturtemp.gcpnro "
            StrSql = StrSql & " WHERE (ternro = " & Ternro & ")"
            StrSql = StrSql & " AND (gttempdesde <= " & ConvFecha(Fecha) & ") "
            StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= gttemphasta)"
            StrSql = StrSql & " AND ttempobs = 'Parte automatico por notificaciones.'"
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                Do While Not rs.EOF
                    'Borro los detalles para el empleado en cuestion
                    StrSql = "DELETE FROM gti_detturtemp "
                    StrSql = StrSql & " WHERE ternro =" & Ternro
                    StrSql = StrSql & " AND gcpnro = " & rs!GCPNRO
                    StrSql = StrSql & " AND gttempdesde = " & ConvFecha(Fecha)
                    StrSql = StrSql & " AND ttempobs = 'Parte automatico por notificaciones.'"
                    objConn.Execute StrSql, , adExecuteNoRecords
            
                    StrSql = "DELETE FROM gti_cabparte "
                    StrSql = StrSql & " WHERE gcpnro = " & rs!GCPNRO
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    rs.MoveNext
                Loop
            End If
        MyCommitTrans
                
        'EAM- (5.48) - Elimino el desglose para el tercero
        MyBeginTrans
            StrSql = "DELETE FROM gti_desgloce_hc WHERE ternro = " & Ternro & " AND fecha = " & ConvFecha(Fecha)
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "Reproceso. Fin Depuracion"
        End If
    End If
    
    'Inicio la suma de horas del empleado, esto por si tiene un turno libre */
    CantH_Acum = 0
      
    'FGZ - 17/10/2014 ---------------------------
    'Se restauró llamada perdida a la poltica que setae si los relojes distinguen entradas y salidas
    Call Politica(5)
    'FGZ - 17/10/2014 ---------------------------
    
    'Politica para depurar reg repetidas en un lapso de tiempo dado.
    Call Politica(20)
    
    'FGZ - 17/03/2011 ----------------------
    'Busca las notificaciones Horarias y las pasa a parte de asignacion horaria
    Call Politica(4)
    'FGZ - 17/03/2011 ----------------------
   
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = CnTraza
    esFeriado = objFeriado.Feriado(p_fecha, Empleado.Ternro, depurar)
    
    Set objBTurno.Conexion = objConn
    Set objBTurno.ConexionTraza = CnTraza
    objBTurno.Buscar_Turno p_fecha, Empleado.Ternro, depurar
    Call initVariablesTurno(objBTurno)
    If Not tiene_turno And Not Tiene_Justif Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "Sin turno y no tiene justificacion"
        End If
        EmpleadoSinError = False
        Exit Sub
    End If
    
    If tiene_turno Then
        Set objBDia.Conexion = objConn
        Set objBDia.ConexionTraza = CnTraza
        objBDia.Buscar_Dia p_fecha, Fecha_Inicio, Nro_Turno, Empleado.Ternro, P_Asignacion, depurar
        Call initVariablesDia(objBDia)
        
        'FGZ - 20/04/2007 - Agregué estas lineas
        Horario_Movil = False
        Horario_Flexible_Rotativo = False
        Horario_Flexible_sinParte = False
        'Politica Turno de Hoario variable.
        Call Politica(14)
        If depurar Then
            Flog.writeline Espacios(Tabulador * 2) & "Horario Movil     ? " & Horario_Movil
            Flog.writeline Espacios(Tabulador * 2) & "Horario Flexible  ? " & Horario_Flexible_Rotativo
        End If
        If Not Horario_Movil And Not Horario_Flexible_Rotativo Then
            Call Horario_Teorico
        End If
    End If
    
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & " Entrada Teorica 1: " & FE1 & " - " & E1
        Flog.writeline Espacios(Tabulador * 1) & " Salida Teorica 1 : " & FS1 & " - " & S1
        Flog.writeline Espacios(Tabulador * 1) & " Entrada Teorica 2: " & FE2 & " - " & E2
        Flog.writeline Espacios(Tabulador * 1) & " Salida Teorica 2 : " & FS2 & " - " & S2
        Flog.writeline Espacios(Tabulador * 1) & " Entrada Teorica 3: " & FE3 & " - " & E3
        Flog.writeline Espacios(Tabulador * 1) & " Salida Teorica 3 : " & FS3 & " - " & S3
        Flog.writeline
    End If
    
    
    'FGZ - 15/01/2015 -------------------------
    'Habilita el control de feriados
    '   Esta llamada estaba mas abajo y debía estar antes
    Call Politica(25)
    'FGZ - 15/01/2015 -------------------------
    
    
    'Si es lunes llamo a la politica 460
    If Weekday(p_fecha) = 2 Then Call Politica(460)
    
    If depurar Then
        Flog.writeline
    End If
    If (Tipo_Turno = 1 Or Tipo_Turno = 2) And (Not Dia_Libre) And ((UsaFeriadoConControl And esFeriado) Or (Not esFeriado)) Then
        Call Politica(60)
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & " Despues Politica 60 "
        End If
    Else
        Call Politica(50)
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & " Despues Politica 50 "
        End If
    End If
    
    If depurar Then
        Flog.writeline Espacios(Tabulador * 1) & " VENTANA FECHA DESDE    : " & fecha_desde
        Flog.writeline Espacios(Tabulador * 1) & " VENTANA HORA  DESDE    : " & hora_desde
        Flog.writeline Espacios(Tabulador * 1) & " VENTANA FECHA HASTA    : " & fecha_hasta
        Flog.writeline Espacios(Tabulador * 1) & " VENTANA HORA  HASTA    : " & hora_hasta
        Flog.writeline Espacios(Tabulador * 1) & " CODIGO DEL TURNO       : " & Nro_Turno
        Flog.writeline Espacios(Tabulador * 1) & " TIPO DE TURNO          : " & Tipo_Turno
        Flog.writeline Espacios(Tabulador * 1) & " CODIGO DEL GRUPO       : " & Nro_Grupo
        Flog.writeline Espacios(Tabulador * 1) & " CODIGO DE FORMA DE PAGO: " & Nro_fpgo
        Flog.writeline Espacios(Tabulador * 1) & " FECHA DE INICIO        : " & Fecha_Inicio
        Flog.writeline Espacios(Tabulador * 1) & " TRABAJA                : " & Trabaja
        Flog.writeline Espacios(Tabulador * 1) & " ORDEN DIA              : " & Orden_Dia
        Flog.writeline Espacios(Tabulador * 1) & " CODIGO DE DIA          : " & Nro_Dia
        Flog.writeline Espacios(Tabulador * 1) & " CODIGO DE SUBTURNO     : " & Nro_Subturno
        Flog.writeline Espacios(Tabulador * 1) & " DIA LIBRE              : " & Dia_Libre
        Flog.writeline
    End If
    
    
    'FGZ - 18/03/2009 -------------
    'Crea las registraciones en los limites del dia cuando entra en un dia y sale en otro
    RegAuto_Permanentes = False
    Call Politica(21)   'Esta politica aun no está activa
    'FGZ - 18/03/2009 -------------
    
    'Posibles cambios de turnos no informados
    'FGZ - 30/07/2008 ------
    'Si el turno es libre esta politica no tiene sentido
    If Tipo_Turno = 1 Or Tipo_Turno = 2 Then
        Call Politica(15)
    Else
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & " El turno es libre. No se puede analizar posibles cambios de turno en base a su horario teorico porque el turno no lo tiene."
        End If
    End If
    
    'FGZ - 15/01/2015 -------------------------
    'Se sacó de acá y se puso mas arriba
    'Habilita el control de feriados
    'Call Politica(25)
    'FGZ - 15/01/2015 -------------------------
    
    
    '--------------------------------------------------MDF - 14/10/2015
    If Not Tiene_Justif And (FE1 <> FS1) Then
        'Tiene_Justif = BuscarTurnoNocturno 'es lo mismo que està al inicio de buscar turno
       buscar_justif_nocturna p_fecha, Empleado.Ternro, depurar
    End If
    '------------------------------------------------- MDF - 14/10/2015
    
    'FGZ - 06/11/2008 - Analizo inconsistencias --------------------------------
    If Tiene_Justif And Not justif_turno Then
        Call Buscar_Justif(p_fecha, Empleado.Ternro, Nro_Justif, No_Trabaja_just, fecha_desde, Hora_desde_aux, fecha_hasta, Hora_Hasta_aux, nro_jus_ent, nro_jus_sal)
        'Inconsistencia - Licencia/Novedad dia completo y registraciones
        If Justif_Completa And Tipo_de_Justificacion = 1 Then
            Continua_Procesando = True
            Call Politica(91)
            If Not Continua_Procesando Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 1) & "Inconsistencia - Licencia/Novedad dia completo y registraciones. Se configuró que no se continue procesando."
                    Flog.writeline
                End If
                Exit Sub
            End If
        End If
    End If
    'FGZ - 06/11/2008 - Analizo inconsistencias --------------------------------
    
    
    If Tiene_Justif And Not justif_turno Then
       ' Es una justificacion (PARCIAL Fija O DIA COMPLETO), sin turno, por lo tanto, Ajusta el Turno Asignado
       ' Busco la justif dentro de la ventana de Analisis
       Hora_desde_aux = hora_desde
       Hora_Hasta_aux = hora_hasta
       Call Buscar_Justif(p_fecha, Empleado.Ternro, Nro_Justif, No_Trabaja_just, fecha_desde, Hora_desde_aux, fecha_hasta, Hora_Hasta_aux, nro_jus_ent, nro_jus_sal)
       
       Trabaja = True 'Si tiene una Justificacion, ese día le hubiese correspondido trabajar
       
       'En Expofrut, los días libres (domingos, feriados) deben justificarse,
       'para evitar que se generen "horas domingo". O.D.A. 21/03/2003.
       If No_Trabaja_just Then
            If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Posee una Justificación de Día Completo"
            Call generar_dia_justificacion(Empleado.Ternro, p_fecha, Nro_Turno, Nro_Justif)
            ' GENERAR VALES COMEDOR
            'Call Politica(402)
            
            Call Politica(470)
            If Not Sigo_Generando Then
                'FGZ - 08/06/2007 - Agregué esto ------
                'Imputacion de horas sin control de prescencia
                'FGZ - 07/11/2008 - se le agregó este control
                If Genero_Sin_Control_Presencia Then
                    Call Politica(450)
                End If
                'FGZ - 08/06/2007 - Agregué esto ------
                'Call Insertar_GTI_Proc_Emp(Ternro, Fecha)
                'Exit Sub
                GoTo ParteFinal
            Else
                If depurar Then
                    Flog.writeline
                    Flog.writeline Espacios(Tabulador * 1) & "Sigo generando ... "
                    Flog.writeline
                End If
            End If
       End If
       If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Posee una Justificaci¢n de Día Parcial"
       Call generar_justificacion_Parcial(Nro_Turno, nro_jus_ent, nro_jus_sal)
'       'FGZ - 17/08/2006 -------------------
'       If Not Trabaja Then
'            Call generar_justificacion_Parcial(Nro_Turno, nro_jus_ent, nro_jus_sal)
'       End If
'       'FGZ - 17/08/2006 -------------------
    End If
    
    If Trabaja Then
       ' If (Tipo_Turno = 1 Or Tipo_Turno = 2) And _
       '     (Not Dia_Libre) And _
       '     ((UsaFeriadoConControl And esFeriado) Or (Not esFeriado)) Then
      If (Tipo_Turno = 1 Or Tipo_Turno = 2) And (Not Dia_Libre) And ((FeriadoConControldeAnormalidades Or Feriado_Laborable) Or (UsaFeriadoConControl And esFeriado) Or (Not esFeriado)) Then
  
            'Politicas de Ventanas de Embudo
            Call Politica(95)
            If Forma_embudo Then
                Call Politica(100)
                If E1 <> "" Then InsertarWFDia 1, E1, FE1, True: Cant_emb = 1
                If S1 <> "" Then InsertarWFDia 2, S1, FS1, False: Cant_emb = 2
                If E2 <> "" Then InsertarWFDia 3, E2, FE2, True: Cant_emb = 3
                If S2 <> "" Then InsertarWFDia 4, S2, FS2, False: Cant_emb = 4
                If E3 <> "" Then InsertarWFDia 5, E3, FE3, True: Cant_emb = 5
                If S3 <> "" Then InsertarWFDia 6, S3, FS3, False: Cant_emb = 6
            Else
                'FGZ - 16/09/2010 -------------------------------------------------------------------------------------
                'Call Ventana(Nro_Dia, p_fecha, Empleado.Ternro, hora_desde, fecha_desde, hora_hasta, fecha_hasta, P_Asignacion)
                If Horario_Movil Then
                    Call Ventana_Movil(Nro_Dia, p_fecha, Empleado.Ternro, hora_desde, fecha_desde, hora_hasta, fecha_hasta, P_Asignacion)
                Else
                    Call Ventana(Nro_Dia, p_fecha, Empleado.Ternro, hora_desde, fecha_desde, hora_hasta, fecha_hasta, P_Asignacion)
                End If
                'FGZ - 16/09/2010 -------------------------------------------------------------------------------------
            End If
        End If
        'FGZ - 05/11/2008 - se agregó una nueva version de la politica que determina si se continua procesando o no
        Continua_Procesando = True
        Call Politica(70)
        If Not Continua_Procesando Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 1) & Now & "Se configuró que no se continue procesando."
                End If
                Exit Sub
        End If
        'FGZ - 05/11/2008 - se agregó una nueva version de la politica que determina si se continua procesando o no
        
        If Not ok Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * 1) & Now & "Not OK."
            End If
            Exit Sub
        End If
        If (Tipo_Turno = 1 Or Tipo_Turno = 2) Then
            ok = True
            Call Politica(90)
            If Not ok Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * 1) & Now & "Not ok. Mas de 6 Registraciones. El sistema no puede tomar mas de 6 registraciones, depurarlas."
                End If
                Exit Sub
            End If
        End If
                

        If (Tipo_Turno <> 1 And Tipo_Turno <> 2) Then Call Politica(420)
        
        Existe_Reg = Existe_Registracion(Empleado.Ternro, fecha_desde, hora_desde, fecha_hasta, hora_hasta)
        Existe_Reg_LLamada = Existe_Registracion_LLamada(Empleado.Ternro, fecha_desde, hora_desde, fecha_hasta, hora_hasta)
        If depurar Then
            Flog.writeline Espacios(Tabulador * 1) & "" & IIf(Existe_Reg, "Existe Registracion", "No Existe Registracion")
        End If
            
        'Imputacion de horas sin control de prescencia
        Call Politica(450)
        
        If Not (Existe_Reg Or Existe_Reg_LLamada) And Not Dia_Libre And Not esFeriado Then
            ' Estuvo ausente todo el día
            ' Si el turno es libre se crea un reg. para que usa la política
           If (Tipo_Turno <> 1) And (Tipo_Turno <> 2) Then
                Select Case TipoBD
                Case 1: ' DB2
                    StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
                         Nro_Dia & "," & ConvFecha(p_fecha) & ",null,null)"
                Case 2: ' Informix
                     StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
                         Nro_Dia & "," & ConvFecha(p_fecha) & ",null,null)"
                Case 3: ' SQL Server
                    StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
                         Nro_Dia & "," & ConvFecha(p_fecha) & ",null,null)"
                End Select
                objConn.Execute StrSql, , adExecuteNoRecords
           End If
            'Tipo_de_Justificacion
            If Not Tiene_Justif Or (Tiene_Justif And Not Justif_Completa) Then
                Call Politica(30)
            End If
        'FGZ - 27/11/2006   -   Ler agregué esta parte para
        '      controlar Ausencias en Dias feriados que no caen en dias francos
        '      trabaja con la nueva politica 35
        Else
            If Not (Existe_Reg Or Existe_Reg_LLamada) And Not Dia_Libre And esFeriado Then
                'Estuvo ausente todo el día
                'Si el turno es libre se crea un reg. para que usa la política
               If (Tipo_Turno <> 1) And (Tipo_Turno <> 2) Then
                    Select Case TipoBD
                    Case 1: ' DB2
                        StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
                             Nro_Dia & "," & ConvFecha(p_fecha) & ",null,null)"
                    Case 2: ' Informix
                         StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
                             Nro_Dia & "," & ConvFecha(p_fecha) & ",null,null)"
                    Case 3: ' SQL Server
                        StrSql = "INSERT INTO " & TTempWFDia & "(Codigo,Fecha,Hora,Entrada) VALUES (" & _
                             Nro_Dia & "," & ConvFecha(p_fecha) & ",null,null)"
                    End Select
                    objConn.Execute StrSql, , adExecuteNoRecords
               End If
                  
                If Not Tiene_Justif Or (Tiene_Justif And Not Justif_Completa) Then
                    Call Politica(35)
                End If
            End If
        End If
            
        If Not (Existe_Reg Or Existe_Reg_LLamada) And (Dia_Libre Or esFeriado) Then
            'Politica de horas de sábado y domingo
            If Weekday(Fecha) = 1 Or Weekday(Fecha) = 7 Then
                usaSabadoDomingo = False
                Call Politica(550)
                If usaSabadoDomingo Then
                    Call Horas_Sabado_Domingo(Fecha)
                End If
            End If
            
            'Inserto dia procesado
            'Call Insertar_GTI_Proc_Emp(Ternro, Fecha)
            'Exit Sub
            GoTo ParteFinal
        Else
            If depurar Then
                Flog.writeline
                Flog.writeline Espacios(Tabulador * 1) & "Sigo generando ... "
                Flog.writeline
            End If
        End If
        
        If Existe_Reg Or Existe_Reg_LLamada Then
            'Registraciones a evaluar
            Call Politica(80)
            
            'C.A.T 22/08/08 No pagaba llamadas en Dias Francos o Feriados
            'Se movio la llamada a la politica 85 para que se ejecute siempre
            'es decir Laborable Franco o Feriado
            Call Politica(85)

            'FGZ - 02/08/2010 - Control de FR especifica --------------------------------
            Call Politica(111)
            'FGZ - 02/08/2010 - Control de FR especifica --------------------------------

            'FGZ - 06/11/2008 - Analizo inconsistencias --------------------------------
            If Tiene_Justif And Not justif_turno Then
                If Not Justif_Completa And Tipo_de_Justificacion = 2 Then
                    'Inconsistencia - Licencias/Novedades Parciales y registraciones superpuestas
                    Continua_Procesando = True
                    Call Politica(92)
                    If Not Continua_Procesando Then
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * 1) & "Inconsistencia - Licencias/Novedades Parciales y registraciones superpuestas. Se configuró que no se continue procesando."
                            Flog.writeline
                        End If
                        Exit Sub
                    End If
                End If
            End If
            'FGZ - 06/11/2008 - Analizo inconsistencias --------------------------------


            If (Tipo_Turno = 1 Or Tipo_Turno = 2) And (Not Dia_Libre) And ((UsaFeriadoConControl And esFeriado) Or (Not esFeriado)) Then
                If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Generación del Horario Normal"
                If Forma_embudo Then
                    'ARMAR EMBUDO
                        
                    'FGZ - 28/09/2011 --------------------------------
                        Call Tolerancias
                        
                        'Impuntualidad
                        Call Politica(201)
                        
                        'Generación de Llegadas Tardes por Diferencia
                        Call Politica(190)
    
                        'Generación de Salidas Temprano por Diferencia
                        Call Politica(200)
                        
                    'FGZ - 28/09/2011 --------------------------------
                Else
                    'recorre las registraciones
                    'C.A.T 22/08/08 - Eliminar la llamada a la politica 85 en este punto
                    'Call Politica(85)
                    
                    Call Tolerancias

                    'FGZ - 30/04/2007 - no debe generar este tipo de anormalidades cuando el dia es feriado o franco, a excepcion de
                    '                   los turnos Moviles, Fexibles o Rotativos flexibles (Politica 14 V1, V2 y V3 respectivamente)
                    If Not Horario_Movil And Not Horario_Flexible_Rotativo Then
                        If Not esFeriado Then
                            'Impuntualidad
                            Call Politica(201)
                            
                            'Generación de Llegadas Tardes por Diferencia
                            Call Politica(190)
        
                            'Generación de Salidas Temprano por Diferencia
                            Call Politica(200)
                            
                        Else
                            'FGZ - 27/01/2009 - Se le agregó este control
                            If Not Dia_Libre And UsaFeriadoConControl Then
                                'FGZ - 22/01/2015 -------------------------
                                If FeriadoConControldeAnormalidades Or Feriado_Laborable Then
                                    'Impuntualidad
                                    Call Politica(201)
                                    
                                    'Generación de Llegadas Tardes por Diferencia
                                    Call Politica(190)
                
                                    'Generación de Salidas Temprano por Diferencia
                                    Call Politica(200)
                                End If
                                'FGZ - 22/01/2015 -------------------------
                            End If
                        End If
                    Else
                        'Impuntualidad
                        Call Politica(201)
                        
                        'Generación de Llegadas Tardes por Diferencia
                        Call Politica(190)
    
                        'Generación de Salidas Temprano por Diferencia
                        Call Politica(200)
                        
                    End If
                End If
                
                
                'FGZ - 20/01/2015 -----------------------------------
                'If (UsaFeriadoConControl And esFeriado) Then
                '    Call Politica(1001)
                '    GeneraLaborable_y_Feriado = False
                '    Call Politica(205)
                '    If GeneraLaborable_y_Feriado Then
                '        Call Politica(1000)
                '    End If
                'Else
                '    Call Politica(1000)
                'End If
                
                GeneraLaborable_y_Feriado = False
                Call Politica(205)
                If (UsaFeriadoConControl And esFeriado) Then
                    If Feriado_Laborable Then
                        Select Case FP_Feriado
                        Case 1: 'Solo Feriado
                            Call Politica(1001)
                        Case 2: 'Solo Laborable
                            Call Politica(1000)
                        Case 3: 'Feriado y Laborable
                            Call Politica(1001)
                            Call Politica(1000)
                        Case Else
                            Call Politica(1001)
                        End Select
                    Else
                        Call Politica(1001)
                        If FP_Feriado = 3 Then
                            Call Politica(1000)
                        End If
                    End If
                Else
                    Call Politica(1000)
                End If
                'FGZ - 20/01/2015 -----------------------------------
                
            Else
                ' Si es turno libre
                If depurar Then
                    GeneraTraza Empleado.Ternro, p_fecha, "Generación del Horario Libre"
                End If

                Call Politica(1001)
            End If
        Else
            ' Si no registraciones hay que generar igual el horario libre
            If (Tipo_Turno <> 1 And Tipo_Turno <> 2) And (Not Dia_Libre) And (Not esFeriado) Then
                If depurar Then
                    GeneraTraza Empleado.Ternro, p_fecha, "Generación del Horario Libre"
                End If
                Call Politica(1001)
            End If
        End If
        
        'FGZ - 27/01/2009 --------------------------------------
        '   le cambié esta llamada pues
        '   si el dia es Laborable y feriado tambien debe analizar la politica 410
'        If (Existe_Reg Or Existe_Reg_LLamada) And (Tipo_Turno = 1 Or Tipo_Turno = 2) And (Not Dia_Libre) And (Not esFeriado) Then
'            Call Politica(410)
'        End If
        
        If (Existe_Reg Or Existe_Reg_LLamada) And (Tipo_Turno = 1 Or Tipo_Turno = 2) And (Not Dia_Libre) Then
            If (UsaFeriadoConControl And esFeriado) Or (Not esFeriado) Then
                'dia laborable o feriado pero laborable
                Call Politica(410)
            End If
        End If
        'FGZ - 27/01/2009 --------------------------------------
        
        'EAM- (v5.44) - Hs de Insalubridad
        Call Politica(720)
        
        'FGZ - 26/05/2009 --------------------------------------
        'Control de X con Hs Nocturnas
        Call Politica(600)
        'Control de X con Hs Nocturnas
        Call Politica(597)
        'Control de X con Hs Nocturnas
        Call Politica(598)
        'FGZ - 04/10/2010 --------------------------------------
        
        'FGZ - 30/09/2010 --------------------------------------
        'Control de Hs Nocturnas
        Call Politica(599)
        'FGZ - 30/09/2010 --------------------------------------
        
        
        'Prueba de Generar las Horas de Descuento. Esta política va luego de la
        'generación del HC, porque en ese momento es cuando se setea la fecha de procesamiento
        Call Politica(430)

        'Genero las Justificaciones Parciales Variables
        Call Politica(480)
        
        'Genero las Justificaciones Parciales Automaticas
        'FGZ - 27/10/2008 - Selo se generarán las Justificaciones de almuerzo variable si tiene registraciones
        If (Existe_Reg Or Existe_Reg_LLamada) Then
            'Call Politica(400)
            Call Politica(401)
        End If

        'FGZ - 21/10/2011 ----------------
        'Se analiza la pol 400 y dependiendo de la version controla si hay registraciones o no.
        'Antes solo llamaba a la politica 400 si habia registraciones
        Call Politica(400)
        'FGZ - 21/10/2011 ----------------
        
        'Politica de generacion de horas de sábado y domingo
        If Weekday(Fecha) = 1 Or Weekday(Fecha) = 7 Then
            usaSabadoDomingo = False
            Call Politica(550)
            If usaSabadoDomingo Then
                Call Horas_Sabado_Domingo(Fecha)
            End If
        End If
        'FGZ - 06/12/2006 - Nueva version de la politica 220
        If Not Dia_Libre And Not esFeriado Then
            Call Politica(220)
        End If
    End If
    
    'FGZ - -19/05/2008 - Politica de Horas de LLamadas para CARGIL
    Call Politica(510)
    'FGZ - -19/05/2008 - Politica de Horas de LLamadas para CARGIL
    
    'FGZ - 21/10/2010 - Descuentos Autorizados
    Call Politica(490)
    
    'FGZ - 19/10/2010 - Paros Gremiales (MONRESA)
    Call Politica(495)
ParteFinal:
    'FGZ - 02/06/2009 --------------------------------------
    'VALES
    Call Politica(402)
    'FGZ - 02/06/2009 --------------------------------------

    'FGZ - 07/07/2008  - El antiguo proceso de Feriados ahora se insertó como una nueva politica (499)
    If esFeriado Then
        If Sigo_Generando Or (Not Sigo_Generando And Genero_Sin_Control_Presencia) Then
            Call Politica(499)
        End If
    End If
    
    'Inserto dia procesado
    Call Insertar_GTI_Proc_Emp(Ternro, Fecha)
    
    '--------------------------------------------------------------------
    'Esto no se usa mas pero puede que queden algunas versiones antiguas
    '   que lo sigan utilizando por lo cual se deja
    usaTurnoTrasnoche = False
    Call Politica(560)
    If usaTurnoTrasnoche Then
        Call Horario_Trasnoche(Fecha, Ternro)
    End If
    '--------------------------------------------------------------------
    
    'EAM- Nueva politic de desglose.
    Call Politica(730)
    
    
    '*22657
    
    
    'FGZ - 03/04/2006 - CUSTOM HALLIBURTON ------------------------------
    Call Politica(700)
    'FGZ - 03/04/2006 - CUSTOM HALLIBURTON ------------------------------
    
    
    Call ActualizarFT(1, Fecha, Ternro)
    
    'FGZ - 02/11/2009 - Borro las Registraciones que se generaron por la politica 21
    If Not RegAuto_Permanentes Then
        'FGZ - 10/12/2010 - Borro las Registraciones temporales que se generaron por la politica 21
        MyBeginTrans
            StrSql = "DELETE gti_registracion "
            StrSql = StrSql & " WHERE ternro = " & Empleado.Ternro
            StrSql = StrSql & " AND (fechagen = " & ConvFecha(fecha_proceso)
            StrSql = StrSql & " OR ((regestado = 'I' OR regestado = 'L') AND regfecha = " & ConvFecha(fecha_proceso) & "))"
            'StrSql = StrSql & " OR (regestado = 'L' AND regfecha = " & ConvFecha(fecha_proceso) & "))"
            StrSql = StrSql & " AND hornro = -121"
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        'FGZ - 10/12/2010 - Borro las Registraciones temporales que se generaron por la politica 21
        
        MyBeginTrans
            StrSql = "DELETE gti_registracion "
            StrSql = StrSql & " WHERE ternro = " & Empleado.Ternro
            StrSql = StrSql & " AND (fechagen = " & ConvFecha(fecha_proceso)
            StrSql = StrSql & " OR (regestado = 'L' AND regfecha = " & ConvFecha(fecha_proceso) & "))"
            StrSql = StrSql & " AND hornro = 21"
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        'FGZ - 18/03/2009 - Borro las Registraciones que se generaron por la politica 21
    End If
    
    
    
    'cierro y libero
    Exit Sub
ce:
    HuboErrores = True
    EmpleadoSinError = False
    Flog.writeline Espacios(Tabulador * 1) & "Error. Empleado abortado " & Now
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
    Flog.writeline Espacios(Tabulador * 1) & " ------------------------------"
End Sub


'-------------------------------------------------------------------------------MDF just nocturnas
Public Sub buscar_justif_nocturna(Fecha As Date, lngTernro As Long, depurar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el turno del empleado en la fecha.
' Autor      : MDF
' Fecha      : 13/10/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objRsTurnoFPago As New ADODB.Recordset
Dim objRsReldTur As New ADODB.Recordset
Dim objRsTurnoFPGrupo As New ADODB.Recordset
Dim objLicAut As New ADODB.Recordset
'Dim Parte_FT As Boolean
Dim rs_FT As New ADODB.Recordset
Dim rs_Firma As New ADODB.Recordset
Dim Firmado As Boolean
Dim Nov_Firmada As Boolean
Dim Encontro_Jus As Boolean

Dim rs As New ADODB.Recordset
Dim rs_Tur As New ADODB.Recordset

   ' PFecha = Fecha
    Tiene_Justif = False
    justif_turno = False
    
    StrSql = ""
    Select Case TipoBD
    Case 4:
        StrSql = "SELECT * FROM ("
    End Select
    StrSql = StrSql & "(SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN emp_lic ON gti_justificacion.juscodext = emp_lic.emp_licnro "
    StrSql = StrSql & " WHERE (ternro = " & lngTernro & ") "
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(DateAdd("d", 1, Fecha)) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND emp_lic.licestnro = 2"
    StrSql = StrSql & " AND jussigla = 'LIC'"
    StrSql = StrSql & " AND juseltipo = 2 " '----
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
    StrSql = StrSql & " )UNION ("
    StrSql = StrSql & " SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN gti_novedad ON gti_justificacion.juscodext = gti_novedad.gnovnro "
    StrSql = StrSql & " WHERE (Ternro = " & lngTernro & ")"
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(DateAdd("d", 1, Fecha)) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND jussigla = 'NOV'"
    StrSql = StrSql & " AND jussigla <> 'ALM'"
    StrSql = StrSql & " AND juseltipo = 2 " '----
    StrSql = StrSql & " AND (gti_novedad.ft = 0 OR (gti_novedad.ft = -1 AND gti_novedad.ftap = -1))"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " )UNION ("
    StrSql = StrSql & " SELECT gti_justificacion.* FROM gti_justificacion "
    StrSql = StrSql & " WHERE (Ternro = " & lngTernro & ")"
    StrSql = StrSql & " AND (jusdesde <= " & ConvFecha(DateAdd("d", 1, Fecha)) & ")"
    StrSql = StrSql & " AND (" & ConvFecha(Fecha) & " <= jushasta)"
    StrSql = StrSql & " AND jussigla <> 'LIC'"
    StrSql = StrSql & " AND jussigla <> 'NOV'"
    'FGZ - 14/8/2008 - le agregué esta linea por las justificaciones automaticas de la Politica 400
    StrSql = StrSql & " AND jussigla <> 'ALM'"
     StrSql = StrSql & " AND juseltipo = 2 " '----
    'FGZ - 14/8/2008 - le agregué esta linea por las justificaciones automaticas de la Politica 400
    StrSql = StrSql & ")"
    Select Case TipoBD
    Case 4:
        StrSql = StrSql & ")"
    End Select
    StrSql = StrSql & " ORDER BY juseltipo "
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        'Esto no quiere decir que no hay justificaciones sino que no hay o no hay aprobadas
    Else
        'FGZ - 10/06/2011 -------------------------------------------------------
        ' Se agregó control de firmas para Novedades horarias (que no sean automaticas - pol 400)
        Encontro_Jus = False
        Do While Not objRs.EOF And Not Encontro_Jus
            If objRs!jussigla = "NOV" Then
                'Verificar si esta en el NIVEL FINAL DE FIRMA ACTIVO para Novedades Horarias de GTI
                If Firma_Novedades Then
                    StrSql = "SELECT * FROM cysfirmas "
                    StrSql = StrSql & " WHERE cysfirfin = -1"
                    StrSql = StrSql & " AND cysfircodext = '" & objRs!juscodext & "' "
                    StrSql = StrSql & " AND cystipnro = 7"
                    OpenRecordset StrSql, rs
                    If rs.EOF Then
                        Nov_Firmada = False
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * 6) & "Hay una novedad horaria NO aprobado. Se descarta."
                        End If
                    Else
                        Nov_Firmada = True
                        Encontro_Jus = True
                    End If
                Else
                    Nov_Firmada = True
                    Encontro_Jus = True
                End If
            Else
                Nov_Firmada = True
                Encontro_Jus = True
            End If
                
            'FGZ - 10/06/2011 -------------------------------------------------------
            If Nov_Firmada Then
                Tiene_Justif = True
                Nro_Justif = objRs!jusnro
                If Not IsNull(objRs!turnro) Then
                    StrSql = "SELECT * FROM gti_turno WHERE turnro = " & objRs!turnro
                    OpenRecordset StrSql, rs_Tur
                    If Not rs_Tur.EOF Then
                       ' Tur.tiene_turno = True
                       ' Tur.Numero = rs_Tur!turnro
                       ' Tur.Nombre = Trim(rs_Tur!turdesabr)
                       ' Tur.Tipo = rs_Tur!TipoTurno
                        justif_turno = True
                        'FGZ - 10/06/2011 -------------------------------------------------------
                        'le saque el exit y lo manejo abajo con la condicion de si tiene justificacion
                        'Exit Sub
                    End If
                End If
            End If
            objRs.MoveNext
        Loop
    
    End If
End Sub
'-------------------------------------------------------------------------------MDF fin justif nocturnas


Private Sub Buscar_Justif(Fecha As Date, p_ternro As Long, Nro_Justif As Long, ByRef No_Trabaja_just As Boolean, ByRef fecha_desde As Date, ByRef hora_desde As String, ByRef fecha_hasta As Date, ByRef hora_hasta As String, ByRef nro_jus_ent As Long, ByRef nro_jus_sal As Long)
Dim Entro As Boolean
Dim SQLJustif As String
Dim AuxE As String
Dim AuxS As String
Dim AuxFE As Date
Dim AuxFS As Date

Dim Cant_Int As Double
Dim Cant_HC1 As Double
Dim Cant_HC2 As Double
Dim Caso As String
Dim JFDesde As Date
Dim JFHasta As Date
Dim PrimerJust As Long
    'FGZ - 22/08/2006
    'No estaban inicializadas estas variables y procesando dias consecutivos hacen macanas
    nro_jus_ent = 0
    nro_jus_sal = 0
    PrimerJust = 0
    lista_justificacionesparciales = "0" 'MDf
    ' Verificar Justificaci¢n a la fecha con un turno especial
    StrSql = "SELECT jusdiacompleto,juseltipo FROM gti_justificacion WHERE gti_justificacion.jusnro = " & Nro_Justif
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'FGZ - 02/12/2004
        Justif_Completa = CBool(objRs!jusdiacompleto)
        
        'FGZ - 26/12/2005
        'Siempre las justificaciones estan cargadas como de dia completo ==>
        'segun el tipo las pongo como de dia completo o no
        Select Case objRs!juseltipo
        Case 1:
            Justif_Completa = True
        Case 2:
            Justif_Completa = False
        Case 3:
            Justif_Completa = False
        Case Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * 2) & "Tipo de justificacion desconocido: " & objRs!juseltipo
            End If
        End Select
        
        'Tipo_de_Justificacion = objRs!juseltipo = 1
        'FGZ - 18/04/2006
        Tipo_de_Justificacion = objRs!juseltipo
        
        ' Existe Una Justificacion Particular
        If objRs!juseltipo = 1 Then 'And (CDate(FE1) = CDate(FS1)) Then  'No trabaja en todo el dia
            No_Trabaja_just = True
        Else
            No_Trabaja_just = False
            
           'Si no es dia completo puede haber + de 1 Justificacion, 1 de Entrada y otra
           'de Salida, por lo tanto se deben ordenar por Hora para determinar las menores
           'y se deben buscar en la ventana ppal sino puede haber solapamiento
            
            'FGZ - 26/08/2014 ------------
            'SQLJustif = "SELECT jusnro,jushoradesde,jushorahasta FROM gti_justificacion WHERE (ternro = " & p_ternro & ") AND " & _
            '         " (juseltipo = 2 or juseltipo = 3) AND (jusdesde <= " & ConvFecha(Fecha) & ") AND (jushasta >= " & ConvFecha(Fecha) & ") " & _
            '         " ORDER BY ternro,jusdesde,jushasta,jushoradesde,jushorahasta"
            
            AuxE = E1
            AuxFE = FE1
            AuxS = S1
            AuxFS = FS1
            If Not EsNulo(FS2) And Not EsNulo(S2) Then
                AuxS = S2
                AuxFS = FS2
            End If
            If Not EsNulo(FS3) And Not EsNulo(S3) Then
                AuxS = S3
                AuxFS = FS3
            End If
            'FGZ - 30/12/2014 ------------
            'SQLJustif = "SELECT jusnro,jusdesde,jushasta,jushoradesde,jushorahasta,juseltipo FROM gti_justificacion WHERE (ternro = " & p_ternro & ") AND " & _
            '         " (juseltipo = 2 or juseltipo = 3) AND (jusdesde <= " & ConvFecha(AuxFS) & ") AND (jushasta >= " & ConvFecha(AuxFE) & ") " & _
            '         " ORDER BY ternro,jusdesde,jushasta,jushoradesde,jushorahasta"
            
          
            '------------------------------ MDF 12/05/2015
            StrSql = "SELECT jusnro,jusdesde,jushasta,jushoradesde,jushorahasta,juseltipo FROM gti_justificacion WHERE (ternro = " & p_ternro & ") AND " & _
                   " (juseltipo = 2) AND (jusdesde <= " & ConvFecha(AuxFS) & ") AND (jushasta >= " & ConvFecha(AuxFE) & ") " & _
                   "and jusnro not in (" & lista_procesadas & ")" & _
                   " ORDER BY ternro,jusdesde,jushasta,jushoradesde,jushorahasta"
                     
            'StrSql = "SELECT jusnro,jusdesde,jushasta,jushoradesde,jushorahasta,juseltipo,jussigla FROM gti_justificacion WHERE (ternro = " & p_ternro & ") AND " & _
            '         " (juseltipo = 2 or juseltipo = 1) AND (jusdesde <= " & ConvFecha(AuxFS) & ") AND (jushasta >= " & ConvFecha(AuxFE) & ") " & _
            '         " ORDER BY ternro,jusdesde,jushasta,jushoradesde,jushorahasta"
            '         'Puse jussigla tb mdf
            '-----------------------------MDF 12/05/2015
            
            'FGZ - 26/08/2014 ------------
            'OpenRecordset SQLJustif, objRs
            OpenRecordset StrSql, objRs
            
            Do While Not objRs.EOF
                'FGZ - 07/10/2014 -------------
                'If ((objRs!jushoradesde <= hora_desde Or p_fecha > fecha_desde) And _
                '   (objRs!jushorahasta <= hora_hasta Or p_fecha < fecha_hasta)) And _
                '   (objRs!jushorahasta >= hora_desde Or p_fecha > fecha_desde) _
                'Then Entro = True
                'If ((objRs!jushoradesde >= hora_desde Or p_fecha > fecha_desde) And _
                '   (objRs!jushorahasta >= hora_hasta Or p_fecha < fecha_hasta)) And _
                '   (objRs!jushoradesde <= hora_hasta Or p_fecha < fecha_hasta) _
                'Then Entro = True
                'If ((objRs!jushoradesde >= hora_desde Or p_fecha > fecha_desde) And _
                '   (objRs!jushorahasta <= hora_hasta Or p_fecha < fecha_hasta)) _
                'Then Entro = True
                'If ((objRs!jushoradesde <= hora_desde Or p_fecha > fecha_desde) And _
                '   (objRs!jushorahasta >= hora_hasta Or p_fecha < fecha_hasta)) _
                'Then Entro = True
                
                'FGZ - 07/10/2014 -------------
                Entro = False
                If Not EsNulo(AuxE) And Not EsNulo(AuxS) Then
                  If Not IsNull(objRs!jushoradesde) Then
                    Call CalcularInterseccionHoras3(AuxFE, AuxE, AuxFS, AuxS, objRs!jusdesde, objRs!jushoradesde, objRs!jushasta, objRs!jushorahasta, 8, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
                    If Cant_Int = 0 And AuxFE <> AuxFS Then
                       Call CalcularInterseccionHoras3(DateAdd("d", -1, AuxFE), AuxE, AuxFE, AuxS, objRs!jusdesde, objRs!jushoradesde, objRs!jushasta, objRs!jushorahasta, 8, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
                    End If
                  Else 'MDF 12/05/2015 - Si vienen null hora desde y hasta, es todo el dia
                     Call CalcularInterseccionHoras3(AuxFE, AuxE, AuxFS, AuxS, objRs!jusdesde, "0000", objRs!jushasta, "2400", 12, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
                  End If
                  
                    If Cant_Int <> 0 Then
                        Entro = True
                        'FGZ - 07/10/2014 --------------
                            
                        'FGZ - 07/10/2014 --------------
                        If PrimerJust = 0 Then
                            
                            PrimerJust = objRs!jusnro
                            
                            If (objRs!jushorahasta <> "0000") And (objRs!jushorahasta > hora_desde) Then
                               
                                hora_desde = objRs!jushorahasta
                            End If
                            nro_jus_ent = objRs!jusnro
                        Else
                            'Si es la ultima
                            If (objRs!juseltipo = 2) Then
                             lista_justificacionesparciales = lista_justificacionesparciales & "," & objRs!jusnro
                            End If
                            
                            objRs.MoveNext
                            If objRs.EOF Then
                                
                                objRs.MovePrevious
                                If (objRs!jushoradesde <> "0000") And (objRs!jushoradesde < hora_hasta) Then
                                    hora_hasta = objRs!jushoradesde
                                End If
                                nro_jus_sal = objRs!jusnro
                            Else
                                
                                objRs.MovePrevious
                            End If
                        End If
                    End If
                End If
                'FGZ - 07/10/2014 -------------
                
                'FGZ - 07/10/2014 -------------
                'If Entro Then
                '    'First_of (jusdesde)
                '    'If FIRST_OF(objrs!jusnro, "gti_justificacion", "jusnro", "ternro,jusdesde,jushasta,jushoradesde,jushorahasta", "ternro", objrs!ternro) Then
                '    'If FIRST_OF(SQLJustif, "ternro", empleado.Ternro) Then
                '    If Primer_Justificacion(SQLJustif, objRs!jusnro, Empleado.Ternro) Then
                '        If (objRs!jushorahasta <> "0000") And (objRs!jushorahasta > hora_desde) Then
                '            hora_desde = objRs!jushorahasta
                '        End If
                '        nro_jus_ent = objRs!jusnro
                '    End If
                '    'Last_of (jushasta)
                '    'If LAST_OF(objrs!jusnro, "gti_justificacion", "jusnro", "ternro,jusdesde,jushasta,jushoradesde,jushorahasta", "ternro", objrs!ternro) Then
                '    'If LAST_OF(SQLJustif, "ternro", empleado.Ternro) Then
                '    If Ultima_Justificacion(SQLJustif, objRs!jusnro, Empleado.Ternro) Then
                '        If (objRs!jushoradesde <> "0000") And (objRs!jushoradesde < hora_hasta) Then
                '            hora_hasta = objRs!jushoradesde
                '        End If
                '        nro_jus_sal = objRs!jusnro
                '    End If
                'End If
                'FGZ - 07/10/2014 -------------
                
                objRs.MoveNext
            Loop
        End If
        If hora_desde <> "" And hora_hasta <> "" Then
            If Not objFechasHoras.ValidarHora(hora_desde) Then Exit Sub
            If Not objFechasHoras.ValidarHora(hora_hasta) Then Exit Sub
        End If
        If depurar Then
            GeneraTraza Empleado.Ternro, p_fecha, "No trabaja justificado", Str(No_Trabaja_just)
            GeneraTraza Empleado.Ternro, p_fecha, "Justificación Parcial Ent: ", Str(nro_jus_ent)
            GeneraTraza Empleado.Ternro, p_fecha, "Justificación Parcial Sal: ", Str(nro_jus_sal)
        End If
    End If
End Sub
Private Sub generar_dia_justificacion(p_ternro As Long, p_fecha As Date, p_turnro As Long, p_justif As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Nueva version del procedimiento para generar las justificaciones.
'              Motivo: version anterior procesa una sola justificacion.
' Autor      : FGZ
' Fecha      : 02/05/2005
' Ultima Mod.: FGZ - 17/11/2005
' Descripcion: Levanta solo las licencias en estado AUTORIZADAS
' ---------------------------------------------------------------------------------------------
Dim j_tipo As String
Dim objRsJustif As New ADODB.Recordset
Dim aux_canthorasjust As Single
Dim max_horas As Single
Dim horas_min As Single

    Aux_Tipohora = 0
    Call buscar_horas_turno(aux_canthorasjust, max_horas, horas_min)
    
    Total_horas = aux_canthorasjust
    
    StrSql = "SELECT gti_justificacion.jusnro,gti_justificacion.jussigla,gti_justificacion.juscodext, gti_justificacion.juscanths, gti_tipojust.thnro "
    StrSql = StrSql & " FROM gti_justificacion "
    StrSql = StrSql & " INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro "
    StrSql = StrSql & " WHERE gti_justificacion.jusnro = " & p_justif
    OpenRecordset StrSql, objRsJustif
    If objRsJustif.EOF Then
       Exit Sub
    End If
    
        Select Case objRsJustif!jussigla
            Case "NOV", "ALM"
                j_tipo = "NOVEDAD"
                StrSql = "SELECT gti_tiponovedad.thnro FROM gti_novedad INNER JOIN gti_tiponovedad ON gti_novedad.gtnovnro = gti_tiponovedad.gtnovnro WHERE (gnovotoa = " & p_ternro & ") AND "
                StrSql = StrSql & " (gnovnro = " & objRsJustif!juscodext & ")"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                StrSql = StrSql & " AND (gti_novedad.ft = 0 OR (gti_novedad.ft = -1 AND gti_novedad.ftap = -1))"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    ' MESSAGE "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)"
                    Exit Sub
                End If
                Aux_Tipohora = objRs!thnro
                Call Politica(440)
            Case "LIC"
                j_tipo = "LICENCIA"
                StrSql = "SELECT tipdia.thnro,tipdia.tdnro FROM emp_lic "
                StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro "
                StrSql = StrSql & " WHERE (empleado = " & p_ternro & ") "
                StrSql = StrSql & " AND (emp_licnro = " & objRsJustif!juscodext & ")"
                'FGZ - 17/11/2005
                StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                'FGZ - 19/05/2010 ------------ Control FT -------------
                StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    ' MESSAGE "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)" VIEW_AS ALERT_BOX. */
                    Exit Sub
                End If
                aux_TipoDia = objRs!tdnro
                Aux_Tipohora = objRs!thnro ' Tipo de Hora equivalente de la Licencia
                Call Politica(440)
            Case "CUR"
                j_tipo = "CURSO"
                Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para el Curso
                aux_canthorasjust = objRsJustif!juscanths
            Case "SUS"
                j_tipo = "SUSPENCION"
                Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para Suspenci¢n
                aux_canthorasjust = objRsJustif!juscanths
        End Select
        If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Tipo de la Justificación", j_tipo
        
        If Aux_Tipohora = 0 Then
            ' La Justificacion no se paga, no tiene tipo de hora asignado
            If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Falta Tipo de Hora en la justificación", Str(p_justif)
            Exit Sub
        Else
            'Call ValidarTipoDeHora(aux_Tipohora, p_turnro, tipo_hora)
                        
            Total_Hs_Justificadas = Total_horas
            
            'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
            Fecha_Generacion = p_fecha
            'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
            
            StrSql = "INSERT INTO gti_horcumplido(horas, horcant,hordesde,horestado,horhasta,hormanual,horvalido,jusnro,jusnro2,ternro,thnro,turnro,empleg,horfecrep,horfecgen) VALUES (" & _
                     CHoras(Total_horas, 60) & "," & Total_horas & "," & ConvFecha(p_fecha) & ",' '," & ConvFecha(p_fecha) & ",0,-1," & objRsJustif!jusnro & "," & objRsJustif!jusnro & "," & Empleado.Ternro & "," & Aux_Tipohora & "," & p_turnro & "," & Empleado.Legajo & "," & ConvFecha(Fecha_Generacion) & "," & ConvFecha(p_fecha) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            If depurar Then
                Flog.writeline Espacios(Tabulador * 3) & "  ==> Insertó Justificación --> Tipo de Hora: " & Aux_Tipohora & "- Cantidad: " & Total_horas & " hs."
            End If
        End If
    
End Sub


Private Sub generar_dia_justificacion_old(p_ternro As Long, p_fecha As Date, p_turnro As Long, p_justif As Long)
Dim j_tipo As String
Dim objRsJustif As New ADODB.Recordset
Dim aux_canthorasjust As Single
Dim max_horas As Single
Dim horas_min As Single

    Aux_Tipohora = 0
    Call buscar_horas_turno(aux_canthorasjust, max_horas, horas_min)
    
    Total_horas = aux_canthorasjust
    
    StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & p_justif
    OpenRecordset StrSql, objRsJustif
    If objRsJustif.EOF Then
      ' MESSAGE "Hay problemas con la Justificaci¢n (o su tipo), avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)" VIEW_AS ALERT_BOX.
      Exit Sub
    End If
    Select Case objRsJustif!jussigla
        Case "NOV"
            j_tipo = "NOVEDAD"
            StrSql = "SELECT gti_novedad.*,gti_tiponovedad.thnro FROM gti_novedad INNER JOIN gti_tiponovedad ON gti_novedad.gtnovnro = gti_tiponovedad.gtnovnro WHERE (gnovotoa = " & p_ternro & ") AND " & _
                     " (gnovnro = " & objRsJustif!juscodext & ")"
            'FGZ - 19/05/2010 ------------ Control FT -------------
            StrSql = StrSql & " AND (gti_novedad.ft = 0 OR (gti_novedad.ft = -1 AND gti_novedad.ftap = -1))"
            'FGZ - 19/05/2010 ------------ Control FT -------------
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                ' MESSAGE "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)"
                Exit Sub
            End If
            Aux_Tipohora = objRs!thnro
            Call Politica(440)
            
        Case "LIC"
            j_tipo = "LICENCIA"
            StrSql = "SELECT emp_lic.*,tipdia.thnro,tipdia.tdnro FROM emp_lic INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro WHERE (empleado = " & p_ternro & ") AND " & _
                     " (emp_licnro = " & objRsJustif!juscodext & ")"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                ' MESSAGE "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)" VIEW_AS ALERT_BOX. */
                Exit Sub
            End If
            aux_TipoDia = objRs!tdnro
            Aux_Tipohora = objRs!thnro ' Tipo de Hora equivalente de la Licencia
            Call Politica(440)
            
        Case "CUR"
            j_tipo = "CURSO"
            Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para el Curso
            aux_canthorasjust = objRsJustif!juscanths
        Case "SUS"
            j_tipo = "SUSPENCION"
            Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para Suspenci¢n
            aux_canthorasjust = objRsJustif!juscanths
    End Select
    If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Tipo de la Justificación", j_tipo
    
    If Aux_Tipohora = 0 Then
        ' La Justificacion no se paga, no tiene tipo de hora asignado
        If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Falta Tipo de Hora en la justificación", Str(p_justif)
        Exit Sub
    Else
        'Call ValidarTipoDeHora(aux_Tipohora, p_turnro, tipo_hora)
        
        
        StrSql = "INSERT INTO gti_horcumplido(horas, horcant,hordesde,horestado,horhasta,hormanual,horvalido,jusnro,jusnro2,ternro,thnro,turnro,empleg,horfecrep) VALUES (" & _
                 Total_horas & "," & ConvFecha(p_fecha) & ",' '," & ConvFecha(p_fecha) & ",0,-1," & objRsJustif!jusnro & "," & objRsJustif!jusnro & "," & Empleado.Ternro & "," & Aux_Tipohora & "," & p_turnro & "," & Empleado.Legajo & "," & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub

Private Sub generar_justificacion_Parcial(p_turnro As Long, ByVal just_ent As Long, ByVal just_sal As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Nueva version del procedimiento para generar las justificaciones.
'              Motivo: version anterior procesa una sola justificacion.
' Autor      : FGZ
' Fecha      : 02/05/2005
' Ultima Mod.: FGZ - 17/11/2005 - Levanta solo las licencias en estado AUTORIZADAS
' Ultima Mod.: FGZ - 02/10/2006 - Si tenia solo justif en la salida no generaba ninguna
' ---------------------------------------------------------------------------------------------
Dim j_tipo As String
Dim HDesde As String
Dim HHasta As String
Dim aux_canthorasjust As Single
Dim frac_desde As Integer
Dim frac_hasta As Integer
Dim hora_desde As String
Dim hora_hasta As String
Dim Genera_Just_Ent As Boolean
Dim JFDesde As Date
Dim JFHasta As Date

Dim AuxE As String
Dim AuxS As String
Dim AuxFE As Date
Dim AuxFS As Date
Dim Cant_Int As Double
Dim Cant_HC1 As Double
Dim Cant_HC2 As Double
Dim Caso As String
Dim dia_anterior
dia_anterior = False
Dim objRsJustif As New ADODB.Recordset

    Aux_Tipohora = 0
    StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion"
    StrSql = StrSql & " INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro "
    StrSql = StrSql & " WHERE gti_justificacion.jusnro = " & just_ent
    'FGZ - 04/02/2014 -------------------------------------------
    StrSql = StrSql & " AND gti_justificacion.juseltipo <> 4 "
    'FGZ - 04/02/2014 -------------------------------------------

    
    StrSql = StrSql & " ORDER BY gti_justificacion.juseltipo "
    OpenRecordset StrSql, objRsJustif
'    If objRsJustif.EOF Then
'        Exit Sub
'    End If
    'FGZ - 02/10/2006
    If objRsJustif.EOF Then
        Genera_Just_Ent = False
    Else
        Genera_Just_Ent = True
    End If
    
    If Genera_Just_Ent Then
        'Primer justificacion
        Select Case objRsJustif!jussigla
            Case "NOV", "ALM"
                j_tipo = "NOVEDAD"
                StrSql = "SELECT gti_novedad.*,gti_tiponovedad.thnro FROM gti_novedad INNER JOIN gti_tiponovedad ON gti_novedad.gtnovnro = gti_tiponovedad.gtnovnro WHERE (gnovotoa = " & Empleado.Ternro & ") AND " & _
                         " (gnovnro = " & objRsJustif!juscodext & ")"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                StrSql = StrSql & " AND (gti_novedad.ft = 0 OR (gti_novedad.ft = -1 AND gti_novedad.ftap = -1))"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    Exit Sub
                End If
                Aux_Tipohora = objRs!thnro
                Call Politica(440)
                
            Case "LIC"
                j_tipo = "LICENCIA"
                StrSql = "SELECT emp_lic.*,tipdia.thnro FROM emp_lic "
                StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro "
                StrSql = StrSql & " WHERE (empleado = " & Empleado.Ternro & ") "
                StrSql = StrSql & " AND (emp_licnro = " & objRsJustif!juscodext & ")"
                'FGZ - 17/11/2005
                StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                'FGZ - 19/05/2010 ------------ Control FT -------------
                StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    Exit Sub
                End If
                Aux_Tipohora = objRs!thnro ' Tipo de Hora equivalente de la Licencia
                Call Politica(440)
            Case "CUR"
                j_tipo = "CURSO"
                Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para el Curso
                'aux_canthorasjust = objrsJustif!juscanths
            Case "SUS"
                j_tipo = "SUSPENCION"
                Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para Suspenci¢n
                'aux_canthorasjust = objrsJustif!juscanths
        End Select
        If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Tipo de la Justificación", j_tipo
        
        If objRsJustif!juscanths <> Null Then
            aux_canthorasjust = objRsJustif!juscanths
        Else
            aux_canthorasjust = 0
        End If
        
        If Aux_Tipohora = 0 Then
            ' La Justificacion no se paga, no tiene tipo de hora asignado
            If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Falta Tipo de Hora en la justificación", Str(just_ent)
            Exit Sub
        Else
'            HDesde = objRsJustif!jushoradesde
'            HHasta = objRsJustif!jushorahasta
'            If Not objFechasHoras.ValidarHora(HDesde) Then Exit Sub
'            If Not objFechasHoras.ValidarHora(HHasta) Then Exit Sub
'
'            hora_desde = objFechasHoras.FraccionaHs(objRsJustif!jushoradesde, frac_desde)
'            hora_hasta = objFechasHoras.FraccionaHs(objRsJustif!jushorahasta, frac_hasta)
'            'FGZ - 22/07/2014 -------
'            'objFechasHoras.RestaHs p_fecha, hora_desde, p_fecha, hora_hasta, Tdias, Thoras, Tmin
'            JFDesde = objRsJustif!jusdesde
'            JFHasta = objRsJustif!jushasta
'            objFechasHoras.RestaHs JFDesde, hora_desde, JFHasta, hora_hasta, Tdias, Thoras, Tmin
'            'FGZ - 22/07/2014 -------
'            aux_canthorasjust = (Tdias * 24) + (Thoras + (Tmin / 60))
'
'            Total_Hs_Justificadas = Total_Hs_Justificadas + aux_canthorasjust
'            Arr_Justificaciones(Indice_Justif).Ent = hora_desde
'            Arr_Justificaciones(Indice_Justif).Sal = hora_hasta
'            Arr_Justificaciones(Indice_Justif).Cantidad = aux_canthorasjust
'            Indice_Justif = Indice_Justif + 1
'
'            'Call ValidarTipoDeHora(aux_Tipohora, p_turnro, tipo_hora)
'
'            'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
'            'Fecha_Generacion = CalcularFechadeGeneracion(Nro_Subturno, p_fecha, fecpar1, fecpar2, Cambio_dia)
'            Fecha_Generacion = p_fecha
'            'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
'
'            'FGZ - 22/07/2014 ------------------------------
'            'StrSql = "INSERT INTO gti_horcumplido(horas, horcant,horhoradesde,horhorahasta,hordesde,horestado,horhasta,hormanual,horvalido,jusnro,ternro,thnro,turnro,empleg,horfecrep,horfecgen) VALUES (" & _
'            '         CHoras(aux_canthorasjust, 60) & "," & aux_canthorasjust & ",'" & HDesde & "','" & HHasta & "'," & ConvFecha(p_fecha) & ",' '," & ConvFecha(p_fecha) & ",0,-1," & objRsJustif!jusnro & "," & Empleado.Ternro & "," & Aux_Tipohora & "," & p_turnro & "," & Empleado.Legajo & "," & ConvFecha(Fecha_Generacion) & "," & ConvFecha(p_fecha) & ")"
'            StrSql = "INSERT INTO gti_horcumplido(horas, horcant,horhoradesde,horhorahasta,hordesde,horestado,horhasta,hormanual,horvalido,jusnro,ternro,thnro,turnro,empleg,horfecrep,horfecgen) VALUES (" & _
'                     CHoras(aux_canthorasjust, 60) & "," & aux_canthorasjust & ",'" & HDesde & "','" & HHasta & "'," & ConvFecha(JFDesde) & ",' '," & ConvFecha(JFHasta) & ",0,-1," & objRsJustif!jusnro & "," & Empleado.Ternro & "," & Aux_Tipohora & "," & p_turnro & "," & Empleado.Legajo & "," & ConvFecha(Fecha_Generacion) & "," & ConvFecha(p_fecha) & ")"
'            'FGZ - 22/07/2014 ------------------------------
'            objConn.Execute StrSql, , adExecuteNoRecords
'            If depurar Then
'                Flog.writeline Espacios(Tabulador * 3) & "  ==> Insertó Justificación --> Tipo de Hora: " & Aux_Tipohora & "- Cantidad: " & aux_canthorasjust & " hs."
'            End If
            
            'FGZ - 25/08/2014 -------------------------------------------------------
            'HDesde = objRsJustif!jushoradesde
            'HHasta = objRsJustif!jushorahasta
            'If Not objFechasHoras.ValidarHora(HDesde) Then Exit Sub
            'If Not objFechasHoras.ValidarHora(HHasta) Then Exit Sub
            'hora_desde = objFechasHoras.FraccionaHs(objRsJustif!jushoradesde, frac_desde)
            'hora_hasta = objFechasHoras.FraccionaHs(objRsJustif!jushorahasta, frac_hasta)
            'JFDesde = objRsJustif!jusdesde
            'JFHasta = objRsJustif!jushasta
            
            'Reviso la cantidad de horas que se intersectan entre el torico y la justificacion
            AuxE = E1
            AuxFE = FE1
            AuxS = S1
            AuxFS = FS1
            If Not EsNulo(FS2) And Not EsNulo(S2) Then
                AuxS = S2
                AuxFS = FS2
            End If
            If Not EsNulo(FS3) And Not EsNulo(S3) Then
                AuxS = S3
                AuxFS = FS3
            End If
            
            If Not EsNulo(AuxE) And Not EsNulo(AuxS) Then
                'Call CalcularInterseccionHoras3(AuxFE, AuxE, AuxFS, AuxS, objRsJustif!jusdesde, objRsJustif!jushoradesde, objRsJustif!jushasta, objRsJustif!jushorahasta, objRsJustif!juselMaxHoras, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
              If Not IsNull(objRsJustif!jushoradesde) And (objRsJustif!jushoradesde <> "") Then
                Call CalcularInterseccionHoras3(AuxFE, AuxE, AuxFS, AuxS, objRsJustif!jusdesde, objRsJustif!jushoradesde, objRsJustif!jushasta, objRsJustif!jushorahasta, 8, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
                 If Cant_Int = 0 And AuxFE <> AuxFS Then
                   Call CalcularInterseccionHoras3(DateAdd("d", -1, AuxFE), AuxE, AuxFE, AuxS, objRsJustif!jusdesde, objRsJustif!jushoradesde, objRsJustif!jushasta, objRsJustif!jushorahasta, 8, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
                   dia_anterior = True
                 End If
              Else 'MDF 12/05/2015 - Si vienen null hora desde y hasta, es todo el dia
                Call CalcularInterseccionHoras3(AuxFE, AuxE, AuxFS, AuxS, objRsJustif!jusdesde, "0000", objRsJustif!jushasta, "2400", 12, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
              End If
                If Cant_Int <> 0 Then
                    HDesde = hora_desde
                    HHasta = hora_hasta
                    objFechasHoras.RestaHs JFDesde, hora_desde, JFHasta, hora_hasta, Tdias, Thoras, Tmin
                    aux_canthorasjust = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                    Total_Hs_Justificadas = Total_Hs_Justificadas + aux_canthorasjust
                    Arr_Justificaciones(Indice_Justif).Ent = hora_desde
                    Arr_Justificaciones(Indice_Justif).Sal = hora_hasta
                    Arr_Justificaciones(Indice_Justif).Cantidad = aux_canthorasjust
                    Indice_Justif = Indice_Justif + 1
                    If Not dia_anterior Then
                      Fecha_Generacion = p_fecha
                    Else
                       Fecha_Generacion = DateAdd("d", -1, p_fecha)
                    End If
                    StrSql = "INSERT INTO gti_horcumplido(horas, horcant,horhoradesde,horhorahasta,hordesde,horestado,horhasta,hormanual,horvalido,jusnro,ternro,thnro,turnro,empleg,horfecrep,horfecgen) VALUES (" & _
                             CHoras(aux_canthorasjust, 60) & "," & aux_canthorasjust & ",'" & HDesde & "','" & HHasta & "'," & ConvFecha(JFDesde) & ",' '," & ConvFecha(JFHasta) & ",0,-1," & objRsJustif!jusnro & "," & Empleado.Ternro & "," & Aux_Tipohora & "," & p_turnro & "," & Empleado.Legajo & "," & ConvFecha(Fecha_Generacion) & "," & ConvFecha(p_fecha) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    lista_procesadas = lista_procesadas & "," & objRsJustif!jusnro '---MDF
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * 3) & "  ==> Insertó Justificación --> Tipo de Hora: " & Aux_Tipohora & "- Cantidad: " & aux_canthorasjust & " hs."
                    End If
                End If
            End If
            'FGZ - 25/08/2014 -------------------------------------------------------
        End If
    End If
    
    'Segunda justificacion
    If nro_jus_ent <> nro_jus_sal Then
        StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion"
        StrSql = StrSql & " INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro "
        '--------------------MDF 07/07/2015
        If lista_justificacionesparciales = "0" Then
         StrSql = StrSql & " WHERE gti_justificacion.jusnro = " & just_sal
        Else
          StrSql = StrSql & " WHERE gti_justificacion.jusnro in (" & lista_justificacionesparciales & ")"
        End If
        '--------------------MDF 07/07/2015
        StrSql = StrSql & " ORDER BY gti_justificacion.juseltipo "
        If objRsJustif.State = adStateOpen Then objRsJustif.Close
        OpenRecordset StrSql, objRsJustif
        If objRsJustif.EOF Then
          Exit Sub
        End If
        Do While Not objRsJustif.EOF
        Select Case objRsJustif!jussigla
            Case "NOV"
                j_tipo = "NOVEDAD"
                StrSql = "SELECT gti_novedad.*,gti_tiponovedad.thnro FROM gti_novedad INNER JOIN gti_tiponovedad ON gti_novedad.gtnovnro = gti_tiponovedad.gtnovnro WHERE (gnovotoa = " & Empleado.Ternro & ") AND " & _
                         " (gnovnro = " & objRsJustif!juscodext & ")"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                StrSql = StrSql & " AND (gti_novedad.ft = 0 OR (gti_novedad.ft = -1 AND gti_novedad.ftap = -1))"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    Exit Sub
                End If
                Aux_Tipohora = objRs!thnro
                Call Politica(440)
                
            Case "LIC"
                j_tipo = "LICENCIA"
                StrSql = "SELECT emp_lic.*,tipdia.thnro FROM emp_lic"
                StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro "
                StrSql = StrSql & " WHERE (empleado = " & Empleado.Ternro & ") "
                StrSql = StrSql & " AND (emp_licnro = " & objRsJustif!juscodext & ")"
                'FGZ - 17/11/2005
                StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                'FGZ - 19/05/2010 ------------ Control FT -------------
                StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    Exit Sub
                End If
                Aux_Tipohora = objRs!thnro ' Tipo de Hora equivalente de la Licencia
                Call Politica(440)
            Case "CUR"
                j_tipo = "CURSO"
                Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para el Curso
                'aux_canthorasjust = objrsJustif!juscanths
            Case "SUS"
                j_tipo = "SUSPENCION"
                Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para Suspenci¢n
                'aux_canthorasjust = objrsJustif!juscanths
            Case "ALM"
                j_tipo = "ALMUERZO"
                StrSql = "SELECT gti_novedad.*,gti_tiponovedad.thnro FROM gti_novedad INNER JOIN gti_tiponovedad ON gti_novedad.gtnovnro = gti_tiponovedad.gtnovnro WHERE (gnovotoa = " & Empleado.Ternro & ") AND " & _
                         " (gnovnro = " & objRsJustif!juscodext & ")"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                StrSql = StrSql & " AND (gti_novedad.ft = 0 OR (gti_novedad.ft = -1 AND gti_novedad.ftap = -1))"
                'FGZ - 19/05/2010 ------------ Control FT -------------
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    Exit Sub
                End If
                Aux_Tipohora = objRs!thnro
                'Aux_Tipohora = objRsJustif!thnro
        End Select
        If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Tipo de la Justificación", j_tipo
        
        If objRsJustif!juscanths <> Null Then
            aux_canthorasjust = objRsJustif!juscanths
        Else
            aux_canthorasjust = 0
        End If
        
        If Aux_Tipohora = 0 Then
            ' La Justificacion no se paga, no tiene tipo de hora asignado
            If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Falta Tipo de Hora en la justificación", Str(just_ent)
            Exit Sub
        Else
'            HDesde = objRsJustif!jushoradesde
'            HHasta = objRsJustif!jushorahasta
'            If Not objFechasHoras.ValidarHora(HDesde) Then Exit Sub
'            If Not objFechasHoras.ValidarHora(HHasta) Then Exit Sub
'
'            hora_desde = objFechasHoras.FraccionaHs(objRsJustif!jushoradesde, frac_desde)
'            hora_hasta = objFechasHoras.FraccionaHs(objRsJustif!jushorahasta, frac_hasta)
'            'FGZ - 01/08/2014 ------------------------------
'            'objFechasHoras.RestaHs p_fecha, hora_desde, p_fecha, hora_hasta, Tdias, Thoras, Tmin
'            JFDesde = objRsJustif!jusdesde
'            JFHasta = objRsJustif!jushasta
'            objFechasHoras.RestaHs JFDesde, hora_desde, JFHasta, hora_hasta, Tdias, Thoras, Tmin
'            'FGZ - 01/08/2014 ------------------------------
'
'            aux_canthorasjust = (Tdias * 24) + (Thoras + (Tmin / 60))
'
'            Total_Hs_Justificadas = Total_Hs_Justificadas + aux_canthorasjust
'            Arr_Justificaciones(Indice_Justif).Ent = hora_desde
'            Arr_Justificaciones(Indice_Justif).Sal = hora_hasta
'            Arr_Justificaciones(Indice_Justif).Cantidad = aux_canthorasjust
'            Indice_Justif = Indice_Justif + 1
'
'            'FGZ - 18/04/2005
'            'Ojo con la cantidad de horas que se genera porque se deberia valida que las horas desde hasta
'            ' caigan dentro de una franja del horario teorico, cosa que no sucede o no se valida cuando el turno es libre
'
'            'Call ValidarTipoDeHora(aux_Tipohora, p_turnro, tipo_hora)
'
'            'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
'            'Fecha_Generacion = CalcularFechadeGeneracion(Nro_Subturno, p_fecha, fecpar1, fecpar2, Cambio_dia)
'            Fecha_Generacion = p_fecha
'            'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
'
'            StrSql = "INSERT INTO gti_horcumplido(horas, horcant,horhoradesde,horhorahasta,hordesde,horestado,horhasta,hormanual,horvalido,jusnro2,ternro,thnro,turnro,empleg,horfecrep,horfecgen) "
'            StrSql = StrSql & " VALUES ("
'            StrSql = StrSql & CHoras(aux_canthorasjust, 60)
'            StrSql = StrSql & "," & aux_canthorasjust
'            StrSql = StrSql & ",'" & HDesde & "'"
'            StrSql = StrSql & ",'" & HHasta & "'"
'            'FGZ - 01/08/2014 ------------------------------
'            'StrSql = StrSql & "," & ConvFecha(p_fecha)
'            StrSql = StrSql & "," & ConvFecha(JFDesde)
'            'FGZ - 01/08/2014 ------------------------------
'            StrSql = StrSql & ",' '"
'            'FGZ - 01/08/2014 ------------------------------
'            'StrSql = StrSql & "," & ConvFecha(p_fecha)
'            StrSql = StrSql & "," & ConvFecha(JFHasta)
'            'FGZ - 01/08/2014 ------------------------------
'            StrSql = StrSql & ",0"
'            StrSql = StrSql & ",-1"
'            StrSql = StrSql & "," & objRsJustif!jusnro
'            StrSql = StrSql & "," & Empleado.Ternro
'            StrSql = StrSql & "," & Aux_Tipohora
'            StrSql = StrSql & "," & p_turnro
'            StrSql = StrSql & "," & Empleado.Legajo
'            StrSql = StrSql & "," & ConvFecha(Fecha_Generacion)
'            StrSql = StrSql & "," & ConvFecha(p_fecha)
'            StrSql = StrSql & ")"
'            objConn.Execute StrSql, , adExecuteNoRecords
'
'            If depurar Then
'                Flog.writeline Espacios(Tabulador * 3) & "  ==> Insertó Justificación --> Tipo de Hora: " & Aux_Tipohora & "- Cantidad: " & aux_canthorasjust & " hs."
'            End If
        
        
            'FGZ - 25/08/2014 -------------------------------------------------------
            'HDesde = objRsJustif!jushoradesde
            'HHasta = objRsJustif!jushorahasta
            'If Not objFechasHoras.ValidarHora(HDesde) Then Exit Sub
            'If Not objFechasHoras.ValidarHora(HHasta) Then Exit Sub
            'hora_desde = objFechasHoras.FraccionaHs(objRsJustif!jushoradesde, frac_desde)
            'hora_hasta = objFechasHoras.FraccionaHs(objRsJustif!jushorahasta, frac_hasta)
            'JFDesde = objRsJustif!jusdesde
            'JFHasta = objRsJustif!jushasta
            
            'Reviso la cantidad de horas que se intersectan entre el torico y la justificacion
            AuxE = E1
            AuxFE = FE1
            AuxS = S1
            AuxFS = FS1
            If Not EsNulo(FS2) And Not EsNulo(S2) Then
                AuxS = S2
                AuxFS = FS2
            End If
            If Not EsNulo(FS3) And Not EsNulo(S3) Then
                AuxS = S3
                AuxFS = FS3
            End If
            
            If Not EsNulo(AuxE) And Not EsNulo(AuxS) Then
                'Call CalcularInterseccionHoras3(AuxFE, AuxE, AuxFS, AuxS, objRsJustif!jusdesde, objRsJustif!jushoradesde, objRsJustif!jushasta, objRsJustif!jushorahasta, objRsJustif!juselMaxHoras, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
                If Not IsNull(objRsJustif!jushoradesde) Then
                 Call CalcularInterseccionHoras3(AuxFE, AuxE, AuxFS, AuxS, objRsJustif!jusdesde, objRsJustif!jushoradesde, objRsJustif!jushasta, objRsJustif!jushorahasta, 8, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
                Else  'MDF 12/05/2015 - Si vienen null hora desde y hasta, es todo el dia
                  Call CalcularInterseccionHoras3(AuxFE, AuxE, AuxFS, AuxS, objRsJustif!jusdesde, "0000", objRsJustif!jushasta, "2400", 12, Cant_Int, JFDesde, hora_desde, JFHasta, hora_hasta, Caso, Cant_HC1, Cant_HC2)
                End If
                If Cant_Int <> 0 Then
                    HDesde = hora_desde
                    HHasta = hora_hasta
                    objFechasHoras.RestaHs JFDesde, hora_desde, JFHasta, hora_hasta, Tdias, Thoras, Tmin
                    aux_canthorasjust = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                    Total_Hs_Justificadas = Total_Hs_Justificadas + aux_canthorasjust
                    Arr_Justificaciones(Indice_Justif).Ent = hora_desde
                    Arr_Justificaciones(Indice_Justif).Sal = hora_hasta
                    Arr_Justificaciones(Indice_Justif).Cantidad = aux_canthorasjust
                    Indice_Justif = Indice_Justif + 1
                            
                    Fecha_Generacion = p_fecha
                    
                    StrSql = "INSERT INTO gti_horcumplido(horas, horcant,horhoradesde,horhorahasta,hordesde,horestado,horhasta,hormanual,horvalido,jusnro2,ternro,thnro,turnro,empleg,horfecrep,horfecgen) "
                    StrSql = StrSql & " VALUES ("
                    StrSql = StrSql & CHoras(aux_canthorasjust, 60)
                    StrSql = StrSql & "," & aux_canthorasjust
                    StrSql = StrSql & ",'" & HDesde & "'"
                    StrSql = StrSql & ",'" & HHasta & "'"
                    StrSql = StrSql & "," & ConvFecha(JFDesde)
                    StrSql = StrSql & ",' '"
                    StrSql = StrSql & "," & ConvFecha(JFHasta)
                    StrSql = StrSql & ",0"
                    StrSql = StrSql & ",-1"
                    StrSql = StrSql & "," & objRsJustif!jusnro
                    StrSql = StrSql & "," & Empleado.Ternro
                    StrSql = StrSql & "," & Aux_Tipohora
                    StrSql = StrSql & "," & p_turnro
                    StrSql = StrSql & "," & Empleado.Legajo
                    StrSql = StrSql & "," & ConvFecha(Fecha_Generacion)
                    StrSql = StrSql & "," & ConvFecha(p_fecha)
                    StrSql = StrSql & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                     lista_procesadas = lista_procesadas & "," & objRsJustif!jusnro
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * 3) & "  ==> Insertó Justificación ---> Tipo de Hora: " & Aux_Tipohora & "- Cantidad: " & aux_canthorasjust & " hs."
                    End If
                End If
            End If
            'FGZ - 25/08/2014 -------------------------------------------------------
        End If
        objRsJustif.MoveNext
      Loop
    End If
If objRsJustif.State = adStateOpen Then objRsJustif.Close
Set objRsJustif = Nothing
End Sub



Private Sub generar_justificacion_Parcial_old(p_turnro As Long, just_ent As Long, just_sal As Long)
Dim j_tipo As String
Dim HDesde As String
Dim HHasta As String
Dim objRsJustif As New ADODB.Recordset
Dim aux_canthorasjust As Single
Dim frac_desde As Integer
Dim frac_hasta As Integer
Dim hora_desde As String
Dim hora_hasta As String

    Aux_Tipohora = 0
    StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & just_ent
    OpenRecordset StrSql, objRsJustif
    If objRsJustif.EOF Then
      ' MESSAGE "Hay problemas con la Justificaci¢n (o su tipo), avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)" VIEW_AS ALERT_BOX.
      Exit Sub
    End If
    Select Case objRsJustif!jussigla
        Case "NOV"
            j_tipo = "NOVEDAD"
            StrSql = "SELECT gti_novedad.*,gti_tiponovedad.thnro FROM gti_novedad INNER JOIN gti_tiponovedad ON gti_novedad.gtnovnro = gti_tiponovedad.gtnovnro WHERE (gnovotoa = " & Empleado.Ternro & ") AND " & _
                     " (gnovnro = " & objRsJustif!juscodext & ")"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                ' MESSAGE "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)"
                Exit Sub
            End If
            Aux_Tipohora = objRs!thnro
            Call Politica(440)
            
        Case "LIC"
            j_tipo = "LICENCIA"
            StrSql = "SELECT emp_lic.*,tipdia.thnro FROM emp_lic INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro WHERE (empleado = " & Empleado.Ternro & ") AND " & _
                     " (emp_licnro = " & objRsJustif!juscodext & ")"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                ' MESSAGE "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)" VIEW_AS ALERT_BOX. */
                Exit Sub
            End If
            Aux_Tipohora = objRs!thnro ' Tipo de Hora equivalente de la Licencia
            Call Politica(440)
        Case "CUR"
            j_tipo = "CURSO"
            Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para el Curso
            'aux_canthorasjust = objrsJustif!juscanths
        Case "SUS"
            j_tipo = "SUSPENCION"
            Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para Suspenci¢n
            'aux_canthorasjust = objrsJustif!juscanths
    End Select
    If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Tipo de la Justificación", j_tipo
    
    If objRsJustif!juscanths <> Null Then
        
        aux_canthorasjust = objRsJustif!juscanths
    Else
        aux_canthorasjust = 0
    End If
    
    If Aux_Tipohora = 0 Then
        ' La Justificacion no se paga, no tiene tipo de hora asignado
        If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Falta Tipo de Hora en la justificación", Str(just_ent)
        Exit Sub
    Else
        HDesde = objRsJustif!jushoradesde
        HHasta = objRsJustif!jushorahasta
        If Not objFechasHoras.ValidarHora(HDesde) Then Exit Sub
        If Not objFechasHoras.ValidarHora(HHasta) Then Exit Sub
         
        hora_desde = objFechasHoras.FraccionaHs(objRsJustif!jushoradesde, frac_desde)
        hora_hasta = objFechasHoras.FraccionaHs(objRsJustif!jushorahasta, frac_hasta)
        objFechasHoras.RestaHs p_fecha, hora_desde, p_fecha, hora_hasta, Tdias, Thoras, Tmin
        aux_canthorasjust = (Tdias * 24) + (Thoras + (Tmin / 60))
                
        'Call ValidarTipoDeHora(aux_Tipohora, p_turnro, tipo_hora)
        StrSql = "INSERT INTO gti_horcumplido(horas, horcant,horhoradesde,horhorahasta,hordesde,horestado,horhasta,hormanual,horvalido,jusnro,ternro,thnro,turnro,empleg,horfecrep) VALUES (" & _
                 aux_canthorasjust & ",'" & HDesde & "','" & HHasta & "'," & ConvFecha(p_fecha) & ",' '," & ConvFecha(p_fecha) & ",0,-1," & objRsJustif!jusnro & "," & Empleado.Ternro & "," & Aux_Tipohora & "," & p_turnro & "," & Empleado.Legajo & "," & ConvFecha(p_fecha) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
    
    If nro_jus_ent <> nro_jus_sal Then
    
        StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojus ON gti_justificacion.tjusnro = gti_tipojus.tjusnro WHERE gti_justificacion.jusnro = " & just_sal
        OpenRecordset StrSql, objRsJustif
        If objRsJustif.EOF Then
          ' MESSAGE "Hay problemas con la Justificaci¢n (o su tipo), avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)" VIEW_AS ALERT_BOX.
          Exit Sub
        End If
        Select Case objRs!jussigla
            Case "NOV"
                j_tipo = "NOVEDAD"
                StrSql = "SELECT gti_novedad.*,gti_tiponovedad.thnro FROM gti_novedad INNER JOIN gti_tiponovedad ON gti_novedad.gtnovnro = gti_tiponovedad.gtnovnro WHERE (gnovotoa = " & Empleado.Ternro & ") AND " & _
                         " (gnovnro = " & objRsJustif!juscodext & ")"
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    ' MESSAGE "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)"
                    Exit Sub
                End If
                Aux_Tipohora = objRs!thnro
                Call Politica(440)
            
            Case "LIC"
                j_tipo = "LICENCIA"
                StrSql = "SELECT emp_lic.*,tipdia.thnro FROM emp_lic INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro WHERE (empleado = " & Empleado.Ternro & ") AND " & _
                         " (emp_licnro = " & objRsJustif!juscodext & ")"
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    ' MESSAGE "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)" VIEW_AS ALERT_BOX. */
                    Exit Sub
                End If
                Aux_Tipohora = objRs!thnro ' Tipo de Hora equivalente de la Licencia
                Call Politica(440)
            Case "CUR"
                j_tipo = "CURSO"
                Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para el Curso
                'aux_canthorasjust = objrsJustif!juscanths
            Case "SUS"
                j_tipo = "SUSPENCION"
                Aux_Tipohora = objRsJustif!thnro  'Tipo de Hora default para Suspenci¢n
                'aux_canthorasjust = objrsJustif!juscanths
        End Select
        If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Tipo de la Justificación", j_tipo
        
        If objRsJustif!juscanths <> Null Then
            
            aux_canthorasjust = objRsJustif!juscanths
        Else
            aux_canthorasjust = 0
        End If
        
        If Aux_Tipohora = 0 Then
            ' La Justificacion no se paga, no tiene tipo de hora asignado
            If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Falta Tipo de Hora en la justificación", Str(just_ent)
            Exit Sub
        Else
            HDesde = objRsJustif!jushoradesde
            HHasta = objRsJustif!jushorahasta
            If Not objFechasHoras.ValidarHora(HDesde) Then Exit Sub
            If Not objFechasHoras.ValidarHora(HHasta) Then Exit Sub
             
            hora_desde = objFechasHoras.FraccionaHs(objRs!jushoradesde, frac_desde)
            hora_hasta = objFechasHoras.FraccionaHs(objRs!jushorahasta, frac_hasta)
            objFechasHoras.RestaHs p_fecha, hora_desde, p_fecha, hora_hasta, Tdias, Thoras, Tmin
            aux_canthorasjust = (Tdias * 24) + (Thoras + (Tmin / 60))
                    
            'Call ValidarTipoDeHora(aux_Tipohora, p_turnro, tipo_hora)
            StrSql = "INSERT INTO gti_horcumplido(horas, horcant,horhoradesde,horhorahasta,hordesde,horestado,horhasta,hormanual,horvalido,jusnro,ternro,thnro,turnro,empleg,horfecrep) VALUES (" & _
                     aux_canthorasjust & ",'" & HDesde & "','" & HHasta & "'," & ConvFecha(p_fecha) & ",' '," & ConvFecha(p_fecha) & ",0,-1," & objRsJustif!jusnro2 & "," & Empleado.Ternro & "," & Aux_Tipohora & "," & p_turnro & "," & Empleado.Legajo & "," & ConvFecha(p_fecha) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
        End If
    
    End If
                 
    
End Sub

Private Sub Ventana(Nro_Dia As Integer, Fecha As Date, Ternro As Long, venthoradesde As String, ventFdesde As Date, venthorahasta As String, ventFhasta As Date, P_Asignacion As Boolean)
Dim objRsJustif As New ADODB.Recordset
Dim i As Integer

    GeneraDia_WF_Turno Ternro, Fecha, Nro_Dia, P_Asignacion
    'Flog.writeline Now & " GenDia"
    
    
    StrSql = "SELECT * FROM  gti_justificacion WHERE (ternro = " & Ternro & " ) AND " & _
             " (juseltipo = 2) AND (jusdesde <= " & ConvFecha(p_fecha) & " AND " & _
             " jushasta >= " & ConvFecha(p_fecha) & ")"
             '2010-05-10 EGO - Cambio el actuar el "(jusdesde <= " & ConvFecha(p_fecha) & " AND jushasta >= " & ConvFecha(p_fecha) & ")""
    OpenRecordset StrSql, objRsJustif
    Do While Not objRsJustif.EOF
        
        StrSql = "SELECT Codigo FROM " & TTempWFDiaLaboral & " WHERE (('" & objRsJustif!jushoradesde & "' <= hora_entrada or " & _
                 ConvFecha(Fecha) & " > fecha_entrada) AND ('" & objRsJustif!jushorahasta & "' <= hora_salida or " & _
                 ConvFecha(Fecha) & " < fecha_salida)) AND ('" & objRsJustif!jushorahasta & "' >= hora_entrada Or " & _
                 ConvFecha(Fecha) & " > Fecha_entrada )ORDER BY Codigo DESC"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            StrSql = "UPDATE " & TTempWFDiaLaboral & " SET Hora_entrada = '" & objRsJustif!jushorahasta & "'," & _
                    "nrojustif = " & objRsJustif!jusnro & " WHERE Codigo = " & objRs!Codigo
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
                 
        StrSql = "SELECT Codigo FROM " & TTempWFDiaLaboral & " WHERE (('" & objRsJustif!jushoradesde & "' >= hora_entrada or " & ConvFecha(Fecha) & "> fecha_entrada)AND " & _
                 "('" & objRsJustif!jushorahasta & "'>= hora_salida or " & ConvFecha(Fecha) & " < fecha_salida)) AND " & _
                 "('" & objRsJustif!jushoradesde & "'<= hora_salida or " & ConvFecha(Fecha) & "< fecha_salida)"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            StrSql = "UPDATE " & TTempWFDiaLaboral & " SET Hora_Salida = '" & objRsJustif!jushoradesde & "'," & _
                    "nrojustif = " & objRsJustif!jusnro & " WHERE Codigo = " & objRs!Codigo
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        StrSql = "SELECT * FROM " & TTempWFDiaLaboral & " WHERE (('" & objRsJustif!jushoradesde & "' >= hora_entrada or " & ConvFecha(Fecha) & " > fecha_entrada) AND " & _
                 " ('" & objRsJustif!jushorahasta & "' <= hora_salida or " & ConvFecha(Fecha) & " < fecha_salida))"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            StrSql = "DELETE FROM " & TTempWFDiaLaboral & " WHERE Codigo = " & objRs!Codigo
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'InsertarWFDiaLaboral -1, objRs!Fecha_entrada, IIf(objRs!hora_entrada > objRsJustif!jushoradesde, DateAdd("d", 1, objRs!Fecha_entrada), objRs!Fecha_entrada), objRs!hora_entrada, objRsJustif!jushoradesde, -1, objRsJustif!jusnro
            'InsertarWFDiaLaboral -1, objRs!Fecha_Salida, objRs!Fecha_Salida, objRsJustif!jushorahasta, objRs!hora_salida, -1, objRsJustif!jusnro
            ' G. Bauer - 30/10/2007 - se modifico para que traigas bien la entrada y salida cuando utiliza ventanas.
            InsertarWFDiaLaboral 1, objRs!Fecha_entrada, IIf(objRs!hora_entrada > objRsJustif!jushoradesde, DateAdd("d", -1, objRs!Fecha_entrada), objRs!Fecha_entrada), objRs!hora_entrada, objRsJustif!jushoradesde, 1, objRsJustif!jusnro
            InsertarWFDiaLaboral 2, objRs!Fecha_Salida, objRs!Fecha_Salida, objRsJustif!jushorahasta, objRs!hora_salida, 2, objRsJustif!jusnro
            
        End If
        objRsJustif.MoveNext
    Loop
    
    StrSql = "SELECT * FROM " & TTempWFDiaLaboral & " ORDER BY Fecha_entrada,Hora_Entrada,Fecha_Salida,Hora_Salida "
    OpenRecordset StrSql, objRs
    i = 1
    Do While Not objRs.EOF
        InsertarWFDia i, objRs!hora_entrada, objRs!Fecha_entrada, True
        i = i + 1
        InsertarWFDia i, objRs!hora_salida, objRs!Fecha_Salida, False
        i = i + 1
        objRs.MoveNext
    Loop
    
    Call Generar_Embudo_Dinamico(ventFdesde, ventFhasta, venthoradesde, venthorahasta)
    'Flog.writeline Now & " GenEmbudo"
End Sub

Private Sub Ventana_Movil(Nro_Dia As Integer, Fecha As Date, Ternro As Long, venthoradesde As String, ventFdesde As Date, venthorahasta As String, ventFhasta As Date, P_Asignacion As Boolean)
Dim objRsJustif As New ADODB.Recordset
Dim i As Integer

Dim JFDesde As Date
Dim JFHasta As Date

Dim AuxE As String
Dim AuxS As String
Dim AuxFE As Date
Dim AuxFS As Date
Dim Cant_Int As Double
Dim Cant_HC1 As Double
Dim Cant_HC2 As Double
Dim Caso As String


    'GeneraDia_WF_Turno Ternro, Fecha, Nro_Dia, P_Asignacion
    GeneraDia_WF_Turno_Movil Ternro, Fecha, Nro_Dia, P_Asignacion
    
    StrSql = "SELECT * FROM  gti_justificacion WHERE (ternro = " & Ternro & " ) AND " & _
             " (juseltipo = 2) AND (jusdesde <= " & ConvFecha(p_fecha) & " AND " & _
             " jushasta >= " & ConvFecha(p_fecha) & ")"
             '2010-05-10 EGO - Cambio el actuar el "(jusdesde <= " & ConvFecha(p_fecha) & " AND jushasta >= " & ConvFecha(p_fecha) & ")""
    OpenRecordset StrSql, objRsJustif
    Do While Not objRsJustif.EOF
            
                'StrSql = "SELECT Codigo FROM " & TTempWFDiaLaboral & " WHERE (('" & objRsJustif!jushoradesde & "' <= hora_entrada or " & _
                '         ConvFecha(Fecha) & " > fecha_entrada) AND ('" & objRsJustif!jushorahasta & "' <= hora_salida or " & _
                '         ConvFecha(Fecha) & " < fecha_salida)) AND ('" & objRsJustif!jushorahasta & "' >= hora_entrada Or " & _
                '         ConvFecha(Fecha) & " > Fecha_entrada )ORDER BY Codigo DESC"
                
        
                'EAM - Busca si la justifijacion es al comienzo del turno
                StrSql = "SELECT Codigo FROM " & TTempWFDiaLaboral & " WHERE (('" & objRsJustif!jushoradesde & "' <= hora_entrada AND " & _
                         ConvFecha(objRsJustif!jusdesde) & " = fecha_entrada) AND ('" & objRsJustif!jushorahasta & "' < hora_salida or " & _
                         ConvFecha(objRsJustif!jusdesde) & " <= fecha_salida)) AND ('" & objRsJustif!jushorahasta & "' >= hora_entrada Or " & _
                         ConvFecha(objRsJustif!jusdesde) & " >= Fecha_entrada )ORDER BY Codigo DESC"
                         
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    StrSql = "UPDATE " & TTempWFDiaLaboral & " SET Hora_entrada = '" & objRsJustif!jushorahasta & "'," & _
                            "nrojustif = " & objRsJustif!jusnro & " WHERE Codigo = " & objRs!Codigo
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                'EAM - Busca si la justifijacion es al final del turno
                StrSql = "SELECT Codigo FROM " & TTempWFDiaLaboral & " WHERE (('" & objRsJustif!jushoradesde & "' >= hora_entrada or " & ConvFecha(Fecha) & "> fecha_entrada)AND " & _
                         "('" & objRsJustif!jushorahasta & "'>= hora_salida or " & ConvFecha(Fecha) & " < fecha_salida)) AND " & _
                         "('" & objRsJustif!jushoradesde & "'<= hora_salida or " & ConvFecha(Fecha) & "< fecha_salida)"
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    StrSql = "UPDATE " & TTempWFDiaLaboral & " SET Hora_Salida = '" & objRsJustif!jushoradesde & "'," & _
                            "nrojustif = " & objRsJustif!jusnro & " WHERE Codigo = " & objRs!Codigo
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                StrSql = "SELECT * FROM " & TTempWFDiaLaboral & " WHERE (('" & objRsJustif!jushoradesde & "' >= hora_entrada or " & ConvFecha(Fecha) & " > fecha_entrada) AND " & _
                         " ('" & objRsJustif!jushorahasta & "' <= hora_salida or " & ConvFecha(Fecha) & " < fecha_salida))"
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    StrSql = "DELETE FROM " & TTempWFDiaLaboral & " WHERE Codigo = " & objRs!Codigo
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'InsertarWFDiaLaboral -1, objRs!Fecha_entrada, IIf(objRs!hora_entrada > objRsJustif!jushoradesde, DateAdd("d", 1, objRs!Fecha_entrada), objRs!Fecha_entrada), objRs!hora_entrada, objRsJustif!jushoradesde, -1, objRsJustif!jusnro
                    'InsertarWFDiaLaboral -1, objRs!Fecha_Salida, objRs!Fecha_Salida, objRsJustif!jushorahasta, objRs!hora_salida, -1, objRsJustif!jusnro
                    ' G. Bauer - 30/10/2007 - se modifico para que traigas bien la entrada y salida cuando utiliza ventanas.
                    InsertarWFDiaLaboral 1, objRs!Fecha_entrada, IIf(objRs!hora_entrada > objRsJustif!jushoradesde, DateAdd("d", -1, objRs!Fecha_entrada), objRs!Fecha_entrada), objRs!hora_entrada, objRsJustif!jushoradesde, 1, objRsJustif!jusnro
                    InsertarWFDiaLaboral 2, objRs!Fecha_Salida, objRs!Fecha_Salida, objRsJustif!jushorahasta, objRs!hora_salida, 2, objRsJustif!jusnro
                    ''FGZ - 29/08/2014 --------------------
                    'If Not EsNulo(objRsJustif!jushorahasta) And Not EsNulo(objRs!hora_salida) Then
                    '    InsertarWFDiaLaboral 2, objRs!Fecha_Salida, objRs!Fecha_Salida, objRsJustif!jushorahasta, objRs!hora_salida, 2, objRsJustif!jusnro
                    'End If
                    'FGZ - 29/08/2014 --------------------
                    
                End If
        '    End If
        'End If
        'FGZ - 29/08/2014 ---------------------
        
        objRsJustif.MoveNext
    Loop
    
    StrSql = "SELECT * FROM " & TTempWFDiaLaboral & " ORDER BY Fecha_entrada,Hora_Entrada,Fecha_Salida,Hora_Salida "
    OpenRecordset StrSql, objRs
    i = 1
    Do While Not objRs.EOF
        InsertarWFDia i, objRs!hora_entrada, objRs!Fecha_entrada, True
        i = i + 1
        InsertarWFDia i, objRs!hora_salida, objRs!Fecha_Salida, False
        i = i + 1
        objRs.MoveNext
    Loop
    
    Call Generar_Embudo_Dinamico(ventFdesde, ventFhasta, venthoradesde, venthorahasta)
End Sub



Private Sub GeneraDia_WF_Turno(Ternro As Long, Fecha As Date, Nro_Dia As Integer, P_Asignacion As Boolean)
Dim fecha_aux As Date

    fecha_aux = Fecha
    If P_Asignacion Then
        StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND " & _
                 " (gttempdesde <= " & ConvFecha(Fecha) & " ) AND " & _
                 " (" & ConvFecha(Fecha) & " <= gttemphasta)"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            
            If objRs!ttemphdesde1 = "0000" And objRs!ttemphhasta1 = "0000" Then Exit Sub
            If objRs!ttemphdesde1 <> "" Then
               If (objRs!diasiguiente = 1) Or (objRs!ttemphdesde1 > objRs!ttemphhasta1) Then fecha_aux = fecha_aux + 1
               InsertarWFDiaLaboral 1, Fecha, fecha_aux, objRs!ttemphdesde1, objRs!ttemphhasta1, 1
            End If
            
            If objRs!ttemphdesde2 = "0000" And objRs!ttemphhasta2 = "0000" Then Exit Sub
            If objRs!ttemphdesde2 <> "" Then
               If (objRs!diasiguiente = 2) Or (objRs!ttemphdesde2 > objRs!ttemphhasta2) Then fecha_aux = fecha_aux + 1
               InsertarWFDiaLaboral 2, Fecha, fecha_aux, objRs!ttemphdesde2, objRs!ttemphhasta2, 2
            End If
       
            If objRs!ttemphdesde3 = "0000" And objRs!ttemphhasta3 = "0000" Then Exit Sub
            If objRs!ttemphdesde3 <> "" Then
               If (objRs!diasiguiente = 3) Or (objRs!ttemphdesde3 > objRs!ttemphhasta3) Then fecha_aux = fecha_aux + 1
               InsertarWFDiaLaboral 3, Fecha, fecha_aux, objRs!ttemphdesde3, objRs!ttemphhasta3, 3
            End If
        End If
    Else
        If Not Horario_Movil And Not Horario_Flexible_Rotativo Then
            StrSql = "SELECT * FROM  gti_dias where dianro = " & Nro_Dia
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                If objRs!diahoradesde1 = "0000" And objRs!diahorahasta1 = "0000" Then Exit Sub
                If objRs!diahoradesde1 <> "" Then
                   If (objRs!diasiguiente = 1) Or (objRs!diahoradesde1 > objRs!diahorahasta1) Then fecha_aux = fecha_aux + 1
                   InsertarWFDiaLaboral 1, Fecha, fecha_aux, objRs!diahoradesde1, objRs!diahorahasta1, 1
                End If
                
                If objRs!diahoradesde2 = "0000" And objRs!diahorahasta2 = "0000" Then Exit Sub
                If objRs!diahoradesde2 <> "" Then
                   If (objRs!diasiguiente = 2) Or (objRs!diahoradesde2 > objRs!diahorahasta2) Then fecha_aux = fecha_aux + 1
                   InsertarWFDiaLaboral 2, Fecha, fecha_aux, objRs!diahoradesde2, objRs!diahorahasta2, 2
                End If
                
                If objRs!diahoradesde3 = "0000" And objRs!diahorahasta3 = "0000" Then Exit Sub
                If objRs!diahoradesde3 <> "" Then
                   If (objRs!diasiguiente = 3) Or (objRs!diahoradesde3 > objRs!diahorahasta3) Then fecha_aux = fecha_aux + 1
                   InsertarWFDiaLaboral 3, Fecha, fecha_aux, objRs!diahoradesde3, objRs!diahorahasta3, 3
                End If
            End If
        Else
            'Es horario movil
            If Pasa_de_Dia Then
                fecha_aux = FE1 + 1
            Else
                fecha_aux = FE1
            End If
            
            If E1 = "0000" And S1 = "0000" Then Exit Sub
            If E1 <> "" Then
               If (Pasa_de_Dia) Or (E1 > S1) Then fecha_aux = fecha_aux + 1
               InsertarWFDiaLaboral 1, Fecha, fecha_aux, E1, S1, 1
            End If
            'InsertarWFDiaLaboral 1, Fecha, fecha_aux, E1, S1, 1
        End If
        
    End If
End Sub


Private Sub GeneraDia_WF_Turno_Movil(Ternro As Long, Fecha As Date, Nro_Dia As Integer, P_Asignacion As Boolean)
Dim fecha_aux As Date

    fecha_aux = Fecha
    If Not Horario_Flexible_sinParte Then
        If P_Asignacion Then
            StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro = " & Ternro & ") AND " & _
                     " (gttempdesde <= " & ConvFecha(Fecha) & " ) AND " & _
                     " (" & ConvFecha(Fecha) & " <= gttemphasta)"
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                
                If objRs!ttemphdesde1 = "0000" And objRs!ttemphhasta1 = "0000" Then Exit Sub
                If objRs!ttemphdesde1 <> "" Then
                
                    If TurnoNocturno_HaciaAtras Then
                        'EAM- Monresa- Turnos nocturnos hacia atras
                        InsertarWFDiaLaboral 1, (Fecha - 1), fecha_aux, objRs!ttemphdesde1, objRs!ttemphhasta1, 1
                    Else
                        If (objRs!diasiguiente = 1) Or (objRs!ttemphdesde1 > objRs!ttemphhasta1) Then fecha_aux = fecha_aux + 1
                        InsertarWFDiaLaboral 1, Fecha, fecha_aux, objRs!ttemphdesde1, objRs!ttemphhasta1, 1
                    End If
                End If
                
                If objRs!ttemphdesde2 = "0000" And objRs!ttemphhasta2 = "0000" Then Exit Sub
                If objRs!ttemphdesde2 <> "" Then
                   If (objRs!diasiguiente = 2) Or (objRs!ttemphdesde2 > objRs!ttemphhasta2) Then fecha_aux = fecha_aux + 1
                   InsertarWFDiaLaboral 2, Fecha, fecha_aux, objRs!ttemphdesde2, objRs!ttemphhasta2, 2
                End If
           
                If objRs!ttemphdesde3 = "0000" And objRs!ttemphhasta3 = "0000" Then Exit Sub
                If objRs!ttemphdesde3 <> "" Then
                   If (objRs!diasiguiente = 3) Or (objRs!ttemphdesde3 > objRs!ttemphhasta3) Then fecha_aux = fecha_aux + 1
                   InsertarWFDiaLaboral 3, Fecha, fecha_aux, objRs!ttemphdesde3, objRs!ttemphhasta3, 3
                End If
            End If
        Else
            If Not Horario_Movil And Not Horario_Flexible_Rotativo Then
                StrSql = "SELECT * FROM  gti_dias where dianro = " & Nro_Dia
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    If objRs!diahoradesde1 = "0000" And objRs!diahorahasta1 = "0000" Then Exit Sub
                    If objRs!diahoradesde1 <> "" Then
                       If (objRs!diasiguiente = 1) Or (objRs!diahoradesde1 > objRs!diahorahasta1) Then fecha_aux = fecha_aux + 1
                       InsertarWFDiaLaboral 1, Fecha, fecha_aux, objRs!diahoradesde1, objRs!diahorahasta1, 1
                    End If
                    
                    If objRs!diahoradesde2 = "0000" And objRs!diahorahasta2 = "0000" Then Exit Sub
                    If objRs!diahoradesde2 <> "" Then
                       If (objRs!diasiguiente = 2) Or (objRs!diahoradesde2 > objRs!diahorahasta2) Then fecha_aux = fecha_aux + 1
                       InsertarWFDiaLaboral 2, Fecha, fecha_aux, objRs!diahoradesde2, objRs!diahorahasta2, 2
                    End If
                    
                    If objRs!diahoradesde3 = "0000" And objRs!diahorahasta3 = "0000" Then Exit Sub
                    If objRs!diahoradesde3 <> "" Then
                       If (objRs!diasiguiente = 3) Or (objRs!diahoradesde3 > objRs!diahorahasta3) Then fecha_aux = fecha_aux + 1
                       InsertarWFDiaLaboral 3, Fecha, fecha_aux, objRs!diahoradesde3, objRs!diahorahasta3, 3
                    End If
                End If
            Else
    '            'Es horario movil
    '            If Pasa_de_Dia Then
    '                fecha_aux = FE1 + 1
    '            Else
    '                fecha_aux = FE1
    '            End If
    '
    '            If E1 = "0000" And S1 = "0000" Then Exit Sub
    '            If E1 <> "" Then
    '               If (Pasa_de_Dia) Or (E1 > S1) Then fecha_aux = fecha_aux + 1
    '               InsertarWFDiaLaboral 1, Fecha, fecha_aux, E1, S1, 1
    '            End If
    '            'InsertarWFDiaLaboral 1, Fecha, fecha_aux, E1, S1, 1
                
                
                If E1 = "0000" And S1 = "0000" Then Exit Sub
                If E1 <> "" Then
                   InsertarWFDiaLaboral 1, FE1, FS1, E1, S1, 1
                End If
            End If
        End If
    Else
        If E1 = "0000" And S1 = "0000" Then Exit Sub
        If E1 <> "" Then
           InsertarWFDiaLaboral 1, FE1, FS1, E1, S1, 1
        End If
    End If
    
End Sub



Private Sub Generar_Embudo_Dinamico(fecha_desde As Date, fecha_hasta As Date, hora_desde As String, hora_hasta As String)
Dim aux_fecdesde As Date
Dim aux_fechasta As Date
Dim Aux_HoraDesde As String
Dim Aux_HoraHasta As String
Dim pto As String
Dim pto_ant As String
Dim fecha_vent As Date
Dim fecha_vent_hasta As Date
Dim i As Long

    i = 0
    
    fecha_vent = fecha_desde
    pto_ant = hora_desde
    Aux_HoraDesde = hora_desde
    aux_fecdesde = fecha_desde
    Aux_HoraHasta = hora_hasta
    aux_fechasta = fecha_hasta

    OpenRecordset "SELECT * FROM " & TTempWFDiaLaboral & " ORDER BY Codigo", objRs
    Do While Not objRs.EOF
        i = i + 1
        fecha_vent_hasta = fecha_vent
        pto = CalcularPto(objRs!hora_entrada, objRs!hora_salida, objRs!hora_entrada > objRs!hora_salida)
        If pto < pto_ant Then fecha_vent_hasta = DateAdd("d", 1, fecha_vent_hasta)
    
        InsertarWFEmbudo i, fecha_vent, pto_ant, fecha_vent_hasta, pto
        Aux_HoraHasta = objRs!hora_salida
                
        objRs.MoveNext
        
        If Not objRs.EOF Then
            
            i = i + 1
            fecha_vent = fecha_vent_hasta
            pto_ant = pto
            pto = CalcularPto(Aux_HoraHasta, objRs!hora_entrada, Aux_HoraHasta > objRs!hora_entrada)
            If pto < pto_ant Then fecha_vent_hasta = DateAdd("d", 1, fecha_vent_hasta)
        
            InsertarWFEmbudo i, fecha_vent, pto_ant, fecha_vent_hasta, pto
            pto_ant = pto
        Else
            i = i + 1
            fecha_vent = fecha_vent_hasta
            pto_ant = pto
'            pto = CalcularPto(aux_horadesde, hora_hasta, aux_horadesde > hora_hasta)
'            If pto < pto_ant Then DateAdd "d", 1, fecha_vent_hasta
        
            InsertarWFEmbudo i, fecha_vent, pto_ant, fecha_hasta, hora_hasta
            pto_ant = pto
        End If
    Loop
End Sub

Private Function CalcularPto(E1 As String, E2 As String, PD As Boolean) As String
Dim minutosE1 As Integer
Dim minutosE2 As Integer
Dim Minutos As Integer
Dim ok As Boolean
Dim mitad As String

    minutosE1 = (Int(Mid(E1, 1, 2)) * 60 + Int(Mid(E1, 3, 2)))
    If PD Then
        minutosE2 = Int(Mid(E2, 1, 2)) * 60 + Int(Mid(E2, 3, 2)) + 1440 ' cant min. del dia
    Else
        minutosE2 = Int(Mid(E2, 1, 2)) * 60 + Int(Mid(E2, 3, 2))
    End If
    Minutos = ((minutosE2 - minutosE1) / 2) + minutosE1
    If Minutos > 1440 Then Minutos = Minutos - 1440
    mitad = (Format(Int(Minutos / 60), "00")) & (Format(Minutos - (Int(Minutos / 60) * 60), "00"))
'    ok = ValidarHora(mitad)
    CalcularPto = mitad

End Function

Private Function Existe_Registracion(Ternro As Long, fecha_desde As Date, hora_desde As String, fecha_hasta As Date, hora_hasta As String) As Boolean

Dim result As Boolean
Dim Continuar As Boolean
Dim salir As Boolean

    salir = False
    Existe_Registracion = False
    result = False
    StrSql = "SELECT regfecha,reghora FROM gti_registracion WHERE (regestado = 'I')"
    StrSql = StrSql & " AND (ternro = " & Ternro & ") AND ( regfecha >= " & ConvFecha(fecha_desde) & ")"
    StrSql = StrSql & " AND (regfecha <=" & ConvFecha(fecha_hasta) & ")"
    StrSql = StrSql & " AND ( regllamada = 0 OR regllamada is null )"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " AND (gti_registracion.ft = 0 OR (gti_registracion.ft = -1 AND gti_registracion.ftap = -1))"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " ORDER BY regfecha ASC, reghora ASC"
    OpenRecordset StrSql, objRs
    Do While Not salir And Not objRs.EOF
        Continuar = True
        If objRs!regfecha = fecha_desde Then
            If objRs!reghora < hora_desde Then
                objRs.MoveNext
                Continuar = False
            End If
        End If
        
        If Continuar Then
            If (objRs!regfecha = fecha_hasta) And (Continuar) Then
                If objRs!reghora > hora_hasta Then
                    Continuar = False
                    salir = True
                End If
            End If
        End If
        
        If (Continuar) Then result = True
        
        If Not objRs.EOF Then objRs.MoveNext
                         
    Loop
    Existe_Registracion = result
End Function

Private Function Existe_Registracion_LLamada(ByVal Ternro As Long, fecha_desde As Date, hora_desde As String, fecha_hasta As Date, hora_hasta As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que busca si existen registraciones con marca de llamada.
' Autor      : FGZ
' Fecha      : 23/05/2008
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
    Existe_Registracion_LLamada = False
    
    StrSql = "SELECT regfecha,reghora FROM gti_registracion WHERE (regestado = 'I')"
    StrSql = StrSql & " AND (ternro = " & Ternro & ") AND ( regfecha >= " & ConvFecha(fecha_desde) & ")"
    StrSql = StrSql & " AND (regfecha <=" & ConvFecha(fecha_hasta) & ")"
    StrSql = StrSql & " AND ( regllamada = -1)"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    StrSql = StrSql & " AND (gti_registracion.ft = 0 OR (gti_registracion.ft = -1 AND gti_registracion.ftap = -1))"
    StrSql = StrSql & " ORDER BY ternro ASC, regfecha ASC, reghora ASC"
    'FGZ - 19/05/2010 ------------ Control FT -------------
    OpenRecordset StrSql, objRs
    Existe_Registracion_LLamada = Not objRs.EOF
End Function


Private Sub Tolerancias()
    ' Politicas Gral de Tol. tarde
    Call Politica(130)
    'Flog.writeline Now & " Pol130"
    ' Politicas Gral de Tol. Temprano
    Call Politica(140)
    'Flog.writeline Now & " Pol140"
    ' Politica de tolerancia de Dto.
    Call Politica(150)
    'Flog.writeline Now & " Pol150"
    ' Politica para ver si acumula en HC las llegas Tarde
    Call Politica(160)
    'Flog.writeline Now & " Pol160"
    ' Politica para ver si acumula en HC salidas Temprano
    Call Politica(180)
    'Flog.writeline Now & " Pol180"
    ' Politica para ver si acumula en HC las Horas Ausente
    Call Politica(170)
    'Flog.writeline Now & " Pol170"
End Sub

Public Sub ValidarTipoDeHora(ByVal TipoHoraABuscar As Integer, ByVal Aux_Turno As Integer, tipo_horaValido As Integer)
' este procedimiento valida el tipo de Hora a Buscar, si no existe ese tipo ==>
' le setea el tipo de Hora Error
' FGZ - 24/03/2003

Dim rsBusqueda As New ADODB.Recordset
Dim rsError As New ADODB.Recordset

StrSql = "SELECT * FROM tiphora WHERE thnro = " & TipoHoraABuscar
OpenRecordset StrSql, rsBusqueda

If rsBusqueda.EOF Then
   ' no lo encontró, seteo el tipo de hora ERROR
    StrSql = "SELECT thnro FROM Gti_Config_tur_hor where conhornro = 40 AND turnro = " & Aux_Turno
    OpenRecordset StrSql, rsError
    If Not rsError.EOF Then
        ' tipo de Hora Error
        tipo_horaValido = rsError!thnro
    Else
        ' no esta configurado el tipo de Hora Error y debería estarlo
    End If
Else
    ' el tipo de hora existe, lo cargo
    tipo_horaValido = TipoHoraABuscar
End If
End Sub







Private Sub Armar_Embudo()

Dim tothoras   As Single
Dim Tdias      As Integer
Dim Thoras     As Integer
Dim Tmin       As Integer
Dim CantR      As Integer
Dim Reg_Oblig  As Integer
Dim regpares As Boolean
Dim haydos As Boolean
Dim siguiente As Integer
Dim cant1 As Integer
Dim cant2 As Integer
Dim cant3 As Integer
Dim cant4 As Integer
Dim cant5 As Integer
Dim cant6 As Integer
Dim paga_almuerzo As Boolean
Dim emb_prox As Boolean
Dim emb_ante As Boolean
Dim objrsReg As New ADODB.Recordset
Dim ok As Boolean

    acumula = False
    acumula_dto = False
    acumula_temp = False
    regpares = True
    haydos = False
    paga_almuerzo = False
    
    ' Politicas Gral de Tol. tarde
    Call Politica(130)
    ' Politicas Gral de Tol. Temprano
    Call Politica(140)
    ' Politica de tolerancia de Dto.
    Call Politica(150)
        ' Politica para ver si acumula en HC las llegas Tarde
    Call Politica(160)
    ' Politica para ver si acumula en HC salidas Temprano
    Call Politica(180)
    ' Politica para ver si acumula en HC las Horas Ausente
    Call Politica(170)
 
    'Recorro las registraciones para saber si son pares o impares
    StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE evenro = 2 ORDER BY fecha,hora"
    OpenRecordset StrSql, objRs
    With objRs
        Do While Not .EOF
            If ((!Fecha > fv1) Or (!Fecha = fv1 And !Hora >= v1)) And ((!Fecha = fv2 And !Hora <= v2) Or (!Fecha < fv2)) Then cant1 = cant1 + 1
            If ((!Fecha > fv2) Or (!Fecha = fv2 And !Hora > v2)) And ((!Fecha = fv3 And !Hora <= v3) Or (!Fecha < fv3)) Then cant2 = cant2 + 1
            If ((!Fecha > fv3) Or (!Fecha = fv3 And !Hora > v3)) And ((!Fecha = fv4 And !Hora <= v4) Or (!Fecha < fv4)) Then cant3 = cant3 + 1
            If ((!Fecha > fv4) Or (!Fecha = fv4 And !Hora > v4)) And ((!Fecha = fv5 And !Hora <= v5) Or (!Fecha < fv5)) Then cant4 = cant4 + 1
            If ((!Fecha > fv5) Or (!Fecha = fv5 And !Hora > v5)) And ((!Fecha = fv6 And !Hora <= v6) Or (!Fecha < fv6)) Then cant5 = cant5 + 1
            If ((!Fecha > fv6) Or (!Fecha = fv6 And !Hora > v6)) And ((!Fecha = fv7 And !Hora <= v7) Or (!Fecha < fv7)) Then cant6 = cant6 + 1
            CantR = CantR + 1
            .MoveNext
        Loop
    End With
    regpares = (CantR Mod 2) = 0
    
    ' PREGUNTA !!!
    ' Política de Falta de Registraciones Obligatorias - Genera Ausente/Presente
    ' Ver bien el lugar donde ejecutar esta política
    ' Ver que pasa con políticas de 0 Reg. y Reg. impares (cantr <> 1 and regpares = false)
    
    ' Se cuentan las E/S teóricas del día
    Reg_Oblig = 0
    StrSql = "SELECT Count(Codigo) as cantidad FROM " & TTempWFDia
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then If Not IsNull(objRs!Cantidad) Then Reg_Oblig = objRs!Cantidad
  
    If (Tiene_Justif) And (Not No_Trabaja_just) Then Reg_Oblig = Reg_Oblig - 2
    
    'Se fija si hay menos reg. que en la definición del turno
    If (CantR < Reg_Oblig) Then Call Politica(110)

    ' Por lo menos debe haber 3 ptos para una E y S
    If v1 <> "" And v2 <> "" And v3 <> "" Then
        'Busco registracion en el primer embudo
        StrSql = " SELECT * FROM " & TTempWFTurno & " WHERE evenro = 2 AND ((fecha > " & ConvFecha(fv1) & ") OR (fecha = " & ConvFecha(fv1) & " AND  hora >= '" & v1 & "'))" & _
                 " AND ((fecha = " & ConvFecha(fv2) & " AND hora <= '" & v2 & "') OR (fecha < " & ConvFecha(fv2) & "))"
        OpenRecordset StrSql, objRs, adLockOptimistic
        If Not objRs.EOF Then
            'Calculo tolerancia tarde */
            objFechasHoras.SumoHoras p_fecha, E1, tol, Fecha_Tol, Hora_Tol
            ok = objFechasHoras.ValidarHora(Hora_Tol)
            If objRs!Fecha = Fecha_Tol And objRs!Hora > Hora_Tol Then
                'Genero anormalidad de llegada tarde */
                'ASSIGN wf-turno.anornro = 5
                 objRs.Update objRs.Fields("anornro"), 5
                 objRs.UpdateBatch
                'Si acumula genero-hc-llegada-tarde
                 If acumula Then
                    Hora_Tol = E1
                    ok = objFechasHoras.ValidarHora(Hora_Tol)
                    objFechasHoras.RestaHs Fecha_Tol, Hora_Tol, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                    tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                    Generar_Llegada_Tarde objRs!Fecha, tothoras, objRs!Hora, Hora_Tol
                 End If
                 objFechasHoras.SumoHoras FE1, E1, toldto, fecha_toldto, hora_toldto
                 ok = objFechasHoras.ValidarHora(hora_toldto)
                 If (objRs!Fecha = fecha_toldto And objRs!Hora > hora_toldto) And (acumula_dto = True) Then
                     objFechasHoras.RestaHs fecha_toldto, hora_toldto, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                     tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                     'Descontar_Horas objRs!Fecha, tothoras
                     'FGZ - 01/12/2005
                     Call Descontar_Horas(objRs!Fecha, tothoras, 1)
                 End If
            Else
                'Entra en tolerancia. La fichada debe ser la del turno
                If (objRs!Fecha = Fecha_Tol) And (objRs!Hora > E1) Then
                    'Llega tarde
                    objRs.Update objRs.Fields("Fecha"), ConvFecha(FE1)
                    objRs.Update objRs.Fields("Hora"), E1
                    objRs.UpdateBatch
                End If
            End If
            'Asigno el registro como la primer componente del par1
            objRs.Update objRs.Fields("par"), 1
            objRs.UpdateBatch
            
            If cant1 >= 2 And cant2 = 0 Then
                'Busco la ult registracion en el 1er Embudo para pasarla al 2do Embudo */
                StrSql = " SELECT * FROM " & TTempWFTurno & " WHERE (wf-turno.evenro = 2) AND ((fecha > " & ConvFecha(fv1) & ") OR (fecha = '" & fv1 & "' AND  hora > '" & v1 & "')) " & _
                         " AND ((fecha = " & ConvFecha(fv2) & " AND hora <= '" & v2 & "') OR (fecha < " & ConvFecha(fv2) & ")) AND par = 0 "
                OpenRecordset StrSql, objRs, adLockOptimistic
                objRs.MoveLast
                If objRs.EOF Then Exit Sub
                    'Calculo tolerancia temprano
                    objFechasHoras.RestaXHoras FS1, S1, toltemp, Fecha_Tol, Hora_Tol
                    ok = objFechasHoras.ValidarHora(Hora_Tol)
                    If objRs!Fecha = Fecha_Tol And objRs!Hora < Hora_Tol And objRs!Hora < S1 Then
                        'Genero anormalidad de salida temprano */
                        objRs.Update objRs.Fields("anornro"), 6
                        objRs.UpdateBatch
                        ' Si acumula genero-hc-salida-temprano
                        If acumula_temp Then
                            objFechasHoras.RestaHs objRs!Fecha, objRs!Hora, Fecha_Tol, Hora_Tol, Tdias, Thoras, Tmin
                            tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                            Generar_Salida_Temprano objRs!Fecha, tothoras
                        End If
                        objFechasHoras.RestaXHoras FS1, S1, toldto, fecha_toldto, hora_toldto
                        ok = objFechasHoras.ValidarHora(hora_toldto)
                        If (objRs!Fecha = fecha_toldto And objRs!Hora < hora_toldto) And (acumula_dto = True) Then
                            objFechasHoras.RestaHs fecha_toldto, hora_toldto, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                            tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                            'Descontar_Horas objRs!Fecha, tothoras
                            'FGZ - 01/12/2005
                            Call Descontar_Horas(objRs!Fecha, tothoras, 1)
                        End If
                    Else
                        'Entra en tolerancia. La fichada debe ser la del turno */
                        If (objRs!Fecha = Fecha_Tol) And (objRs!Hora < S1) Then
                            ' Sale Temprano
                            objRs.Update objRs.Fields("Fecha"), FS1
                            objRs.Update objRs.Fields("Hora"), S1
                            objRs.UpdateBatch
                        End If
                    End If
                    ' Asigno el registro como la segunda componente del par1
                    objRs.Update objRs.Fields("par"), 1
                    objRs.UpdateBatch
                    cant2 = 1
                End If
         Else
            'No es available wf-turno faltan registraciones en el 1er Embudo */
            'Busco la 1er reg del 2do embudo para pasarla al 1er embudo*/
            If cant2 >= 2 Then
                emb_ante = True
                StrSql = " SELECT * FROM " & TTempWFTurno & " WHERE (evenro = 2) AND ((fecha > " & ConvFecha(fv2) & " ) OR (fecha = " & ConvFecha(fv2) & " AND  hora > '" & v2 & "'))" & _
                         " AND ((fecha = " & ConvFecha(fv3) & " AND hora <= '" & v3 & "') OR (fecha < " & ConvFecha(fv3) & ")) AND par = 0 "
                OpenRecordset StrSql, objRs, adLockOptimistic
                If Not objRs.EOF Then
                     'Calculo tolerancia tarde
                     objFechasHoras.SumoHoras FE1, E1, tol, Fecha_Tol, Hora_Tol
                     ok = objFechasHoras.ValidarHora(Hora_Tol)
                  If objRs!Fecha = Fecha_Tol And objRs!Hora > Hora_Tol And objRs!Hora > E1 Then
                      ' Genero anormalidad de llegada tarde */
                      objRs.Update objRs.Fields("anornro"), 5
                      objRs.UpdateBatch
                      ' Si acumula genero-hc-llegada-tarde
                      If acumula Then
                          objFechasHoras.RestaHs Fecha_Tol, Hora_Tol, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                          tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                          Generar_Llegada_Tarde objRs!Fecha, tothoras, objRs!Hora, Hora_Tol
                      End If
                      objFechasHoras.SumoHoras FE1, E1, toldto, fecha_toldto, hora_toldto
                      ok = objFechasHoras.ValidarHora(hora_toldto)
                      If (objRs!Fecha = fecha_toldto And objRs!Hora > hora_toldto) And (acumula_dto = True) Then
                          objFechasHoras.RestaHs fecha_toldto, hora_toldto, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                          tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                          'Descontar_Horas objRs!Fecha, tothoras
                          'FGZ - 01/12/2005
                          Call Descontar_Horas(objRs!Fecha, tothoras, 1)
                      End If
                Else
                    ' Entra en tolerancia. La fichada debe ser la del turno */
                      If (objRs!Fecha = Fecha_Tol) And (objRs!Hora > E1) Then
                         ' Llega tarde
                         objRs.Update objRs.Fields("Fecha"), FE1
                         objRs.Update objRs.Fields("Hora"), E1
                         objRs.UpdateBatch
                      End If
                End If
                ' Asigno el registro como la primer componente del par1
                objRs.Update objRs.Fields("par"), 1
                objRs.UpdateBatch
            Else
                If cant2 = 0 Then
                    StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 2 AND turnro = " & Nro_Turno
                    OpenRecordset StrSql, objRs
                    If Not objRs.EOF Then
                        GeneraTraza Empleado.Ternro, p_fecha, "No está configurado el Tipo de Hora Ausencia para el Turno:", Str(Nro_Turno)
                        'PREGUNTA !! Salgo de todo el prc
                        ' deberia salir de todo el procedimiento
                        Exit Sub
                    End If
                    
                    'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
                    'Fecha_Generacion = CalcularFechadeGeneracion(Nro_Subturno, p_fecha, fecpar1, fecpar2, Cambio_dia)
                    Fecha_Generacion = p_fecha
                    'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
                    
                    StrSql = "INSERT INTO gti_horcumplido(horas, horcant,hordesde,horestado,horhasta,horhoradesde," & _
                            "horhorahasta,hormanual,horvalido,jusnro,jusnro2,normnro,normnro2,Ternro," & _
                            "thnro,turnro,regent,horrealent,regsal,horrealsal,Empleg,horfecrep,horfecgen) VALUES ('00:00'," & _
                            0 & "," & ConvFecha(FE1) & ",' '," & ConvFecha(FS1) & ",'0000','0000'," & _
                            CInt(False) & "," & CInt(True) & "," & ValorNulo & "," & ValorNulo & ",1,2," & _
                            Empleado.Ternro & "," & objRs!thnro & "," & Nro_Turno & "," & _
                            ValorNulo & ",''," & ValorNulo & ",''," & _
                            Empleado.Legajo & "," & ConvFecha(Fecha_Generacion) & "," & ConvFecha(p_fecha) & ")"
    
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
         End If
         'Busco registracion en el segundo embudo
          StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE evenro = 2 AND ((fecha > " & ConvFecha(fv2) & ") OR (fecha = " & ConvFecha(fv2) & " AND hora > '" & v2 & ")) " & _
                      " AND ((fecha = " & ConvFecha(fv3) & " AND hora <= '" & v3 & "') OR (fecha < " & ConvFecha(fv3) & ")) AND wf-turno.par = 0 " & _
                      " ORDER BY Codigo "
         OpenRecordset StrSql, objRs
         If cant2 >= 2 And cant3 = 0 And Cant_emb > 2 Then
            ' El ultimo va al 3 embudo y tomo el anterior para el 2 embudo
             objRs.MoveLast
             objRs.MovePrevious
         Else
             objRs.MoveLast
         End If
         If Not objRs.EOF Then
             ' Calculo tolerancia temprano */
            objFechasHoras.RestaXHoras FS1, S1, toltemp, Fecha_Tol, Hora_Tol
            ok = objFechasHoras.ValidarHora(Hora_Tol)
            If objRs!Fecha = Fecha_Tol And objRs!Hora < Hora_Tol Then
                ' Genero anormalidad de salida temprano */
                objRs.Update objRs.Fields("anornro"), 6
                objRs.UpdateBatch
                'Si acumula genero-hc-salida-temprano
                If acumula_temp Then
                     objFechasHoras.RestaHs objRs!Fecha, objRs!Hora, Fecha_Tol, Hora_Tol, Tdias, Thoras, Tmin
                     tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                     Generar_Salida_Temprano objRs!Fecha, tothoras
                End If
                objFechasHoras.RestaXHoras FS1, S1, toldto, fecha_toldto, hora_toldto
                ok = objFechasHoras.ValidarHora(hora_toldto)
                If (objRs!Fecha = fecha_toldto And objRs!Hora < hora_toldto) And (acumula_dto = True) Then
                     objFechasHoras.RestaHs fecha_toldto, hora_toldto, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                     tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                     'Descontar_Horas objRs!Fecha, tothoras
                     'FGZ - 01/12/2005
                     Call Descontar_Horas(objRs!Fecha, tothoras, 1)
                End If
            Else
               ' Entra en tolerancia. La fichada debe ser la del turno
               If (objRs!Fecha = Fecha_Tol) And (objRs!Hora < S1) Then
                   ' Sale Temprano
                    objRs.Update objRs.Fields("Fecha"), FS1
                    objRs.Update objRs.Fields("Hora"), S1
                    objRs.UpdateBatch
               End If
            End If
            ' Asigno el registro como la segunda componente del par1 */
            objRs.Update objRs.Fields("par"), 1
            objRs.UpdateBatch
            If cant2 >= 2 And cant3 = 0 And Cant_emb > 2 Then
                ' Paso el ultimo del Embudo 2 al Embudo 3
                objRs.MoveLast
                       
                'Calculo tolerancia tarde
                objFechasHoras.SumoHoras FE2, E2, tol, Fecha_Tol, Hora_Tol
                ok = objFechasHoras.ValidarHora(Hora_Tol)
                If objRs!Fecha = Fecha_Tol And objRs!Hora > Hora_Tol And objRs!Hora > E2 Then
                     ' Genero anormalidad de llegada tarde */
                     objRs.Update objRs.Fields("anornro"), 5
                     objRs.UpdateBatch
                     'Si acumula genero_hc_llegada_tarde
                     If acumula Then
                         objFechasHoras.RestaHs Fecha_Tol, Hora_Tol, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                         tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                         Generar_Llegada_Tarde objRs!Fecha, tothoras, objRs!Hora, Hora_Tol
                     End If
                
                     objFechasHoras.SumoHoras FE2, E2, toldto, fecha_toldto, hora_toldto
                     ok = objFechasHoras.ValidarHora(hora_toldto)
                     If (objRs!Fecha = fecha_toldto And objRs!Hora > hora_toldto) And (acumula_dto = True) Then
                         objFechasHoras.RestaHs fecha_toldto, hora_toldto, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                         tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                         'Descontar_Horas objRs!Fecha, tothoras
                        'FGZ - 01/12/2005
                        Call Descontar_Horas(objRs!Fecha, tothoras, 1)
                     End If
                Else
                    ' Entra en tolerancia. La fichada debe ser la del turno */
                    If (objRs!Fecha = Fecha_Tol) And (objRs!Hora > E2) Then
                        ' Llega tarde
                        objRs.Update objRs.Fields("Fecha"), FE2
                        objRs.Update objRs.Fields("Hora"), E2
                        objRs.UpdateBatch
                    End If
                End If
               ' Asigno el registro como la primer componente del par2 */
                objRs.Update objRs.Fields("par"), 2
                objRs.UpdateBatch
                cant3 = 1
            End If
        'End If
        Else
           ' No es available wf-turno faltan registraciones obligatorias en el 2do Embudo */
           If cant3 = 0 And cant2 = 0 Then
                ' Es un caso donde no salio a almorzar
                ' Politica para ver si paga almuerzo por no tomarlo
                Call Politica(390)
                If Not paga_almuerzo Then
                  ' Generar HC teorico con anormalidad sin reg
                  InsertarWFTurno Empleado.Ternro, S1, FS1, ValorNulo, ValorNulo, 2, False, ValorNulo, 1, 2
                End If
           End If
           If cant3 = 1 Then
                ' Generar HC teorico con anormalidad sin reg */
                InsertarWFTurno Empleado.Ternro, S1, FS1, ValorNulo, ValorNulo, 2, False, ValorNulo, 1, 2
                'PREGUNTA : IGUAL QUE ARRIBA
           End If
           If cant3 >= 2 Then
              ' Busco registracion en el tercer embudo */
              StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE evenro = 2 AND ((fecha > " & ConvFecha(fv3) & ") OR (fecha = " & ConvFecha(fv3) & " AND hora > '" & v3 & ")) " & _
                      " AND ((fecha = " & ConvFecha(fv4) & " AND hora <= '" & v4 & "') OR (fecha < " & ConvFecha(fv4) & ")) AND wf-turno.par = 0 " & _
                      " ORDER BY Codigo "
              OpenRecordset StrSql, objRs, adLockOptimistic
              objRs.MoveFirst
              If objRs.EOF Then Exit Sub
              ' Calculo tolerancia temprano */
              objFechasHoras.RestaXHoras FS1, S1, toltemp, Fecha_Tol, Hora_Tol
              ok = objFechasHoras.ValidarHora(Hora_Tol)
              If objRs!Fecha = Fecha_Tol And objRs!Hora < Hora_Tol And objRs!Hora < S1 Then
                  ' Genero anormalidad de salida temprano */
                  objRs.Update objRs.Fields("anornro"), 6
                  objRs.UpdateBatch
                  ' Si acumula genero-hc-salida-temprano */
                  If acumula_temp Then
                      objFechasHoras.RestaHs objRs!Fecha, objRs!Hora, Fecha_Tol, Hora_Tol, Tdias, Thoras, Tmin
                      tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                      Generar_Salida_Temprano objRs!Fecha, tothoras
                  End If
                  objFechasHoras.RestaXHoras FS1, S1, toldto, fecha_toldto, hora_toldto
                  ok = objFechasHoras.ValidarHora(hora_toldto)
                  If (objRs!Fecha = fecha_toldto And objRs!Hora < hora_toldto) And (acumula_dto = True) Then
                      objFechasHoras.RestaHs fecha_toldto, hora_toldto, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                      tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                      'Descontar_Horas objRs!Fecha, tothoras
                     'FGZ - 01/12/2005
                     Call Descontar_Horas(objRs!Fecha, tothoras, 1)
                  End If
              Else
                  ' Entra en tolerancia. La fichada debe ser la del turno
                  If (objRs!Fecha = Fecha_Tol) And (objRs!Hora < S1) Then
                      ' Sale Temprano
                       objRs.Update objRs.Fields("Fecha"), FS1
                       objRs.Update objRs.Fields("Hora"), S1
                       objRs.UpdateBatch
                  End If
              End If
              objRs.Update objRs.Fields("par"), 1
              objRs.UpdateBatch
           End If
        End If
    End If
    If v4 <> "" And v5 <> "" Then
        'Busco registracion en el tercer embudo */
         StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE evenro = 2 AND ((fecha > " & ConvFecha(fv3) & ") OR (fecha = " & ConvFecha(fv3) & " AND hora > '" & v3 & ")) " & _
                  " AND ((fecha = " & ConvFecha(fv4) & " AND hora <= '" & v4 & "') OR (fecha < " & ConvFecha(fv4) & ")) AND wf-turno.par = 0 " & _
                  " ORDER BY Codigo "
         OpenRecordset StrSql, objRs, adLockOptimistic
         objRs.MoveFirst
        If Not objRs.EOF Then
           ' Calculo tolerancia tarde */
           objFechasHoras.SumoHoras FE2, E2, tol, Fecha_Tol, Hora_Tol
           ok = objFechasHoras.ValidarHora(Hora_Tol)
           If objRs!Fecha = Fecha_Tol And objRs!Hora > Hora_Tol And objRs!Hora > E2 Then
                ' Genero anormalidad de llegada tarde */
                objRs.Update objRs.Fields("anornro"), 5
                objRs.UpdateBatch
                ' Si acumula genero-hc-llegada-tarde
                If acumula Then
                    objFechasHoras.RestaHs Fecha_Tol, Hora_Tol, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                    tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                    Generar_Llegada_Tarde objRs!Fecha, tothoras, objRs!Hora, Hora_Tol
                End If
                objFechasHoras.SumoHoras FE2, E2, toldto, fecha_toldto, hora_toldto
                ok = objFechasHoras.ValidarHora(hora_toldto)
                If (objRs!Fecha = fecha_toldto And objRs!Hora > hora_toldto) And (acumula_dto = True) Then
                     objFechasHoras.RestaHs fecha_toldto, hora_toldto, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                     tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                     'Descontar_Horas objRs!Fecha, tothoras
                     'FGZ - 01/12/2005
                     Call Descontar_Horas(objRs!Fecha, tothoras, 1)
                End If
           Else
                ' Entra en tolerancia. La fichada debe ser la del turno */
                If (objRs!Fecha = Fecha_Tol) And (objRs!Hora > E2) Then
                    ' Llega tarde
                    objRs.Update objRs.Fields("Fecha"), FE2
                    objRs.Update objRs.Fields("Hora"), E2
                    objRs.UpdateBatch
                End If
           End If
           ' Asigno el registro como la primer componente del par2 */
           objRs.Update objRs.Fields("par"), 2
           objRs.UpdateBatch
           ' Busco el ultimo del Embudo 3 y lo paso al Embudo 4 */
           If cant3 >= 2 And cant4 = 0 And Cant_emb > 3 Then
                objRs.MoveLast
                ' Calculo tolerancia temprano */
                objFechasHoras.RestaXHoras FS2, S2, toltemp, Fecha_Tol, Hora_Tol
                ok = objFechasHoras.ValidarHora(Hora_Tol)
                If objRs!Fecha = Fecha_Tol And objRs!Hora < Hora_Tol And objRs!Hora < S2 Then
                    ' Genero anormalidad de salida temprano
                    objRs.Update objRs.Fields("anornro"), 6
                    objRs.UpdateBatch
                    ' Si acumula genero-hc-salida-temprano
                    If acumula_temp Then
                        objFechasHoras.RestaHs objRs!Fecha, objRs!Hora, Fecha_Tol, Hora_Tol, Tdias, Thoras, Tmin
                        tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                        Generar_Salida_Temprano objRs!Fecha, tothoras
                    End If
                    objFechasHoras.RestaXHoras FS2, S2, toldto, fecha_toldto, hora_toldto
                    ok = objFechasHoras.ValidarHora(hora_toldto)
                    If (objRs!Fecha = fecha_toldto And objRs!Hora < hora_toldto) And (acumula_dto = True) Then
                         objFechasHoras.RestaHs fecha_toldto, hora_toldto, objRs!Fecha, objRs!Hora, Tdias, Thoras, Tmin
                         tothoras = (Tdias * 24) + (Thoras + (Tmin / 60))
                        'Descontar_Horas objRs!Fecha, tothoras
                        'FGZ - 01/12/2005
                        Call Descontar_Horas(objRs!Fecha, tothoras, 1)
                    End If
                Else
                    ' Entra en tolerancia. La fichada debe ser la del turno
                    If (objRs!Fecha = Fecha_Tol) And (objRs!Hora < S2) Then
                        ' Sale Temprano
                        objRs.Update objRs.Fields("Fecha"), FS2
                        objRs.Update objRs.Fields("Hora"), S2
                        objRs.UpdateBatch
                    End If
                End If
                ' Asigno el registro como la segunda componente del par2
                objRs.Update objRs.Fields("par"), 2
                objRs.UpdateBatch
           End If
        End If
    Else
        ' No es available wf-turno faltan registraciones obligatorias en el Embudo3
        If cant4 = 1 And ((cant2 = 0 And Not paga_almuerzo) Or (cant2 >= 1 And emb_prox)) Then
             ' Genero HC teorico con anor de falta reg */
             InsertarWFTurno Empleado.Ternro, E2, FE2, ValorNulo, ValorNulo, 1, True, ValorNulo, 2, 2
        End If
    End If
End If

End Sub



Private Sub Generar_Salida_Temprano(Fecha_Tol As Date, Horas_Tol As Single)

Dim TH_tol As Long
Dim Horas_Acum As Single
Dim hora_red As String
Dim Hora_a_Red As String

Dim objRs As New ADODB.Recordset

    Horas_Acum = 0
    If Horas_Tol = 0 Then Exit Sub
 
    'Conhornro = 5 'Horas de Dto
    StrSql = "SELECT * FROM gti_config_tur_hor WHERE (conhornro = 5) AND (turnro = " & Nro_Turno & ")"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        TH_tol = objRs!thnro
    Else
        GeneraTraza Empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Salida Temprano para el Turno:", Str(Nro_Turno)
        Exit Sub
    End If
          
    objFechasHoras.Convertir_A_Hora Horas_Tol * 60, Hora_a_Red
    'objFechasHoras.Redondeo_Horas_Tipo hora_a_red, objRs!conhredondeo, horas_tol
    objFechasHoras.Redondeo_Horas_Tipo Hora_a_Red, IIf(Not EsNulo(objRs!conhredondeo), objRs!conhredondeo, 0), Horas_Tol
    
    'Call ValidarTipoDeHora(th_tol, nro_turno, tipo_hora)
    
    'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
    'Fecha_Generacion = CalcularFechadeGeneracion(Nro_Subturno, p_fecha, fecpar1, fecpar2, Cambio_dia)
    Fecha_Generacion = p_fecha
    'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
    
    StrSql = "INSERT INTO gti_horcumplido(horas, horcant,hordesde,horestado,horhasta,horhoradesde," & _
            "horhorahasta,hormanual,horvalido,jusnro,jusnro2,normnro,normnro2,Ternro," & _
            "thnro,turnro,regent,horrealent,regsal,horrealsal,Empleg,horfecrep,horfec) VALUES (" & _
            CHoras(Horas_Tol, 60) & "," & Horas_Tol & "," & ConvFecha(Fecha_Tol) & ",' '," & ConvFecha(Fecha_Tol) & ",'',''," & _
            CInt(False) & "," & CInt(True) & "," & ValorNulo & "," & ValorNulo & ",5," & _
            5 & "," & Empleado.Ternro & "," & TH_tol & "," & Nro_Turno & "," & _
            ValorNulo & ",''," & ValorNulo & ",''," & _
            Empleado.Legajo & "," & ConvFecha(Fecha_Generacion) & "," & ConvFecha(p_fecha) & ")"
        
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub

Private Sub BlanquearVariables()

    fecha_desde = p_fecha
    hora_desde = ""
    fecha_hasta = p_fecha
    hora_hasta = ""
    
    Nro_Turno = 0
    Tipo_Turno = 0
    Nro_Grupo = 0
    Nro_fpgo = 0
    Fecha_Inicio = p_fecha
    Trabaja = False
    Orden_Dia = 0
    Nro_Dia = 0
    Nro_Subturno = 0
    Dia_Libre = False
     
    'FGZ - 26/04/2007 - faltaba esta inicializacion
    UsaFeriadoConControl = False
    Pasa_de_Dia = False
     
    'FGZ - 17/05/2007 - Faltaban estas inicializaciones
    p_turcomp = False
    Nro_Justif = 0
    justif_turno = False
    Tiene_Justif = False
    P_Asignacion = False
    Nro_Dia_Original = 0
    'FGZ - 17/05/2007 - Faltaban estas inicializaciones
     
     
End Sub


Public Sub Descontar_Horas_OLD(fecha_dto As Date, horas_dto As Single)
Dim TH_Dto As Long
Dim Horas_Acum As Single
Dim hora_red   As String
Dim Hora_a_Red As String
Dim fraccionamiento As Single
Dim objRs As New ADODB.Recordset

  Horas_Acum = 0
  If horas_dto = 0 Then Exit Sub
 
  StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 3 AND turnro = " & Nro_Turno
  OpenRecordset StrSql, objRs
  If Not objRs.EOF Then
      TH_Dto = objRs!thnro
  Else
      GeneraTraza Empleado.Ternro, p_fecha, "No está configurado el Tipo de Hora Descuento para el Turno:", Str(Nro_Turno)
      Exit Sub
  End If
  fraccionamiento = objRs!conhfraccionamiento / 60
  Do While fraccionamiento < horas_dto
     Horas_Acum = Horas_Acum + fraccionamiento
     horas_dto = horas_dto - fraccionamiento
  Loop
     
  Horas_Acum = Horas_Acum + fraccionamiento
       
  objFechasHoras.Convertir_A_Hora Horas_Acum * 60, Hora_a_Red
  'objFechasHoras.Redondeo_Horas_Tipo hora_a_red, objRs!conhredondeo, horas_acum
  objFechasHoras.Redondeo_Horas_Tipo Hora_a_Red, IIf(Not EsNulo(objRs!conhredondeo), objRs!conhredondeo, 0), Horas_Acum

  'Call ValidarTipoDeHora(th_dto, nro_turno, tipo_hora)
  StrSql = "INSERT INTO gti_horcumplido(horas, horcant,hordesde,horestado,horhasta,horhoradesde," & _
            "horhorahasta,hormanual,horvalido,jusnro,jusnro2,normnro,normnro2,Ternro," & _
            "thnro,turnro,regent,horrealent,regsal,horrealsal,Empleg,horfecrep) VALUES (" & _
            Horas_Acum & "," & ConvFecha(fecha_dto) & ",' '," & ConvFecha(fecha_dto) & ",'',''," & _
            CInt(False) & "," & CInt(True) & "," & ValorNulo & "," & ValorNulo & ",5," & _
            5 & "," & Empleado.Ternro & "," & TH_Dto & "," & Nro_Turno & "," & _
            ValorNulo & ",''," & ValorNulo & ",''," & _
            Empleado.Legajo & "," & ConvFecha(p_fecha) & ")"
          
  objConn.Execute StrSql, , adExecuteNoRecords

End Sub

Private Sub Generar_Llegada_Tarde(Fecha_Tol As Date, Horas_Tol As Single, HDesde As String, HHasta As String)

Dim TH_tol As Long
Dim Horas_Acum As Single
Dim hora_red As String
Dim Hora_a_Red As String

Dim objRs As New ADODB.Recordset

    Horas_Acum = 0
    If Horas_Tol = 0 Then Exit Sub
 
    'Conhornro = 4 'Horas de Dto
    StrSql = "SELECT * FROM gti_config_tur_hor WHERE (conhornro = 4) AND (turnro = " & Nro_Turno & ")"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        TH_tol = objRs!thnro
    Else
        GeneraTraza Empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Llegada Tarde para el Turno:", Str(Nro_Turno)
        Exit Sub
    End If
          
    objFechasHoras.Convertir_A_Hora Horas_Tol * 60, Hora_a_Red
    'objFechasHoras.Redondeo_Horas_Tipo hora_a_red, objRs!conhredondeo, horas_tol
    objFechasHoras.Redondeo_Horas_Tipo Hora_a_Red, IIf(Not EsNulo(objRs!conhredondeo), objRs!conhredondeo, 0), Horas_Tol
  
    'Call ValidarTipoDeHora(th_tol, nro_turno, tipo_hora)
    
    'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
    'Fecha_Generacion = CalcularFechadeGeneracion(Nro_Subturno, p_fecha, fecpar1, fecpar2, Cambio_dia)
    Fecha_Generacion = p_fecha
    'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
    
    StrSql = "INSERT INTO gti_horcumplido(horas, horcant,hordesde,horestado,horhasta,horhoradesde," & _
            "horhorahasta,hormanual,horvalido,jusnro,jusnro2,normnro,normnro2,Ternro," & _
            "thnro,turnro,regent,horrealent,regsal,horrealsal,Empleg,horfecrep,horfecgen) VALUES (" & _
            CHoras(Horas_Tol, 60) & "," & Horas_Tol & "," & ConvFecha(Fecha_Tol) & ",' '," & ConvFecha(Fecha_Tol) & ",'" & HDesde & "','" & HHasta & "'," & _
            CInt(False) & "," & CInt(True) & "," & ValorNulo & "," & ValorNulo & ",5," & _
            5 & "," & Empleado.Ternro & "," & TH_tol & "," & Nro_Turno & "," & _
            ValorNulo & ",''," & ValorNulo & ",''," & _
            Empleado.Legajo & "," & ConvFecha(Fecha_Generacion) & "," & ConvFecha(p_fecha) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub
       
Private Sub initVariablesTurno(ByRef T As BuscarTurno)
   p_turcomp = T.Compensa_Turno
   Nro_Grupo = T.Empleado_Grupo
   Nro_Justif = T.Justif_Numero
   justif_turno = T.justif_turno
   Tiene_Justif = T.Tiene_Justif
   Fecha_Inicio = T.FechaInicio
   Nro_fpgo = T.Numero_FPago
   Nro_Turno = T.Turno_Numero
   tiene_turno = T.tiene_turno
   Tipo_Turno = T.Turno_Tipo
   P_Asignacion = T.Tiene_PAsignacion
   
End Sub
Private Sub initVariablesDia(ByRef D As BuscarDia)
   Dia_Libre = D.Dia_Libre
   Nro_Dia = D.Numero_Dia
   Nro_Subturno = D.SubTurno_Numero
   Orden_Dia = D.Orden_Dia
   Trabaja = D.Trabaja
End Sub



Public Sub Horas_Sabado_Domingo_RHPRO(Fecha As Date)


Dim TotHor As Single
Dim tip_hora As Integer

Dim Horas_Dia As Single
Dim max_horas As Single
Dim horas_min As Single
Dim Horas_Oblig As Single

Call buscar_horas_turno(Horas_Oblig, max_horas, horas_min)


If P_Asignacion Then
    StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro =" & Empleado.Ternro & " ) and (" & _
             "gttempdesde <= " & ConvFecha(p_fecha) & ") and (" & _
             ConvFecha(p_fecha) & " <= gttemphasta)"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Horas_Dia = objRs!diacanthoras
    End If
Else
    StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_Subturno
    StrSql = StrSql & " ORDER BY diaorden"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
         Horas_Dia = objRs!diacanthoras
    End If
        
End If


If (Weekday(Fecha) = 7) And (Horas_Oblig > 5) Then
    StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_Subturno
    StrSql = StrSql & " ORDER BY diaorden"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'objRs.MoveNext
        Horas_Oblig = objRs!diacanthoras
    End If
End If

'StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_subturno
'StrSql = StrSql & " ORDER BY diaorden"
'OpenRecordset StrSql, objRs
'If Not objRs.EOF Then
'
'    horas_dia = objRs!diacanthoras
'
'    If (Weekday(Fecha) = 7) And (horas_oblig > 5) Then
'        objRs.MoveNext
'        horas_oblig = objRs!diacanthoras
'    End If
'
'End If

If (Weekday(Fecha) = 1 Or (Dia_Libre) Or (esFeriado)) Then
 TotHor = Horas_Dia
 tip_hora = 57
End If

If (Weekday(Fecha) = 7) And (Not Dia_Libre) And (Not esFeriado) Then
 TotHor = Horas_Dia - Horas_Oblig
 tip_hora = 56
End If

'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
'Fecha_Generacion = CalcularFechadeGeneracion(Nro_Subturno, p_fecha, fecpar1, fecpar2, Cambio_dia)
Fecha_Generacion = Fecha
'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------

StrSql = "INSERT INTO gti_horcumplido(horas, horcant,hordesde,horestado,horhasta," & _
         "hormanual,horvalido,jusnro,jusnro2,normnro,normnro2,Ternro," & _
         "thnro,turnro,regent,horrealent,regsal,horrealsal,Empleg,horfecrep,horfecgen) VALUES (" & _
         CHoras(Round(TotHor, 2), 60) & "," & Round(TotHor, 2) & "," & ConvFecha(Fecha) & ",' '," & ConvFecha(Fecha) & "," & _
         CInt(False) & "," & CInt(True) & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
         ValorNulo & "," & Empleado.Ternro & "," & tip_hora & "," & Nro_Turno & "," & _
         ValorNulo & ",''," & ValorNulo & ",''," & _
         Empleado.Legajo & "," & ConvFecha(Fecha_Generacion) & "," & ConvFecha(Fecha) & ")"
          
'FLog.writeline StrSql
          
objConn.Execute StrSql, , adExecuteNoRecords


End Sub


Public Sub Horas_Sabado_Domingo(Fecha As Date)

Dim TotHor As Single
Dim tip_hora As Integer

Dim Horas_Dia As Single
Dim max_horas As Single
Dim horas_min As Single
Dim Horas_Oblig As Single
Dim HoraFinDeSemana As Integer

Dim rs_TipHora As New ADODB.Recordset
Dim objRs As New ADODB.Recordset


    Flog.writeline "Horas Sabado y Domingo"
    
    ' Segun gti_config_hora
    HoraFinDeSemana = 44

     ' busco el tipo de hora Fin de Semana
     StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraFinDeSemana & _
                " AND turnro = " & Nro_Turno & " ORDER BY conhornro ASC, turnro ASC"
     OpenRecordset StrSql, rs_TipHora

     If Not rs_TipHora.EOF Then
         tip_hora = rs_TipHora!thnro
     Else
        If depurar Then
            Flog.writeline "Error, no esta configurado el tipo de hora  -" & Fecha
        End If
        Exit Sub
     End If

    If rs_TipHora.State = adStateOpen Then rs_TipHora.Close

    Call buscar_horas_turno(Horas_Oblig, max_horas, horas_min)

    StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_Subturno
    StrSql = StrSql & " ORDER BY diaorden"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Horas_Dia = objRs!diacanthoras
    End If
    
    If Weekday(Fecha) = 1 Then
     TotHor = Horas_Dia
     'tip_hora = 57
    End If
    If Weekday(Fecha) = 7 Then
     TotHor = Horas_Dia - Horas_Oblig
     'tip_hora = 56
    End If
    
    'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
    'Fecha_Generacion = CalcularFechadeGeneracion(Nro_Subturno, p_fecha, fecpar1, fecpar2, Cambio_dia)
    Fecha_Generacion = Fecha
    'FGZ - 27/07/2009  ------------------------------------------------------------------------------------------------
    
    StrSql = "INSERT INTO gti_horcumplido(horas, horcant,hordesde,horestado,horhasta," & _
             "hormanual,horvalido,jusnro,jusnro2,normnro,normnro2,Ternro," & _
             "thnro,turnro,regent,horrealent,regsal,horrealsal,Empleg,horfecrep,horfecgen) VALUES (" & _
             CHoras(TotHor, 60) & "," & TotHor & "," & ConvFecha(Fecha) & ",' '," & ConvFecha(Fecha) & "," & _
             CInt(False) & "," & CInt(True) & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
             ValorNulo & "," & Empleado.Ternro & "," & tip_hora & "," & Nro_Turno & "," & _
             ValorNulo & ",''," & ValorNulo & ",''," & _
             Empleado.Legajo & "," & ConvFecha(Fecha_Generacion) & "," & ConvFecha(Fecha) & ")"
    objConn.Execute StrSql, , adExecuteNoRecords

    If depurar Then
        Flog.writeline "Fin Horas Sabado y Domingo"
    End If

    'Libero y dealoco
    If rs_TipHora.State = adStateOpen Then rs_TipHora.Close
    If objRs.State = adStateOpen Then objRs.Close
    
    Set rs_TipHora = Nothing
    Set objRs = Nothing
End Sub




Public Sub Horas_Sabado_Domingo_old(Fecha As Date)


Dim TotHor As Single
Dim tip_hora As Integer

Dim Horas_Dia As Single
Dim max_horas As Single
Dim horas_min As Single
Dim Horas_Oblig As Single

Call buscar_horas_turno(Horas_Oblig, max_horas, horas_min)

StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_Subturno
StrSql = StrSql & " ORDER BY diaorden"
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
        
    Horas_Dia = objRs!diacanthoras
        
End If

If Weekday(Fecha) = 1 Then
 TotHor = Horas_Dia
 tip_hora = 57

End If
If Weekday(Fecha) = 7 Then
 TotHor = Horas_Dia - Horas_Oblig
 tip_hora = 56
End If

'Call ValidarTipoDeHora(tip_hora, nro_turno, tipo_hora)
StrSql = "INSERT INTO gti_horcumplido(horas, horcant,hordesde,horestado,horhasta," & _
         "hormanual,horvalido,jusnro,jusnro2,normnro,normnro2,Ternro," & _
         "thnro,turnro,regent,horrealent,regsal,horrealsal,Empleg,horfecrep) VALUES (" & _
         TotHor & "," & ConvFecha(Fecha) & ",' '," & ConvFecha(Fecha) & "," & _
         CInt(False) & "," & CInt(True) & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
         ValorNulo & "," & Empleado.Ternro & "," & tip_hora & "," & Nro_Turno & "," & _
         ValorNulo & ",''," & ValorNulo & ",''," & _
         Empleado.Legajo & "," & ConvFecha(Fecha) & ")"
          
'FLog.writeline StrSql
          
objConn.Execute StrSql, , adExecuteNoRecords


End Sub


Sub Horario_Trasnoche(Fecha As Date, Ternro As Long)

Dim rs As New ADODB.Recordset

Dim objBTurno As New BuscarTurno
Dim objBDia As New BuscarDia
Dim objFeriado As New Feriado
Dim objFechasHoras As New FechasHoras

Dim FechaDesde As Date
Dim FechaHasta As Date

Dim aux_nrodia As Long

Call BorrarTempTable(TTempWFTurno)
Call CreateTempTable(TTempWFTurno)

'Call BorrarTempTable("WF_TURNO")
'Call CreateTempTable("WF_TURNO")

'Call BorrarTablaWFTurno
'Call CreateTempWFTurno

FechaDesde = DateAdd("d", -1, Fecha)
FechaHasta = Fecha

p_fecha = FechaDesde

If depurar Then
    Flog.writeline "Horario Trasnoche"
End If

' Si es feriado no proceso
Set objFeriado.Conexion = objConn
'Set objFeriado.ConexionTraza = CnTraza
esFeriado = objFeriado.Feriado(FechaDesde, Ternro, depurar)
If esFeriado Then
    Exit Sub
End If

' Si no tiene turno no proceso
Set objBTurno.Conexion = objConn
'Set objBTurno.ConexionTraza = CnTraza
objBTurno.Buscar_Turno FechaDesde, Ternro, depurar
initVariablesTurno objBTurno
'Flog.writeline Now & " Bturno"
If Not tiene_turno And Not Tiene_Justif Then
    'MyRollbackTrans
    Exit Sub
End If


' Busco el último horario de entrada del día anterior
If tiene_turno Then
    Set objBDia.Conexion = objConn
    'Set objBDia.ConexionTraza = cnTraza
    objBDia.Buscar_Dia p_fecha, Fecha_Inicio, Nro_Turno, Ternro, P_Asignacion, depurar
    initVariablesDia objBDia
    'Flog.writeline Now & " BDia"
    Call Horario_Teorico
    'Flog.writeline Now & " HorarioTeorico"
End If


aux_nrodia = Nro_Dia

If Not Dia_Libre Then

    If E3 <> "" Then
        hora_desde = S2
        fecha_desde = FechaDesde
    Else
        hora_desde = S1
        fecha_desde = FechaDesde
    End If
Else
    hora_desde = "0000"
    fecha_desde = FechaHasta
End If

' Busco la primer entrada del siguiente día
objBDia.Buscar_Dia p_fecha + 1, Fecha_Inicio, Nro_Turno, Ternro, P_Asignacion, depurar
initVariablesDia objBDia
'Flog.writeline Now & " BDia"
Call Horario_Teorico

hora_hasta = E1
fecha_hasta = FechaHasta

Call Politica(80)
Call Politica(85)

Nro_Dia = aux_nrodia

Call Politica(1000)
If depurar Then
    Flog.writeline "Fin Horario Trasnoche"
End If

End Sub

Private Sub FechaProceso()
' FGZ - 20/09/2004
' Esto proceso fué hecho por EPL 23/08/2004
    If P_Asignacion Then
        StrSql = "SELECT * FROM gti_detturtemp WHERE ternro= " & Empleado.Ternro & _
        " AND gttempdesde <= " & ConvFecha(p_fecha) & " AND gttemphasta >= " & ConvFecha(p_fecha)
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            If objRs!diaanterior = 1 Then
                p_fecha = p_fecha - 1
                'Fecha = Fecha - 1
            End If
        End If
    Else
        StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            If objRs!diaanterior = 1 Then
                p_fecha = p_fecha - 1
                'Fecha = Fecha - 1
            End If
        End If
    End If
End Sub


Public Function Primer_Justificacion(ByVal SQLString As String, ByVal Codigo As Long, ByVal Empleado As Long) As Boolean
Dim objRs As New ADODB.Recordset

    Primer_Justificacion = False
    
    OpenRecordset SQLString, objRs
    If Not objRs.EOF Then
        objRs.MoveFirst
        Primer_Justificacion = objRs!jusnro = Codigo
    End If
    
If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing
End Function

Public Function Ultima_Justificacion(ByVal SQLString As String, ByVal Codigo As Long, ByVal Empleado As Long) As Boolean
Dim objRs As New ADODB.Recordset

    Ultima_Justificacion = False
    
    OpenRecordset SQLString, objRs
    If Not objRs.EOF Then
        objRs.MoveLast
        Ultima_Justificacion = objRs!jusnro = Codigo
    End If
    
If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing
End Function


Public Sub Insertar_GTI_Proc_Emp(ByVal Ternro As Long, ByVal Fecha As Date)
' --------------------------------------------------------------
' Descripcion: Genera la informacion del dia procesado.
' Autor: FGZ - 27/10/2005
' Ultima modificacion: FGZ - 30/11/2006 modificaciones para Horario_Flexible_Rotativo
'                      Diego Rosso - 12/11/2007 - PARA HORARIO FLEXIBLE: Cuando ModificaHT es falso graba el horario que le corresponderia en el dia
'                                                   y cuando es verdadero graba el numero de dia calculado.
' --------------------------------------------------------------
Dim rs_gti_Proc_Emp As New ADODB.Recordset

        On Error GoTo ME_Local
        
     
        
        StrSql = "SELECT ternro FROM gti_proc_emp WHERE ternro =" & Ternro
        StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
        OpenRecordset StrSql, rs_gti_Proc_Emp
        If rs_gti_Proc_Emp.EOF Then
            StrSql = "INSERT INTO gti_proc_emp (ternro,fecha,turnro,fpgnro,dianro,feriado,jusnro,pasig,dialibre,trabaja,estrnro"
            
            'EAM- (v5.72) - Se saco el if para que inserte siempre el horario teorico y se dejo solo adentro la propiedad manual
            If Horario_Movil Then
                StrSql = StrSql & ",manual "
            End If
            If Not EsNulo(E1) Then
                StrSql = StrSql & ",horadesde1, horahasta1 "
            End If
            If Not EsNulo(E2) Then
                StrSql = StrSql & ",horadesde2, horahasta2 "
            End If
            If Not EsNulo(E3) Then
                StrSql = StrSql & ",horadesde3, horahasta3 "
            End If
            'End If
            StrSql = StrSql & " ) VALUES ("
            StrSql = StrSql & Ternro & ","
            StrSql = StrSql & ConvFecha(Fecha) & ","
            StrSql = StrSql & Nro_Turno & ","
            StrSql = StrSql & Nro_fpgo & ","
            'FGZ - 30/11/2006 -----
            If Horario_Flexible_Rotativo Then
                If ModificaHT Then 'Diego Rosso - 12/11/2007
                    StrSql = StrSql & IIf(Nro_Dia_Original = 0, Nro_Dia, Nro_Dia_Original) & ","
                Else
                    StrSql = StrSql & Nro_Dia_Default & ","
                End If
            Else
                StrSql = StrSql & Nro_Dia & ","
            End If
            'FGZ - 30/11/2006 -----
            StrSql = StrSql & CInt(esFeriado) & ","
            StrSql = StrSql & Nro_Justif & ","
            StrSql = StrSql & CInt(P_Asignacion) & ","
            StrSql = StrSql & CInt(Dia_Libre) & ","
            StrSql = StrSql & CInt(Trabaja) & ","
            StrSql = StrSql & Nro_Grupo
          
            'EAM- (v5.72) - Se saco el if para que inserte siempre el horario teorico y se dejo solo adentro la propiedad manual
            If Horario_Movil Then
                  StrSql = StrSql & ",-1 "
            End If
            
            If Not EsNulo(E1) Then
                StrSql = StrSql & ",'" & E1 & "','" & S1 & "'"
            End If
            If Not EsNulo(E2) Then
                StrSql = StrSql & ",'" & E2 & "','" & S2 & "'"
            End If
            If Not EsNulo(E3) Then
                StrSql = StrSql & ",'" & E3 & "','" & S3 & "'"
            End If
            'End If
            StrSql = StrSql & ")"
        Else
            StrSql = "UPDATE gti_proc_emp SET "
            StrSql = StrSql & " turnro = " & Nro_Turno
            StrSql = StrSql & ",fpgnro = " & Nro_fpgo
            'FGZ - 30/11/2006 -----
'            If Horario_Flexible_Rotativo And ModificaHT Then
'                StrSql = StrSql & ",dianro = " & IIf(Nro_Dia_Original = 0, Nro_Dia, Nro_Dia_Original) & ","
'            Else
'                StrSql = StrSql & ",dianro = " & Nro_Dia
'            End If

            If Horario_Flexible_Rotativo Then
                If ModificaHT Then 'Diego Rosso - 12/11/2007
                    StrSql = StrSql & IIf(Nro_Dia_Original = 0, Nro_Dia, Nro_Dia_Original) & ","
                Else
                    StrSql = StrSql & Nro_Dia_Default & ","
                End If
            Else
                StrSql = StrSql & Nro_Dia & ","
            End If
            'FGZ - 30/11/2006 -----
            StrSql = StrSql & ",feriado = " & CInt(esFeriado)
            StrSql = StrSql & ",jusnro = " & Nro_Justif
            StrSql = StrSql & ",pasig = " & CInt(P_Asignacion)
            StrSql = StrSql & ",dialibre = " & CInt(Dia_Libre)
            StrSql = StrSql & ",trabaja = " & CInt(Trabaja)
            StrSql = StrSql & ",estrnro = " & Nro_Grupo
            If Horario_Movil Then
                StrSql = StrSql & ",manual = -1 "
                If Not EsNulo(E1) Then
                    StrSql = StrSql & ",horadesde1 ='" & E1 & "',horahasta1 ='" & S1 & "'"
                End If
                If Not EsNulo(E2) Then
                    StrSql = StrSql & ",horadesde2 ='" & E2 & "',horahasta2 ='" & S2 & "'"
                End If
                If Not EsNulo(E3) Then
                    StrSql = StrSql & ",horadesde3 ='" & E3 & "',horahasta3 ='" & S3 & "'"
                End If
            End If
            StrSql = StrSql & " WHERE TERNRO =" & Ternro
            StrSql = StrSql & " AND fecha = " & ConvFecha(Fecha)
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        If depurar Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 1) & "Dia Procesado "
            Flog.writeline Espacios(Tabulador * 1) & "SQL " & StrSql
            Flog.writeline
        End If
        
'        Flog.writeline
'        Flog.writeline Espacios(Tabulador * 1) & "Dia Procesado "
'        Flog.writeline Espacios(Tabulador * 1) & "--------------------------------------"
        
    'Libero y dealoco
    If rs_gti_Proc_Emp.State = adStateOpen Then rs_gti_Proc_Emp.Close
    Set rs_gti_Proc_Emp = Nothing
Exit Sub

ME_Local:
    Flog.writeline
    Flog.writeline "***"
    Flog.writeline " ---------------------------------------------------------------------------------------------------"
    Flog.writeline "Error generando gti_proc_emp. La informacion del horario teorico en el tablero no estará disponible."
    Flog.writeline "Error: " & Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ---------------------------------------------------------------------------------------------------"
    Flog.writeline "***"
    Flog.writeline
End Sub


Public Sub LimpiarJustificaciones()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que limpia el array de justificaciones.
' Autor      : FGZ
' Fecha      : 18/04/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer

For i = 1 To 10
    Arr_Justificaciones(i).Cantidad = 0
    Arr_Justificaciones(i).Ent = ""
    Arr_Justificaciones(i).Sal = ""
Next i
Indice_Justif = 1
Total_Hs_Justificadas = 0
End Sub

Public Sub Inicializar_Globales()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que Carga los array globales.
' Autor      : FGZ
' Fecha      : 17/05/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    'Politicas de alcance global
    Call Cargar_PoliticasGlobales
    Call Cargar_DetallePoliticas
    
    'FGZ - 15/06/2011
    Call ParametrosGlobales
    
End Sub





