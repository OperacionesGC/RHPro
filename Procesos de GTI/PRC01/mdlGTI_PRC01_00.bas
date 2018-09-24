Attribute VB_Name = "MdlGTI_PRC01_00"

Option Explicit

'Const Version = 1.01    'Inicial con nro de version
'Const FechaVersion = "11/10/2005"

'Const Version = 1.02    'Politica 820. Ctas Ctes
'Const FechaVersion = "12/10/2005"

'Const Version = 1.03    'Desgloses Produccion y Ausentismo
'Const FechaVersion = "18/10/2005"

'Const Version = 1.04    'Politica 800
'Const FechaVersion = "06/12/2005"

'Const Version = 1.05    'Mas detalle de Log en la compensacion de horas
'Const FechaVersion = "09/12/2005"

'Const Version = 1.06    'Controla que el periodo de gti no este cerrado
'Const FechaVersion = "19/12/2005"

'Const Version = 1.07    'Logs en politica 820
'Const FechaVersion = "22/12/2005"

'Const Version = 1.08    'politica 820
'Const FechaVersion = "23/12/2005"

'Const Version = 1.09                'politica 820 - Rehice el procedimiento Politica820_Actualiza_Detalles_Estandar
'Const FechaVersion = "19/01/2006"   'Porque no borraba los detalles de los thnro que ya no forman parte de la configuracion

'Const Version = 1.10                 'Maxi Conversion de horas se toco el AD05 y AD02 y AD01 para tome bien la conversion
'Const FechaVersion = "22/03/2006"   'porque solo andaba si el rango era 0

'Const Version = 2.01                'Nuevas Conversiones para Divino SA
'Const FechaVersion = "14/02/2006"  '

'Const Version = 2.02
'Const FechaVersion = "06/04/2006"
'Modificaciones:
    'Desglose_Jornada_Productiva_Nuevo.
        'Cuando Busco los tipos de estructuras a Desglozar faltaba el campo del Order en el Selct
        'StrSql = "SELECT confval, confnrocol FROM confrep WHERE repnro = 53 "
    'Modulo de conversiones: no andaba muy bien, antes de autorizar no andaba y despues andaba + o -
        'Modificaciones
        '   sub ad_02 y sub ad_05 para antes de autorizar
        '   sub ad_06 para despues de autorizar
        '   sub crearAD

'Const Version = 2.03
'Const FechaVersion = "06/04/2006"
''Modificaciones:
''       CAT - 17/05/2006 Conversion DIVINO SA

'Const Version = 2.04
'Const FechaVersion = "23/08/2006"
'    'FGZ - Modulo de conversiones: despues de autorizar
'        'Modificaciones
'        '   sub ad_06 para despues de autorizar
'        '   sub crearAD

'Const Version = 2.05
'Const FechaVersion = "22/09/2006"
''                   FGZ - Modulo de conversiones: despues de autorizar
''                   Modificaciones
''                       sub ad_07 Nueva conversion para Moño Azul (ConversionProd)

'Const Version = 2.06
'Const FechaVersion = "27/11/2006"
''                   FGZ - Modulo de conversiones: despues de autorizar
''                   Modificaciones
''                       sub ad_07 Nuevas conversion para AGD (FeriadosAGD)

'Const Version = 2.07
'Const FechaVersion = "03/01/2007"
''                   FGZ
''                   Modificaciones
''                       sub Compensar_Horas(): no estaba actualizando bien las horas compensadas parcialmente.

'Const Version = 2.08
'Const FechaVersion = "01/02/2007"
''                   FGZ
''                   Modificaciones
'''                       sub ad_07 Nuevas conversion para GORINA (HORASDESTAJO, ADICALMUERZO, PEFICIENCIA)

'Const Version = 2.09
'Const FechaVersion = "09/03/2007"
''                   FGZ
''                   Modificaciones
''                       Function Feriado del modulo de clase Feriado. Estaba calculando mal la variable Feriado_Por_Estructura cuando el feriado es nacional.

'Const Version = "2.10"
'Const FechaVersion = "18/04/2007"
''                   FGZ
''                   Modificaciones
''                       Agregado de log en modulo de conversiones sun ad_07
''                       Modulo de conversiones: Antes de Autorizar
''                           sub sub ad_05 para antes de autorizar

'Const Version = "2.11"
'Const FechaVersion = "30/04/2007"
'                   FGZ
'                   Modificaciones
'                       genera los segun parametro
'                       Modulo de conversiones: Antes de Autorizar y Despues de Autorizar

'Const Version = "2.12"
'Const FechaVersion = "28/05/2007"
''                   FAF
''                   Modificaciones
''                       En el modulo FechasHoras
''                       Se modificaron las funciones Convertir_A_Hora y Redondeo_Horas_Tipo para que soporten mas de 99 horas.

'Const Version = "3.00"
'Const FechaVersion = "01/06/2007"
''Modificaciones: FGZ
''      Mejoras generales de performance.

'Const Version = "3.01"
'Const FechaVersion = "12/06/2007"
''Modificaciones: FGZ
'''    Modulo Conversiones:    Ad_03, AD_05 y AD_06 Definian algunas variables globales lo que causaba problemas en las conversiones

'Const Version = "3.02"
'Const FechaVersion = "27/06/2007"
''Modificaciones: FGZ
'''    Modulo Conversiones:    Ad_07 problemas con las conversiones de feriados de AGD
'''    Modulo de clase    :    BuscarTurno: faltaban inicializar algunas variables

'Const Version = "3.03"
'Const FechaVersion = "08/08/2007"
''Modificaciones: FGZ
'''    Politica 890:    Se le agregó detalle de log
''     Problemas con el alcance por estructura de las politicas en el 1er dia del rango

'Const Version = "3.04"
'Const FechaVersion = "02/10/2007"
'Modificaciones: FGZ
'       Se modifico la funcion Feriado para que busque en todas las estructuras asignadas en la pol. de alcance para GTI ya que buscaba solo en la primera. CAS-04896

'Const Version = "3.05"
'Const FechaVersion = "21/11/2007"
''Modificaciones: Diego Rosso
''       Se agrego la Politica 571. Se agrego una nueva Conversion denominada Completar para Scheneider
''       Utiliza el Parametro 17. Lista de valores. donde los valores de la lista se corresponden a un dia de la semana

'Const Version = "3.06"
'Const FechaVersion = "15/01/2008"
'Modificaciones: FGZ
'       Se agrego la Politica 891 - Ajustes de AD.


'Const Version = "3.07"
'Const FechaVersion = "25/01/2008"
''Modificaciones: Diego Rosso - Se creo la politica 541 para Multivoice. Caso 5431.
''                              Se creo una nueva conversion denominada SABADODOMINGO MV tambien para Multivoice.
''                              Configuracion politica 541: Parametros:
''                              Version Politica, Tipo de hora: numero de tipo de hora de convenio que es donde se acumularan.
''                              Lista de valores: donde se pondran separados por coma los tipos de horas que acumulan en Convenio.

'Const Version = "3.08"
'Const FechaVersion = "13/02/2008"
'Modificaciones: FGZ
'          Politica 571. Se hicieron algunas mdificaciones a la politica 571(se agregó un nuevo parametro)
'          Se modificó la Conversion denominada "Completar" para Scheneider.

'Const Version = "3.09"
'Const FechaVersion = "15/02/2008"
'Modificaciones: Diego Rosso
'          Politica 541. Se cambio + por - en esta linea RestoHsOriginales = RestoHsOriginales + SumaA100
'          Conversion SABADODOMINGO MV= se cambio el numero del tipo de estructura convenio de 19 a 55.

'Const Version = "3.10"
'Const FechaVersion = "19/03/2008"
'Modificaciones: Diego Rosso - Se guarda version anterior la politica 541 como  politica 541_old
'                              Se crea una nueva version de la Politica 541
'                              Configuracion politica 541: Parametros:
'                              Version Politica, Tipo de hora: numero de tipo de hora de convenio que es donde se acumularan.
'                              Lista de valores: donde se pondran separados por coma los tipos de horas que acumulan en Convenio.

'Const Version = "3.11"
'Const FechaVersion = "28/03/2008"
''Modificaciones:
''                Diego Rosso - Se cambio 52 por 45 (tipo de hora) cuando busca las horas adicionales.
''                                               ya que en MV tienen 45 como hora adicional
''                                    Cuando genera horas al 50 o al 100 borra las adicionales.

'Const Version = "3.12"
'Const FechaVersion = "08/04/2008"
''Modificaciones:
''                FGZ - Politica 541: Se corrigió problema cuando busca hs adicionales no autorizada los sabados.

'Const Version = "3.13"
'Const FechaVersion = "16/04/2008"
''Modificaciones: FGZ
''   Modulo politicas: sub Cargar_DetallePoliticas.
''           Se cambió en el where el <> '' por IS NOT NULL

'Const Version = "3.14"
'Const FechaVersion = "18/04/2008"
''Modificaciones: FGZ
''   Conversiones: sub AD_07 programa "SABADODOMINGO MV".
''        ahora los dias
''           sabado hasta las 13 se convierten en %50 NA
''           sabado despues de las 13 se convierten en %100 NA
''           Domingos todo el dia se convierten en %100 NA
'''  Politica 541: Se corrigió problema cuando busca hs adicionales no autorizada los sabados.

'Const Version = "3.15"
'Const FechaVersion = "22/04/2008"
''Modificaciones: FGZ
''   Conversiones: sub AD_07 programa "SABADODOMINGO MV".
''        se volvió todo a como estaba antes
''           sabado hasta las 13 No se convierten
''           sabado despues de las 13 se convierten en %100 NA
''           Domingos todo el dia se convierten en %100 NA
'''  Politica 541: Se corrigió problema cuando busca hs adicionales no autorizada los sabados.

'Const Version = "3.16"
'Const FechaVersion = "26/06/2008"
''Modificaciones: FGZ
''       Mensages de log


'Const Version = "3.17"
'Const FechaVersion = "10/07/2008"
''Modificaciones: FGZ
''       Se agregaron controles a los subs
''               Desglose_Jornada_Productiva_Nuevo
''               Desglose_Jornada_Ausentismo_Nuevo
''   Conversiones: Se modificaron los controles de dia feriado, laborale y no laborable


'Const Version = "3.19"
'Const FechaVersion = "24/07/2008"
''Modificaciones: FGZ
''   Conversiones: Se modificó el sub AD_01

'Const Version = "3.20"
'Const FechaVersion = "06/08/2008"
''Modificaciones: FGZ
''   Conversiones: Se modificó el sub autoriza para que inserte anormalidades cuando no está autorizado
''                   Esto lo tocó CAT para CARGILL

'Const Version = "3.21"
'Const FechaVersion = "18/09/2008"
''Modificaciones: FGZ
''   Conversiones: Se modificó el sub BUSCAR_TURNO_NUEVO
''   Conversiones: Se modificó el sub autoriza para que inserte anormalidades cuando no está autorizado
''                   Esto lo tocó CAT para CARGILL. ademas se saca la anormalidad cuando se encuentra una autorizacion.

'Const Version = "3.22"
'Const FechaVersion = "24/09/2008"
'Modificaciones: FGZ
'   Conversiones: Se modificó el sub BUSCAR_TURNO_NUEVO
'   Conversiones: agregados de log


'Const Version = "3.23"
'Const FechaVersion = "14/11/2008"
''Modificaciones: FGZ
''   Conversiones: Se modificó la Politica 541 Topeo Semanal de hs extras (Custom MultiVoice)


'Const Version = "3.24"
'Const FechaVersion = "16/12/2008"
''Modificaciones: FGZ
''   Conversiones:
''       Se agregaron controles sobre los subs AD_01, AD_02, AD_05, AD_06
''       Nueva conversion TopeMinimo para Trilenium para aplicar en el estandar
''       Nueva conversion TopeMinimo_LV para Trilenium para aplicar en el estandar
''       Nueva conversion TopeMinimo_SD para Trilenium para aplicar en el estandar

'Const Version = "3.25"
'Const FechaVersion = "21/01/2009"
''Modificaciones: FGZ
''   Encriptacion de string de conexion

'Const Version = "3.26"
'Const FechaVersion = "16/02/2009"
''Modificaciones: FGZ
''   Encriptacion de string de conexion


'Const Version = "3.27"
'Const FechaVersion = "25/02/2009"
''Modificaciones: FGZ
''   Conversiones: Se modificó la Politica 541 Topeo Semanal de hs extras (Custom MultiVoice)
''                 le agregaue el control de negativos ---------


'Const Version = "3.28"
'Const FechaVersion = "04/03/2009"
''Modificaciones: FGZ
''   Conversiones: Se modificó la Politica 541 Topeo Semanal de hs extras (Custom MultiVoice)
''                 le agregaue el control de negativos ---------


'Const Version = "3.29"
'Const FechaVersion = "07/05/2009"
''Modificaciones: FGZ
''       Se agrego la Politica 576. Minimo de Hs extras sin autorizar


'Const Version = "3.30"
'Const FechaVersion = "02/06/2009"
''Modificaciones: FGZ
''       Se agrego la Politica 601. Control de Hs Nocturnas

'Const Version = "3.31"
'Const FechaVersion = "07/08/2009"
''Modificaciones: FGZ
''   No se hizo ninguna modificacion al proceso en si, se regeneró el ejecutable por en modulo global
''
''   Se agregó un procedimiento general en el modulo global para hacer chequeos de version
''   Este sub se debe invocar luego de crear el archivo de log y antes de comenzar con la logica especifica de cada proceso
''   Si el proceso no pasase el chequeo de version el proceso finalizará en estado "Error de Version"

'Const Version = "3.32"
'Const FechaVersion = "07/09/2009"
''Modificaciones: FGZ
''   Conversiones: sub AD_07 programa "VALES_SAT".
''       Se agregó el nuevo programa de conversion Custom para TELEARTE.


'Const Version = "3.33"
'Const FechaVersion = "18/09/2009"
''Modificaciones: FGZ
''   Conversiones: sub AD_07 programa "TURNO_PLUS".
''       Se agregó el nuevo programa de conversion Custom para TELEARTE.

'Const Version = "3.34"
'Const FechaVersion = "15/10/2009"
''Modificaciones: FGZ
''   Conversiones: sub AD_07 programa "CONVERSION".
''       Se agregó control de division por cero.

'--------------------------------------------------
'Const Version = "4.00"
'Const FechaVersion = "18/11/2009"
''Modificaciones: FGZ
''    Cambios Importantes
''        ALTER table gti_horcumplido add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_hishc add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_acumdiario add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_hisad add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_achdiario add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_his_achdiario add horas varchar(10) null default ‘0:00’
''        ALTER table Gti_det add horas varchar(10) null default ‘0:00’
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
''    Politica 601: Control de Horas Nocturnos.
''                   Se le agregó la posibilidad de configurar mas de 1 tipo de hora a controlar (no solo nocturnas)


'Const Version = "4.02"
'Const FechaVersion = "11/01/2010"
''Modificaciones: FGZ
''   Conversiones: Se modificó la Politica 541 Topeo Semanal de hs extras (Custom MultiVoice)


'Const Version = "4.03"
'Const FechaVersion = "26/01/2010"
''Modificaciones: FGZ
''   Conversiones: Se modificó la Politica 541 Topeo Semanal de hs extras (Custom MultiVoice)

'Const Version = "4.04"
'Const FechaVersion = "08/03/2010"
''Modificaciones: FGZ
''   Conversiones: Se modificó el sub Prog_23_Turno_Plus correspondiente al programa de conversion Turno_plus

'----------------------------
'Const Version = "5.00"
'Const FechaVersion = "15/06/2010"
''Modificaciones: FGZ
''    Control por entradas fuera de termino.
''           Antes
''               cuando se queria procesar algo en una fecha que caia en un periodo cerrado no se procesaba.
''           Ahora
''               Se puede reprocear un periodo cerrado solo cuando se aprueba una entrada fuera de termino.
''               Para ese reprocesamiento solo se tendrá en cuenta todas las entradas fuera de termino aprobadas.
'
''   Ademas se agregaron mejoras en los detalles de log y se optimizó algo el modulo de conversiones que estaba llamando muchas veces a la politica 810

'Const Version = "5.01"
'Const FechaVersion = "21/07/2010"
''Modificaciones: FGZ
''   Conversiones: Se modificó el sub AD_02 bug en update de campo horas

'Const Version = "5.02"
'Const FechaVersion = "10/08/2010"
''Modificaciones: FGZ
''   Conversiones: Se modificó el sub AD_06 bug en update de campo horas

'Const Version = "5.03"
'Const FechaVersion = "10/09/2010"
''Modificaciones: FGZ
''   Conversiones:
''       Nueva conversion Feriados_MV para Multivoice para aplicar en el estandar

'Const Version = "5.04"
'Const FechaVersion = "02/10/2010"
''Modificaciones: FGZ
''    Politica 577: Horas autorizables sin control.
''                   Se agregó la politica que filtra las horas que son autorizables pero sobre
''                   las cuales no se debe hacer el control de anormalidad ni de autorizacion

'Const Version = "5.05"
'Const FechaVersion = "08/10/2010"
''Modificaciones: FGZ
''    Politica 589: Desglose de Movilidad (Relevos).
''                   Se agregó la politica que genera el desglose de horas por relevo o movilidad


'Const Version = "5.06"
'Const FechaVersion = "12/11/2010"
''Modificaciones: FGZ
''    Politica 710: Distribucion de Horas.
''                   Se agregó la politica que genera y valida la distribucion de horas cargados


'Const Version = "5.07"
'Const FechaVersion = "16/12/2010"
''Modificaciones: Margiotta, Emanuel
''                   Se agrego al desglose de Horas por relevo el detalle de horas relevadas

'Const Version = "5.08"
'Const FechaVersion = "05/04/2011"
''Modificaciones: Margiotta, Emanuel - FGZ
''                   desglose de Horas por relevo el detalle de horas relevadas


'Const Version = "5.09"
'Const FechaVersion = "12/04/2011"
''Modificaciones: Margiotta, Emanuel
''                   Se agregó la política 300 de Infracciones de Sanciones

'Const Version = "5.10"
'Const FechaVersion = "02/05/2011"
''Modificaciones: FGZ
''                   Se agregó/activó la política 465 de francos compensatorios

'Const Version = "5.21"
'Const FechaVersion = "21/06/2011"
''Modificaciones: FGZ
''    Se agregaó el control de firmas a las novedades horarias
''       Se modifico:
''           Buscar_Turno
''           Buscar_Turno_nuevo
''           Politica 480
''           Politica 490

'Const Version = "5.22"
'Const FechaVersion = "25/07/2011"
''Modificaciones: FGZ
''   Conversiones: sub AD_07 programa "NOC_Sykes_CR":
''           Nueva conversion NOC_Sykes_CR para Sykes

'Const Version = "5.23"
'Const FechaVersion = "02/09/2011"
''Modificaciones: FGZ
''    Politica 710: Distribucion de Horas.
''                   habia un campo mal referenciado en his_ot

'Const Version = "5.24"
'Const FechaVersion = "08/09/2011"
''Modificaciones: FGZ
''    Politica 710: Distribucion de Horas.
''       Se le agregó un control de nulo

'Const Version = "5.25"
'Const FechaVersion = "29/09/2011"
''Modificaciones: FGZ
''    Politica 710: Distribucion de Horas. Se restauró la version original de AGD

'Const Version = "5.26"
'Const FechaVersion = "25/10/2011"
''Modificaciones: FGZ
''   Conversiones: sub AD_07 programa "DESC_LUNCH":
''           Nueva conversion DESC_LUNCH para Sykes


'Const Version = "5.27"
'Const FechaVersion = "15/11/2011"
''Modificaciones: FGZ
''  Conversiones: sub AD_07 programa "SABADOS":
''           Nueva conversion SABADOS para Gral Mills


'Const Version = "5.28"
'Const FechaVersion = "31/01/2012"
'''Modificaciones: HJI
'''  Conversiones: sub AD_07 programa "NOC_Sykes_CR":
'''           Se agrego decuento de lunch para horas nocturnas


'Const Version = "5.29"
'Const FechaVersion = "01/10/2012"
''Modificaciones: Margiotta, Emanuel
''  Conversiones: Se mofico las funciones de conversion de horas (AD_02,AD_05) antes de autorizar en una sola función "AD_Antes_Autorizar"

'Const Version = "5.30"
'Const FechaVersion = "18/06/2012"
'Modificaciones: EAM (14825)
'  Compensación de Hs: sub PRC01 : Se agregó la compensación de hs a partir de un parte diario. Este tiene alcance al igual que las politicas (individual,estructura, global)
'           - Se Agrego un nuevo modulo (mdlCompensacion) en el cual se agregaron las funciones existentes de compensación y las nuevas.
'           - Se corrigio en la función vieja de compensación "Compensar_Horas()" los .open por los OpenRecordset


'Const Version = "5.32"
'Const FechaVersion = "25/10/2012"
''Modificaciones: Margiotta, Emanuel
''  Conversiones: Se fusiono la version para Pulsmar que convierte por día con el estandar.
''               Se mejoro la performance.
''               Conversiones antes de autorizar
''   Politica 500: Se corrigio en compensación de horas, en la version 1 (turno) y 2 (parte diario) que cuando una hora compensa mas del 100% se convierta a la unidad original, sino se geran horas irreales.
''   Conversión de horas: Se agrego para que convierta por franja y por acumulado. Por franja convierte cada hora y la va acumulando. Acumulado suma los acumulados y los convierte el resultado de la suma.


'Const Version = "5.33"
'Const FechaVersion = "25/10/2012"
''Modificaciones: Margiotta, Emanuel (14825)
'' Se creo una nueva versión pero esta resulto ya en la 5.32

'Const Version = "5.34"
'Const FechaVersion = "06/12/2012"
'Modificaciones: Margiotta, Emanuel (16553)
'    Conversiones: Prog_Conv_Feriado_Sa_DO - CUSTOM AMR(Asociación Médica de Rosario)
'    Los empleados que tiene el turno full-time (tipo de estructura 21 y estructura 1245) se consideran turnos full-time.
'    Si trabajan un sábado o domingo en un día feriado se le debe abonar medio feriado sino aplica conversión normal.

'Const Version = "5.35"
'Const FechaVersion = "14/05/2013"
'Modificaciones: Deluchi Ezequiel
' FGZ:  sub PRC_01(). Le agregué un control para impresion de log (solo cuando debug es true)
'                   Flog.writeline Espacios(Tabulador * 1) & "  Tiempo DIFERENCIA : " & (GetTickCount - tiempoPrueba)
' LED: Versionado politica 589 - Se agrego version 2 - CAS-18225

'Const Version = "5.36"
'Const FechaVersion = "01/08/2013"
'Modificaciones: Deluchi Ezequiel
'CAS-11450 - H&A - Adecuación Conversiones - Nivelacion del modulo conversionesV3 R3 con R4.

'Const Version = "5.37"
'Const FechaVersion = "18/12/2013"
''Modificaciones: Margiotta, Emanuel
''CAS-22808 - SGS - Distribución Contable
''Politica 589 - Se genero la versión 3 de la poltica que distribuye a partir del Horario Cumplido (tabla -> gti_desgloce_hc).


'Const Version = "5.38"
'Const FechaVersion = "11/03/2014"
''Modificaciones: Margiotta, Emanuel
''CAS-22808 - SGS - Distribución Contable
''Politica 589 - Se genero la versión 3 de la poltica que distribuye a partir del Horario Cumplido (tabla -> gti_desgloce_hc).
''Politica 589 - Se genero la versión 4 de la poltica que distribuye a partir del Horario Cumplido (tabla -> gti_desgloce_hc). Esta version distribuye hs x por hs generadas en el HC.
''               Formula que aplica (por cada reg HC /sum(HC)* AD
'' Se agrego nuevo Programa de Conversión: DESC_CITA_MED
'

'Const Version = "5.39"
'Const FechaVersion = "30/05/2014"
''Modificaciones: FGZ
''       CAS-24229 - Lucaioli -  Nueva Compensación
''       Se agrego nuevo Programa de Conversión: DESTINO

'Const Version = "5.40"
'Const FechaVersion = "12/08/2014"
''Modificaciones: FGZ
''       CAS-21778 - Sykes El Salvador - Bug PRC030 Duplicidad de novedades
''           Se corrigió problema con tabulacion de los mensajes de log


'Const Version = "5.41"
'Const FechaVersion = "15/09/2014"
''Modificaciones: FGZ
''       CAS-21778 -  Sykes El Salvador - Bug en conversión de Hs GT
''           Se corrigió problema en conversiones (antes y despues de autorizar) para feriados
''           Ademas se corrigió problema cuando hay varias conversiones sobre los mismos tipos de horas (origen y destino) y con alcance por estructura


'Const Version = "5.42"
'Const FechaVersion = "08/10/2014"
''Modificaciones: FGZ
''       CAS-26755 - MEDICUS - CUSTOM TOPEO DE HORAS HORAS EXTRAS
''    Nueva Politica 579: Topeo Diario de Horas
''                   Se agregó la politica que topea diariamente todos los tipos de horas configurados
'
''Ademas
''       CAS-25418 - MOÑO AZUL - Error en Conversión
''       Se modificó el programa de conversion "Conversion" del cliente para redondear a 1 cuando la cantidad de hs es menor .


'Const Version = "5.43"
'Const FechaVersion = "21/04/2014"
'Modificaciones: EAM
'       CAS-28352 - Salto Grande - Custom GTI - Franco Compensatorio
'    Nueva Politica 465: Version 3 de la politica de franco compensatorio
'           Se hizo la version 3 de la politica de Franco Compensatorio. Permite configurar los días que se quieren analizar, los tipos de horas que se contabilizan
'           ,el factor de multiplicacion y el tipo de período que se quiere analizar
'


'Const Version = "5.44"
'Const FechaVersion = "26/05/2015"
'Modificacion: Fernandez, Matias -CAS-30614 - MONASTERIO - Bug licencias en día franco -se contempla la conversion de horas cuando es 0.

'Const Version = "5.45"
'Const FechaVersion = "30/09/2015"
'Modificacion: Fernandez, Matias -CAS- 33203 - ILE - Error en conversion de Horas- error en conversion de horas.
                                                     ' en CrearAD acumula debe ser igual a -1

Const Version = "5.46"
Const FechaVersion = "07/10/2015"
'Modificacion: Fernandez, Matias -CAS- 33203 - ILE - se revirtio el cambio hecho en la version 5.45 y se corrige en la funcion
                                  ' se revirtio tambien lo del caso 5.44, ahora se corrige lo solicitado por monasterio con configuracion
                                  ' si bien ambos casos andaban corregian el problema del cliente, estos traian problemas cuando tenes mas de una conversion por dia
                                  '(efecto colateral) entonces se revierten ambos casos.
'-------------------------------------------------------
'VERSION NO LIBERADA
'
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Dim CEmpleadosAProc As Integer
Dim CDiasAProc As Integer
Dim IncPorc As Single
Dim Progreso As Single
Dim IncPorcEmpleado As Single
Dim HuboErrores As Boolean
Dim ProgresoEmpleado As Single

'FGZ - 20/09/2004
Global objBTurno As New BuscarTurno
Global objBDia As New BuscarDia
Global objFeriado As New Feriado
'Global objFechasHoras As New FechasHoras


Public Sub Main()
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Fecha As Date
Dim Legajo As Long
Dim pos1 As Byte
Dim pos2 As Byte
Dim strcmdLine As String
Dim Ternro As Long
Dim fs
Dim objRs As New ADODB.Recordset
Dim myrs As New ADODB.Recordset
Dim objAD_Dup As New ADODB.Recordset
Dim TEmp As Integer
Dim rsEmpleado As New ADODB.Recordset
Dim Cant As Long
Dim PeriodoCerrado As Boolean
Dim ListaPar

Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim rs_Per As New ADODB.Recordset

Dim PID As String
Dim ArrParametros

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

    HuboErrores = False
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    ' inicializa los nombres de las tablas temporales segun la DB
    Call CargarNombresTablasTemporales
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(PathFLog & "PRC01" & "-" & NroProceso & ".log", True)
    
    Cantidad_de_OpenRecordset = 0
    Cantidad_Call_Politicas = 0
    
    'Abro la conexion
    On Error Resume Next
    
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    
    Nivel_Tab_Log = 0
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------------------"
    Flog.writeline
    
    'FGZ - 05/08/2009 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 2, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans

        Flog.writeline
        GoTo Final
    End If
    'FGZ - 05/08/2009 --------- Control de versiones ------
    
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords

    Set objFechasHoras.Conexion = objConn
        

    StrSql = "SELECT IdUser,bprcfecdesde,bprcfechasta, bprcparam FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Parametros: "
        Flog.writeline Espacios(Tabulador * 1) & "Usuario: " & objRs!IdUser
        'IdUser = objRs!IdUser
        Flog.writeline Espacios(Tabulador * 1) & "Desde: " & objRs!bprcfecdesde
        FechaDesde = objRs!bprcfecdesde
        Flog.writeline Espacios(Tabulador * 1) & "Hasta: " & objRs!bprcfechasta
        FechaHasta = objRs!bprcfechasta
        Flog.writeline Espacios(Tabulador * 1) & "bprcparam: " & objRs!bprcparam
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
            Else
                depurar = False
                ReprocesarFT = False
            End If
        Else
            depurar = False
        End If
        FechaDesde = objRs!bprcfecdesde
        FechaHasta = objRs!bprcfechasta
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
    
    Flog.writeline "-------------------------------------------------------------"
    Flog.writeline "-------------------------------------------------------------"
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "Inicio :" & Now
    
    StrSql = "SELECT empleado.ternro, empleado.empleg FROM batch_empleado "
    StrSql = StrSql & " INNER JOIN empleado ON batch_empleado.ternro = empleado.Ternro "
    StrSql = StrSql & " WHERE batch_empleado.bpronro = " & NroProceso
    If myrs.State = adStateOpen Then myrs.Close
    OpenRecordset StrSql, myrs
    
    Fecha = FechaDesde
    CDiasAProc = DateDiff("d", FechaDesde, FechaHasta) + 1
    'FGZ - 20/02/2004
    If Not myrs.EOF Then
        CEmpleadosAProc = myrs.RecordCount
        IncPorc = ((100 / CEmpleadosAProc) * (100 / CDiasAProc)) / 100
        IncPorcEmpleado = (100 / CDiasAProc)
    End If
    
    'Activo el manejador de errores
    On Error GoTo ce
    
    Progreso = 0
    
    'FGZ - Mejoras ----------
    Call Inicializar_Globales
    'FGZ - Mejoras ----------
    
    Nivel_Tab_Log = 1
    Do While Not myrs.EOF
        'MyBeginTrans

        Ternro = myrs!Ternro
        Fecha = FechaDesde
        
        Empleado.Ternro = Ternro
        Empleado.Legajo = myrs!empleg
        
        If depurar Then
            Flog.writeline ""
            Flog.writeline "-------------------------------------------------------------------------------------"
            'Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Inicio Empleado:" & Ternro & " " & Now
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Inicio Empleado:" & Empleado.Legajo & " " & Now
        End If
        
        'FGZ - Mejoras ----------
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cargando politicas de alcance individual... "
        End If
        Call Cargar_PoliticasIndividuales
        'FGZ - Mejoras ----------
        
        UsaConversionHoras = False
        Call Politica(570)
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Usa Conversion de Horas:" & UsaConversionHoras
        End If
        
        Usa_Conv = False
        'politica que indica si usa la conversion de horas
        Call Politica(810)
        
        Nivel_Tab_Log = 2
        ProgresoEmpleado = 0
        Do While Fecha <= FechaHasta
            ReDim arrTempHsAD(0) As THsAD
            Nivel_Tab_Log = 2
            If depurar Then
                Flog.writeline
                Flog.writeline "-----------------------"
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Procesando fecha: " & Fecha
            End If
            
            'FGZ - 27/06/2007 --------------
            Call BlanquearVariables
            Call CreateTempTable(TTempWFInputFT)
            
            'FGZ - 27/06/2007 --------------
            
            'FGZ - Mejoras ----------
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cargando politicas de alcance por estructura... "
            End If
            Call Cargar_PoliticasEstructuras(Fecha)
            'FGZ - Mejoras ----------
            
            'Reviso que el periodo de GTI al que pertenece la fecha no se encuentre cerrado
            StrSql = "SELECT pgtiestado FROM gti_per "
            StrSql = StrSql & " WHERE pgtidesde <= " & ConvFecha(Fecha)
            StrSql = StrSql & " AND pgtihasta >= " & ConvFecha(Fecha)
            If rs_Per.State = adStateOpen Then rs_Per.Close
            OpenRecordset StrSql, rs_Per
            PeriodoCerrado = False
            Do While Not rs_Per.EOF
                If Not rs_Per!pgtiestado Then
                    PeriodoCerrado = True
                    Nivel_Tab_Log = 3
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se va a procesar, Periodo cerrado"
                    End If
                End If
                rs_Per.MoveNext
            Loop
            
            
            'If Not PeriodoCerrado Then
            If Not PeriodoCerrado Or (PeriodoCerrado And ReprocesarFT) Then
                If (PeriodoCerrado And ReprocesarFT) Then
                    Flog.writeline Espacios(Tabulador * 2) & "Se reprocesará teniendo en cuenta todas las entrada fueras de termino aprobadas."
                End If
            
                Call PRC_01(True, Ternro, Fecha, Cant, arrTempHsAD)
                If depurar Then
                    Flog.writeline
                End If
                
                Nivel_Tab_Log = 2
                usaHorasExtras = False
                'Politica de que permite el uso del sub de autorizaciones
                Call Politica(575)
                If usaHorasExtras Then
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Usa Autorizacion de Horas Extras "
                    End If
                    'Autorizo horas extras
                    Nivel_Tab_Log = 3
                    
                    'FGZ - 07/05/2009 ---- le agregué la llamada a esta politica ------
                    HayMinimoExtrasSinAutorizar = False
                    Call Politica(576)
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Minimo de extras sin autorizar: " & HayMinimoExtrasSinAutorizar
                    End If
                    'FGZ - 07/05/2009 ---- le agregué la llamada a esta politica ------
                    Call Autoriza(Fecha, Ternro, TEmp)
                Else
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "NO Usa Autorizacion de Horas Extras "
                    End If
                End If
                    
                'FGZ - 29/05/2009 - se agregó la politica --------------------
                'Control de Hs Nocturnas no autorizadas
                Call Politica(601)
                'FGZ - 29/05/2009 - se agregó la politica --------------------
                
                Nivel_Tab_Log = 2
                If UsaConversionHoras Then
                    
                    ' LLamar gtiad03.p (Conversiones)
                    'FGZ - 07/05/2010 - esta politica no está implementada.
                    'Call Politica(830)
                    'FGZ - 07/05/2010 - esta politica no está implementada.
                    
                    'LLamar gtiad05.p (Conversiones)
                    Nivel_Tab_Log = 3
                    Call AD_06(Fecha, Ternro, depurar, arrTempHsAD())

                    Nivel_Tab_Log = 3
                    Call AD_07(Ternro, Fecha)
                        
                    'Call DepurarDesgloses(Ternro, Fecha)
                    'Call Desglose_Jornada_Productiva(Ternro, Fecha)
                    'FGZ 17/11/2005 - Nuevo procedimiento de desglose
                    'Antes este procedimiento hacia un desglose fijo en 3 niveles fijos
                    'Regimen Horario, Categoria y Producto
                    'Ahora se puede hacer un desglose de hasta 5 niveles y configurable por confrep
                    'confrep 53 columnas 50 a 54 inclusive
                    'ver doc de configuracion
                    Nivel_Tab_Log = 3
                    Call Desglose_Jornada_Productiva_Nuevo(Ternro, Fecha)
                    
                    'Call Desglose_Jornada_Ausentismo(Ternro, Fecha)
                    'FGZ 17/11/2005 - Nuevo procedimiento de desglose
                    'lo mismo pasa con este, en realidad son iguales salvo que busacan otros tipos de horas
                    Nivel_Tab_Log = 3
                    Call Desglose_Jornada_Ausentismo_Nuevo(Ternro, Fecha)
                End If
                Nivel_Tab_Log = 2
                
                
                
                'FGZ - 08/10/2014 - Topeo diario de Horas
                Call Politica(579)
                
                'Redondeo
                ' FGZ - 23/02/2004
                Call Politica(890)
                
                'FGZ - 08/10/2010 --------------
                UsaDesgloseMovilidad = False
                Call Politica(589)
                
                'LED - 14/05/2013 - Se paso dentro de la politica
                'If UsaDesgloseMovilidad Then
                    'Call Desglose_Relevos(Ternro, Fecha)
                'End If
                'LED - 14/05/2013 - Se paso dentro de la politica
                                
                'FGZ - 08/10/2010 --------------
                
                'FGZ - 12/11/2010 --------------
                'Distribucion de hs
                Call Politica(710)
                'FGZ - 12/11/2010 --------------
                
                
                'Diego Rosso
                Call Politica(541)
                
                'Ajustes
                ' FGZ - 08/01/2008
                Call Politica(891)
                
                'Ctas Corrientes
                ' FGZ - 12/03/2004
                Call Politica(820)
                
                'Actualizacion de cta cte de hs para francos compensatorios
                'FGZ - 20/04/2011
                Call Politica(465)
                
                'EAM- Infraccion de Sanciones - 11/04/2011
                Call Politica(300)
                
                'MyCommitTrans
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Periodo cerrado. Empleado:" & Ternro & " " & Fecha & " no se procesa"
                End If
            End If
            
            'actualizo tambien el porcentaje del empleado
            ProgresoEmpleado = ProgresoEmpleado + IncPorcEmpleado
            StrSql = "UPDATE batch_empleado SET progreso = " & ProgresoEmpleado & " WHERE bpronro = " & NroProceso & " AND ternro = " & Ternro
            objConn.Execute StrSql, , adExecuteNoRecords

            
            Progreso = Progreso + IncPorc
            'Debug.Print Progreso
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Call ActualizarFT(2, Fecha, Ternro)
            
            Fecha = DateAdd("d", 1, Fecha)

        Loop
        Nivel_Tab_Log = 1
        
            ' si el empleado se proceso por completo entonces lo borro de batch_empleados
            StrSql = "SELECT progreso FROM batch_empleado WHERE bpronro = " & NroProceso & " AND ternro = " & Ternro
            OpenRecordset StrSql, rsEmpleado
            If Not rsEmpleado.EOF Then
                If rsEmpleado!Progreso = 100 Then
                    ' lo puedo eliminar
                        StrSql = "DELETE FROM batch_empleado WHERE bpronro = " & NroProceso & " And Ternro = " & Ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
    
    
            If depurar Then
                Flog.writeline Espacios(Tabulador * 1) & "Fin Empleado:" & Ternro & " " & Now
            End If
            
SiguienteEmpleado:
        myrs.MoveNext
    Loop

    'Deshabilito el manejador de errores
    On Error GoTo 0

    'Habilito manejador gral
    On Error GoTo ME_Main
    
    StrSql = "DELETE FROM Batch_Procacum WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Actualizo el Btach_Proceso
    If Not HuboErrores Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
   
    ' -----------------------------------------------------------------------------------
    'FGZ - 22/09/2003
    'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
    If Not HuboErrores Then
        StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_Batch_Proceso

        StrSql = "INSERT INTO His_Batch_Proceso (bpronro,btprcnro,bprcfecha,iduser"
        StrSqlDatos = rs_Batch_Proceso!bpronro & "," & rs_Batch_Proceso!btprcnro & "," & _
                 ConvFecha(rs_Batch_Proceso!bprcfecha) & ",'" & rs_Batch_Proceso!IdUser & "'"
        
        If Not IsNull(rs_Batch_Proceso!bprchora) Then
            StrSql = StrSql & ",bprchora"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchora & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcempleados) Then
            StrSql = StrSql & ",bprcempleados"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcempleados & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecdesde) Then
            StrSql = StrSql & ",bprcfecdesde"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecdesde)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfechasta) Then
            StrSql = StrSql & ",bprcfechasta"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfechasta)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcestado) Then
            StrSql = StrSql & ",bprcestado"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcestado & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcparam) Then
            StrSql = StrSql & ",bprcparam"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcparam & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcprogreso) Then
            StrSql = StrSql & ",bprcprogreso"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcprogreso
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecfin) Then
            StrSql = StrSql & ",bprcfecfin"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecfin)
        End If
        If Not IsNull(rs_Batch_Proceso!bprchorafin) Then
            StrSql = StrSql & ",bprchorafin"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprchorafin & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprctiempo) Then
            StrSql = StrSql & ",bprctiempo"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprctiempo & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!Empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!Empnro
        End If
        If Not IsNull(rs_Batch_Proceso!bprcPid) Then
            StrSql = StrSql & ",bprcPid"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcPid
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecInicioEj) Then
            StrSql = StrSql & ",bprcfecInicioEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecInicioEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcfecFinEj) Then
            StrSql = StrSql & ",bprcfecFinEj"
            StrSqlDatos = StrSqlDatos & "," & ConvFecha(rs_Batch_Proceso!bprcfecFinEj)
        End If
        If Not IsNull(rs_Batch_Proceso!bprcUrgente) Then
            StrSql = StrSql & ",bprcUrgente"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!bprcUrgente
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraInicioEj) Then
            StrSql = StrSql & ",bprcHoraInicioEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraInicioEj & "'"
        End If
        If Not IsNull(rs_Batch_Proceso!bprcHoraFinEj) Then
            StrSql = StrSql & ",bprcHoraFinEj"
            StrSqlDatos = StrSqlDatos & ",'" & rs_Batch_Proceso!bprcHoraFinEj & "'"
        End If

        StrSql = StrSql & ") VALUES (" & StrSqlDatos & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
        'Reviso que haya copiado
        StrSql = "SELECT * FROM His_batch_proceso WHERE bpronro =" & NroProceso
        OpenRecordset StrSql, rs_His_Batch_Proceso
        
        If Not rs_His_Batch_Proceso.EOF Then
            ' Borro de Batch_proceso
            StrSql = "DELETE FROM Batch_Proceso WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
        If rs_Batch_Proceso.State = adStateOpen Then rs_Batch_Proceso.Close
        If rs_His_Batch_Proceso.State = adStateOpen Then rs_His_Batch_Proceso.Close
    End If
    ' FGZ - 22/09/2003
    ' -----------------------------------------------------------------------------------


Final:
    If depurar Then
        If CnTraza.State = adStateOpen Then CnTraza.Close
    End If
       
    If myrs.State = adStateOpen Then myrs.Close
    Set myrs = Nothing
    objConn.Close
    Set objConn = Nothing
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    If rs_Per.State = adStateOpen Then rs_Per.Close
    Set rs_Per = Nothing
    
    Flog.writeline
    Flog.writeline
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
    
ce:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & "Error. Empleado abortado " & " " & Fecha
    Flog.writeline Espacios(Tabulador * 0) & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline

    HuboErrores = True
    GoTo SiguienteEmpleado

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
        'objConnProgreso.Execute StrSql, , adExecuteNoRecords
        objConn.Execute StrSql, , adExecuteNoRecords
    MyCommitTrans
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

