Attribute VB_Name = "mdlPagoDto"
Option Explicit

'Version: 1.01  'Inicial
'1.02    'Dias Correspondientes Parciales con fecha desde
'1.03    '1.03    'Nueva version politica 1500. V7

'Const Version = 1.03    'Nueva version politica 1500. V7
'Const FechaVersion = "21/10/2005"

'Const Version = 1.04    'Version con otra conexion para el progreso
'Const FechaVersion = "23/11/2005"

'Const Version = 1.05    'Correccion en el calculo de descuentos, en la politica 1500 - 1
'Const FechaVersion = "28/11/2005"

'Const Version = 1.06    'Correccion en el calculo del pago/descuento, en la politica 1500 - 11
'Const FechaVersion = "29/11/2005"

'Const Version = 1.07    'politica 1500 - jornales programa 1
'Const FechaVersion = "30/11/2005"

'Const Version = 1.08    ' + LOG
'Const FechaVersion = "01/12/2005"

'Const Version = 2.01     'Revision general. Nuevo detalle politica 1500 (programa 12)
'Const FechaVersion = "12/12/2005"

'Const Version = 2.02     'mas log en la politica 1500 - v 12
'Const FechaVersion = "12/12/2005"

'Const Version = 2.03     'Detalle de politica por dias correspondientes. politica 1500 - v 12
'Const FechaVersion = "14/12/2005"

'Const Version = 2.04     'Detalle de politica por dias correspondientes. politica 1500 - v 8 (Paga y Dta por mes)
'Const FechaVersion = "25/01/2006"

'Const Version = 2.05    'politica 1503 - Ahora si hay error de config de parametros no aborta y se usa una config por default
'                        'politicas 1500 - Todas las versiones, cuando busca la forma de Liq, si no la encuentra entonces se asume mensual por default
'Const FechaVersion = "20/02/2006"

'Const Version = 2.06    'politicas 1500 - Todas las versiones, problemas cuando calcula la cantidad de dias en mes de febrero
'Const FechaVersion = "23/02/2006"

'Const Version = 2.07    'politicas 1500 - versiones 12 (por lic y por dias corr)
''FGZ - 27/02/2006 No estaba inicializando estas variables y por lo tanto
''      siempre comenzaba en la primera quincena (cuando Jornal)
''
'Const FechaVersion = "27/02/2006"

'Const Version = 2.08    'politicas 1500 - versiones 12 (por lic y por dias corr)
''FGZ - 27/02/2006 No estaba inicializando estas variables y por lo tanto
''      siempre comenzaba en la primera quincena (cuando Jornal)
'Const FechaVersion = "01/03/2006"

'Const Version = 2.09    'politicas 1500 - versiones 12 (por lic y por dias corr)
'Const FechaVersion = "08/03/2006"

'Const Version = 2.11
'Const FechaVersion = "07/04/2006"
        'politicas 1500 - versiones 12 (por lic y por dias corr)
        ' Seteaba mal el tipo de dia para los jornales

'Const Version = 2.12
'Const FechaVersion = "27/07/2006"
        'nuevo parametro opcional: pliqnro
        
'Const Version = 2.13
'Const FechaVersion = "07/08/2006"
        'Para determinar la fecha de alcance de politicas se debe tomar:
        ' 1   Si es por dias correspondientes: la fecha "a partir del"
        ' 2   Si el por licencias debe tomar la fecha desde de la licencia (de cada empleado)

'Const Version = 2.14
'Const FechaVersion = "10/08/2006"
        'Para determinar la fecha de alcance de politicas se debe tomar:
        ' 1   Si es por dias correspondientes: la fecha "a partir del"
        ' 2   Si el por licencias debe tomar la fecha desde de la licencia (de cada empleado). Se toque este caso. Tomaba mal la fecha!!

'Const Version = 2.15
'Const FechaVersion = "10/08/2006"
        'Se agrego la version 13 a la politica 1500. Es igual a la version 11 pero topea a 30 dias fijo

'Const Version = 2.16
'Const FechaVersion = "25/09/2006"
        'Se modifico la version 8 de la politica 1500 de pago/dtos para días correspondientes (Politica1500_V2PagaDescuenta_PorMes())
        'Se corrijio para CAS 2382

'Const Version = 2.17
'Const FechaVersion = "29/09/2006"
        'Se modifico la version 8 de la politica 1500 de pago/dtos para días correspondientes (Politica1500_V2PagaDescuenta_PorMes())
        'Se corrijio para CAS 2382. Ej 20/11/2006 al 03/12/2006, realizaba los 2 pagos en noviembre

'Const Version = 2.18
'Const FechaVersion = "20/12/2006"
        'Se corrigio un bug en la captura de errores del proceso: se quito un resume next

'Const Version = 2.19
'Const FechaVersion = "13/02/2007"
        'FAF - Se agrego la version 14 a la politica 1500. Paga adelantado y descuenta segun Grupo de Liquidacion.

'Const Version = 2.2
'Const FechaVersion = "20/02/2007"
        'Lisandro Moro - Se agrego la version 15 a la politica 1500 - CUSTOM AGD -. Paga todo y descuenta con tope a 30 dias segun el campo LiquidaVac.

'Const Version = 2.21
'Const FechaVersion = "17/05/2007"
        'Fernando Favre - Se agrego la version 16 a la politica 1500 - CUSTOM SMT -. Paga todo y descuenta con tope a 30 dias para Mensuales y descuenta por dias corridos para Jornales.

'Const Version = 2.22
'Const FechaVersion = "23/10/2007"
        'Lisandro Moro - Se agrego la version 17 a la politica 1500 - CUSTOM .ARLEI

'Const Version = 2.23
'Const FechaVersion = "31/10/2007"
        'Fernando Favre - Se modifico version 16 a la politica 1500 - CUSTOM SMT - Si el mes tenia 31 dias, al estar topeado a 30, sobraba un 1 al final y causaba problemas!!

'Const Version = 2.24
'Const FechaVersion = "13/11/2007"
        'Lisandoro Moro - Papelbril

'Const Version = 2.25
'Const FechaVersion = "26/12/2007"
'        'Fernando Favre - Se modifico la version 15 de la politica 1500 - CUSTOM AGD - Se topea a que los meses tengan maximo 30 dias.

'---------------------------------------------------------------
'Const Version = "2.26"
'Const FechaVersion = "28/01/2007" 'FGZ
'Se cambió la fecha para la cual se resuelve el alcance por estructura de las politicas (sub politica)
'               Se cambió el uso de fecha_desde en los querys por aux_fecha
'                If fecha_desde > Date Then
'                    Aux_fecha = fecha_desde
'                Else
'                    If fecha_hasta > Date Then
'                        Aux_fecha = Date
'                    Else
'                        Aux_fecha = fecha_hasta
'                    End If
'                End If
'---------------------------------------------------------------

'Const Version = 2.27
'Const FechaVersion = "22/04/2008"
       'Gustavo Ring - Se modifico la version 15 de la politica 1500 - CUSTOM AGD - No genera pago dto si no tiene liquida y borra pago dto anteriores si reprocesa.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Const Version = 2.28
'Const FechaVersion = "12/06/2008"
       'Gustavo Ring - Se modifico la version 15 de la politica 1500 - CUSTOM AGD - No genera pago dto si no tiene liquida y borra pago dto anteriores si reprocesa.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Const Version = 2.29
'Const FechaVersion = "16/07/2008"
       'Gustavo Ring - Se modifico la version 12 de la politica 1500 - Toma en cuenta los descuentos quincenales
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Const Version = 2.3
'Const FechaVersion = "04/09/2008"
       'Gustavo Ring - Se modifico la version 12 de la politica 1500 - Cuando se usan Licencias ahora no genera un dia de descuento demás
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Const Version = 2.31
'Const FechaVersion = "28/10/2008"
       'Gustavo Ring - Se modifico la version 12 de la politica 1500 - Se arreglo para que topee en 15 en los meses de 31 días
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Const Version = 2.32
'Const FechaVersion = "19/06/2009"
'       'Lisandro Moro - Se agrego la encriptacion al string de conexion.
'       '                Se agrego la generacion de situacio de revista al pago descuento, modelo 15 - AGD

'Const Version = "2.33"
'Const FechaVersion = "24/06/2009" ' FGZ
'       'Politca 1500 Version AGD: Se modificó la llamada a la poitica 1509.
'       'Politca 1509:  Se modificó el codigo de la politica


'Const Version = "2.34"
'Const FechaVersion = "02/10/2009" ' FGZ
'       'Politca 1500 Version AGD: Se modificó la llamada a la poitica 1509. La otra version (Dias correspondientes)

'Const Version = "2.35"
'Const FechaVersion = "23/10/2009" ' FGZ
'       'Nueva Politca 1512: Vencimiento de vacaciones
''           como consecuencia se modificaron todas las versiones que pagan y/o descuentan por dias correspondientes

'Const Version = "2.36"
'Const FechaVersion = "16/11/2009" 'FGZ
''           Problema con la funcion de validacion de version.


'Const Version = "2.37"
'Const FechaVersion = "14/12/2010" ' Lisandro Moro
'        'Politca 1509:  Correciones en la politica Politica1500_V2PagoDescuento_AGD

'------------------------------------------------------------------------------------
'Const Version = "3.00"
'Const FechaVersion = "14/04/2010" 'FGZ
''           Ahora los periodos de vacaciones ahora pueden tener alcance por estructura


'Const Version = "3.01"
'Const FechaVersion = "01/06/2010" 'FGZ
''           Integracion de versiones


'Const Version = "3.02"
'Const FechaVersion = "13/08/2010" 'FGZ
''       politicas 1500 - Nueva version: 19 Monresa (por lic y por dias corr)

'Const Version = "3.03"
'Const FechaVersion = "26/08/2010" 'FGZ
''       politicas 1500 - version 19 Monresa: Se descuenta la misma cantidad que se paga


'Const Version = "3.04"
'Const FechaVersion = "05/10/2010" 'Margiotta, Emanuel
''       politicas 1500 - version 19 Monresa: Se corrigió el desglose para que contemple los tipos de días configurados cuando paga y descuenta.

'Const Version = "3.05"
'Const FechaVersion = "19/11/2010" 'FGZ
''       politicas 1500 - version 19 Monresa: Se corrigió el desglose para que contemple los tipos de días configurados cuando paga y descuenta.
''                       Ademas ahora pago todo junto y descuenta por mes (siempre por la cantidad de dias de licencia)

'Const Version = "3.06"
'Const FechaVersion = "25/11/2010" 'FGZ
''       politicas 1500 - version 19 Monresa: Se modificó nuevamente la logica a pedido del cliente.
''       Ademas se corrigió un problema con el alcance individual de las politicas.

'Const Version = "3.07"
'Const FechaVersion = "02/12/2010" 'FGZ
''       politicas 1500 - version 19 Monresa: Se corrigió la forma en que desglosaba cuando la cantidad de dias a descontar es mayor quer la licencia.


'Const Version = "3.08"
'Const FechaVersion = "06/01/2011" 'FGZ
'       Se comentaron las siguientes 2 declaraciones de variables globales dado que estaban definias tambien en el modulo de politicas como globales
'       Global fecha_desde As Date
'       Global fecha_hasta As Date

'Const Version = "3.09"
'Const FechaVersion = "25/02/2011" 'FAF
'       Se agrego la version 20 a partir de la 12 a la politica 1500
'       La unica diferencia es que en el descuento lo realiza segun la definicion de Pago/Dto de los modelos
'       definidos en la politica 1503

'Const Version = "3.10"
'Const FechaVersion = "06/12/2011" 'Deluchi, Ezequiel
''       Se modifico la version 8 de la politica 1500 de pago/dtos para días correspondientes Politica1500PagayDescuenta_Mes_a_Mes()
''       Se chequea año para licencias de mas de 33 y empiezan un mes q no es diciembre y terminan el prox año


'Const Version = "3.11"
'Const FechaVersion = "23/01/2012" 'Margiotta, Emanuel (CAS 13972)
''       Se agrego la version 21 de la politica 1500 para cooperativa seguro. Genera el pago de dias correspondientes segun el campo (vdiascorcantcorr).
''       que se configura en la política 1501 st_tipodia2.

'Const Version = "3.12"
'Const FechaVersion = "25/01/2012" 'Margiotta, Emanuel
''       Se agrego a la version 11 de la politica 1500 para que obtenga el vacnro recorriendo los periodos de vacaciones en el rango de fecha procesado,
''       como lo hacen las demas versiones

'Const Version = "3.13"
'Const FechaVersion = "08/02/2012" 'Margiotta, Emanuel (CAS 13972)
''       Se modifico la forma de pago 12 para esta version paga todo y descuenta por mes
''       Se agrego la version 21 de la politica 1500 para cooperativa seguro. Genera el pago de dias correspondientes segun el campo (vdiascorcantcorr).
''       que se configura en la política 1501 st_tipodia2. Se


'Const Version = "3.14"
'Const FechaVersion = "09/04/2012" 'Margiotta, Emanuel (CAS 14995)
''       Se modifico la forma de pago 12 para esta version paga todo y descuenta por mes
''       Se agrego la version 22 de la politica 1500 para TIMBO. Genera el pago y descuento a partir de la licencia y lo genera por Quincena.
''       En caso de que la licencia abarque mas de una quincena parte el pago y descuento para ajustar a la quincena del mes que corresponda.

'Const Version = "3.15"
'Const FechaVersion = "16/05/2012" 'Gonzalez Nicolás - PORTUGAL
''       Se agregó versión 23 de la política 1500 para PORTUGAL. Genera solo pagos.
''       Se agregó Politica1500_PagoLic_PT - Paga por licencias
''       Se agregó Politica1500_PagoXdiasCorr_PT - Paga por licencias


'Const Version = "3.16"
'Const FechaVersion = "22/05/2012" 'Margiotta, Emanuel (15970)
''   Politica 1500 version 22 TIMBO: Se corrigio la version 22 de la poltica ya que pagaba un dia menos cuando se daba una serie de condiciones.

'Const Version = "3.17"
'Const FechaVersion = "20/07/2012" 'Margiotta, Emanuel (13972)
'   Politica 1500 version 21 Coperación Seguro: Se comento la linea que setea a la variable (Aux_TipDiaDescuento=7) cuando es descuento y la forma de pago es Menusal.

'Const Version = "3.18"
'Const FechaVersion = "26/07/2012" 'Gonzalez Nicolás - Se agregaron Logs - PORTUGAL

'Const Version = "3.19"
'Const FechaVersion = "05/12/2012" 'Margiotta, Emanuel (15210)
'        'Se fusionaron agunas versiones en una y se corrigieron otras que no funcionaban
'        'version 12: Se corrigio para el caso de OSDOP la versión ya que no estaba haciendo bien los descuentos.

'Const Version = "3.20"
'Const FechaVersion = "18/12/2012" 'Margiotta, Emanuel (13764)
        'Se fusionaron algunas versiones y se corrigieron otras que no funcionaban, además de sacar algunas que no funcionaba

'Const Version = "3.21"
'Const FechaVersion = "01/08/2013" 'Gonzalez Nicolás - CAS-20507 - HORWATH LITORAL - AMR - Politica de Vacaciones
        'Nueva: Politica 1500 version 25. Paga y desuenta con top de 30 días.

'Const Version = "3.22"
'Const FechaVersion = "08/08/2013" 'Gonzalez Nicolás - CAS-20507 - HORWATH LITORAL - AMR - Politica de Vacaciones
                                'Versión 25:se modifico la forma de calcular los topes.
                
'Const Version = "3.23"
'Const FechaVersion = "12/08/2013" 'Gonzalez Nicolás - CAS-20507 - HORWATH LITORAL - AMR - Politica de Vacaciones
                                'Versión 25: Corrección al partir días en febrero
                
'Const Version = "3.24"
'Const FechaVersion = "21/08/2013" 'Gonzalez Nicolás - CAS-20507 - HORWATH LITORAL - AMR - Politica de Vacaciones
                                'Versión 25: Corrección al partir días.
                
'Const Version = "3.25"
'Const FechaVersion = "23/09/2013" 'Sebastian Stremel - CAS-20818 - OSDOP - PROCESO MASIVO DE VACACIONES
                                  'Se crea version 26 - paga adelantado y descuenta por licencia en el periodo que corresponda.
                
'Const Version = "3.26"
'Const FechaVersion = "22/10/2013" 'Gonzalez Nicolás
                                  'Versión 25: Corrección al partir días. correccion al partir días con los meses de 31 días.

'Const Version = "3.27"
'Const FechaVersion = "24/10/2013" 'Gonzalez Nicolás
'                                  'Versión 25: Corrección al calcular días.
                                  
                                  
'Const Version = "3.28"
'Const FechaVersion = "24/10/2013" 'Margiotta, Emanuel
                                  'Versión 27: Se agrego una nueva versión para SV. Cada 5 días paga 2 días de mas.

'Const Version = "3.29"
'Const FechaVersion = "06/01/2014" 'Carmen Quintero
'                                  'Versión 28: Se agrego una nueva versión para SGS.
'                                  'Paga todo lo que le corresponde y descuenta todo en funcion al primer dia de licencia que se tome .
                                  

                                  
'Const Version = "3.30"
'Const FechaVersion = "24/10/2014" 'Fernandez, Matias - CAS-27575 - SGS - ERROR EN GENERACION DE PEDIDO DE PAGO
                                  ' control de periodos a los cuales imputar los pagos y descuentos y correccion en la query principal cuando se paga por licencia.



'Const Version = "3.31"
'Const FechaVersion = "27/10/2014" 'Fernandez, Matias - CAS-27575 - SGS - ERROR EN GENERACION DE PEDIDO DE PAGO
                                   ' el periodo a imputar tiene q estar abierto


'Const Version = "3.32"
'Const FechaVersion = "10/11/2014" 'Fernandez, Matias - CAS-27575 - SGS - ERROR EN GENERACION DE PEDIDO DE PAGO
'                                   ' Se paso la variable de listapgdto al modulo de politicas.


'Const Version = "3.33"
'Const FechaVersion = "22/01/2015" 'Fernandez, Matias - CAS-29151 - BDO BASE HAVA - BUG EN PAGO Y DESCUENTO
''    Version 14. Se le agregó control que la cantidad de dias a generar no supere los restantes
''     politica 1500 version 22, control sobre campos nulos


'Const Version = "3.34"
'Const FechaVersion = "18/02/2015" 'Fernandez, Matias - CAS-29481 - CDA - Error en pago de vacaciones
'    Se filtra por periodos, llegan por parametros desde el asp a batch_proceso, se separan por @
'    Version 14. Se le corrigió control de ultimo dia en febrero (FZ)
'    politica 1500 version 29, Nueva version para TATA. Copia de V14 pero con calculo de dia de inicio segun licencia(FZ)
   

'Const Version = "3.35"
'Const FechaVersion = "09/03/2015" 'Fernandez, Matias -CAS-29553 - VSO - Interpack - inconveniente en pago descuento
                                  ' control sobre jornales y mensuales para politica 1500 version 22
                                  ' pago y descuento por licencias.

'Const Version = "3.36"
'Const FechaVersion = "10/03/2015" 'Fernandez, Matias -CAS-29553 - VSO - Interpack - inconveniente en pago descuento
                                  ' control entre dias habiles y corridos
                                  
                                  

'Const Version = "3.37"
'Const FechaVersion = "18/03/2015" 'Fernandez, Matias -CAS-21778 - Sykes El Salvador -  Bug Pago Descuento Vacaciones
                                  ' se agrega la funcion Politica1500_V2PagaDescuenta_PorQuincena_sykes (ver. 27 pol 1500)
                                  ' se agrega funcion FormaDeLiquidacion_sykes
                                  
                                  
'Const Version = "3.38"
'Const FechaVersion = "30/03/2015" 'Fernandez, Matias -CAS-29553 - VSO - Interpack - inconveniente en pago descuento
                                  ' control cuando es solo pago o solo descuento
                                  
'Const Version = "3.39"
'Const FechaVersion = "13/05/2015" 'Fernandez, Matias -CAS-29481 - CDA - Error en pago de vacaciones
                                  ' Correccion en mensajes de log cuando se selecciona por periodos y no hay licencias.
Const Version = "3.40"
Const FechaVersion = "27/05/2015" ' Fernandez, Matias - CAS-29553 - VSO - Interpack - inconveniente en pago descuento
                                  ' politica 1500 version 22 - correccion en cambios de quincena, el 15 pertenece a la primer quincena

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

Global Ternro As Long
Global NroProceso As Long
Global CEmpleadosAProc As Integer
Global CDiasAProc As Integer
Global IncPorc As Single
Global Progreso As Single
Global fec_proc As Integer ' 1 - Política Primer Reg.
                           ' 2 - Política Reg. del Turno
                           ' 3 - Política Ultima Reg.
Global Usa_Conv As Boolean
Global diatipo As Byte
Global ok As Boolean

'FGZ - Se comentaron las siguientes 2 declaraciones de variables dado que son globales ----------------
'Global fecha_desde As Date
'Global fecha_hasta As Date
'FGZ - Se comentaron las siguientes 2 declaraciones de variables dado que son globales ----------------
Global fecha_hasta_txt As String
Global Tdias As Integer
Global Thoras As Integer
Global Tmin As Integer
Global Cod_justificacion1 As Long
Global Cod_justificacion2 As Long

Global Existe_Reg As Boolean

Global tiene_turno As Boolean
Global Nro_Turno As Long
Global Tipo_Turno As Integer

Global Tiene_Justif As Boolean
Global nro_justif As Long
Global justif_turno As Boolean
Global p_turcomp As Boolean
Global nro_grupo As Long
Global Nro_fpgo As Integer
Global Fecha_Inicio As Date
Global P_Asignacion  As Boolean
Global Trabaja     As Boolean ' Indica si trabaja para ese dia
Global Orden_Dia As Integer
Global Nro_Dia As Integer
Global Nro_Subturno As Integer
Global Dia_Libre As Boolean
Global dias_trabajados As Integer
Global Dias_laborables As Integer

Global aux_Tipohora As Integer
Global aux_TipoDia As Integer

Global E1 As String
Global E2 As String
Global E3 As String
Global S1 As String
Global S2 As String
Global S3 As String
Global FE1 As Date
Global FE2 As Date
Global FE3 As Date
Global FS1 As Date
Global FS2 As Date
Global FS3 As Date

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

Global Cant_emb As Integer
Global toltemp As String
Global toldto As String
Global acumula As Boolean
Global acumula_dto As Boolean
Global acumula_temp As Boolean
Global convenio As Long

Global tdias_oblig As Single
Global Tipo_Hora As Integer
Global HuboErrores As Boolean
Global SinError As Boolean
'Global listapgdto 'mdf


Public Sub Main()
'-----------------------------------------------------------------------
' Procedimiento Inicial del proceso de Generacion de Pago/dto
'Autor: FGZ
'Fecha: 27/07/2005
'-----------------------------------------------------------------------
Dim Fecha As Date
Dim parametros As String
Dim cantdias As Integer
Dim Columna As Integer
Dim Mensaje As String
Dim Genera As Boolean
Dim NroTPV As String

Dim pos1 As Integer
Dim pos2 As Integer

Dim objReg As New ADODB.Recordset
Dim strCmdLine As String
Dim objconnMain As New ADODB.Connection
Dim Archivo As String

Dim rs As New ADODB.Recordset
Dim rs_Vac As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset
Dim PID As String
Dim Ultimo_Empleado As Long
Dim ArrParametros
Dim Proc_Param
Dim PliqNro
Dim listaperiodos 'mdf
Dim porperiodos As Boolean 'mdf
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
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
            If IsNumeric(strCmdLine) Then
                NroProceso = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    
'    'NG - Se crea log antes de CargarConfiguracionbasica | 27/07/2012
'    Archivo = PathFLog & "Vac_PagoDto" & "-" & NroProceso & ".log"
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set Flog = fs.CreateTextFile(Archivo, True)
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'seteo del nombre del archivo de log
    'Creo el archivo de texto del desglose
    Archivo = PathFLog & "Vac_PagoDto" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Archivo, True)
    
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
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    OpenConnection strconexion, objConnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error GoTo ME_Main
    
    'FGZ - 05/08/2009 --------- Control de versiones ------
    Version_Valida = ValidarVBD(Version, 14, TipoBD, 0)
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
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Flog.writeline "Levanta Proceso y Setea Parámetros:  " & " " & Now
       
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE bpronro =" & NroProceso
    OpenRecordset StrSql, rs_Batch_Proceso
       
    If rs_Batch_Proceso.EOF Then Exit Sub
    parametros = rs_Batch_Proceso!bprcparam
    
    
    If Not IsNull(parametros) Then
        If Len(parametros) >= 1 Then
'            pos1 = 1
'            pos2 = InStr(pos1, Parametros, ".") - 1
'            NroVac = Mid(Parametros, pos1, pos2)
            Proc_Param = Split(parametros, ".")
            

            'pos1 = 1
            'pos2 = InStr(pos1, Parametros, ".") - 1
            'Reproceso = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
            Reproceso = CBool(Proc_Param(0)) '+++

            'pos1 = pos2 + 2
            'pos2 = InStr(pos1, Parametros, ".") - 1
            'fecha_desde = CDate(Mid(Parametros, pos1, pos2 - pos1 + 1))
            fecha_desde = CDate(Proc_Param(1)) '+++

            'pos1 = pos2 + 2
            'pos2 = InStr(pos1, Parametros, ".") - 1
            'If pos1 > pos2 Then ' Viene vacia
            '    StrSql = " SELECT * FROM vacacion " & _
            '             " WHERE vacnro = " & NroVac
            '    OpenRecordset StrSql, rs_vac
            '    If Not rs_vac.EOF Then
            '        fecha_hasta = CDate(rs_vac!vacfechasta)
            '    End If
            'Else
            '    fecha_hasta = CDate(Mid(Parametros, pos1, pos2 - pos1 + 1))
            'End If
            fecha_hasta_txt = Proc_Param(2)
            If fecha_hasta_txt = "" Then
                StrSql = " SELECT * FROM vacacion " & _
                         " WHERE vacnro = " & NroVac
                OpenRecordset StrSql, rs_Vac
                If Not rs_Vac.EOF Then
                    fecha_hasta = CDate(rs_Vac!vacfechasta)
                End If
                rs_Vac.Close
            Else
                fecha_hasta = CDate(fecha_hasta_txt)
            End If
        
            'pos1 = pos2 + 2
            'pos2 = InStr(pos1, Parametros, ".") - 1
            'GeneraPorLicencia = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
            GeneraPorLicencia = CBool(Proc_Param(3))
        
            If GeneraPorLicencia Then
                'pos1 = pos2 + 2
                'pos2 = InStr(pos1, Parametros, ".") - 1
                'TipoLicencia = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
                TipoLicencia = CLng(Proc_Param(4))
                
                'pos1 = pos2 + 2
                'pos2 = InStr(pos1, Parametros, ".") - 1
                'Todas = CBool(Mid(Parametros, pos1, pos2 - pos1 + 1))
                Todas = CBool(Proc_Param(5))
            
                If Not Todas Then
                    'pos1 = pos2 + 2
                    'pos2 = Len(Parametros)
                    'nrolicencia = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
                    nrolicencia = CLng(Proc_Param(6))
                End If
            Else
                'pos1 = pos2 + 5
                'pos2 = InStr(pos1, Parametros, ".") - 1
                'TotalGeneral_Dias_A_Generar = CLng(Mid(Parametros, pos1, pos2 - pos1 + 1))
                TotalGeneral_Dias_A_Generar = CLng(Proc_Param(7))
                
                'pos1 = pos2 + 2
                'pos2 = Len(Parametros)
                'Generar_Fecha_Desde = CDate(Mid(Parametros, pos1, pos2 - pos1 + 1))
                Generar_Fecha_Desde = CDate(Proc_Param(8))
                Aux_Generar_Fecha_Desde = Generar_Fecha_Desde
                TipoLicencia = 0
                Todas = False
                nrolicencia = 0
            End If
            
            'nuevo parametro
            If UBound(Proc_Param) > 8 Then
                Pliq_Nro = CInt(Proc_Param(9))
                If Pliq_Nro <> 0 Then
                    StrSql = " SELECT pliqmes, pliqanio FROM periodo WHERE pliqnro = " & Pliq_Nro
                    OpenRecordset StrSql, rs_Vac
                    If Not rs_Vac.EOF Then
                        Pliq_Anio = CInt(rs_Vac!pliqanio)
                        Pliq_Mes = CInt(rs_Vac!pliqmes)
                    Else
                        Pliq_Nro = 0
                    End If
                    rs_Vac.Close
                    Flog.writeline "  Parámetros: Pliq=" & Pliq_Nro & "  PliqMes=" & Pliq_Mes & " PliqAño=" & Pliq_Anio
                End If
            Else
                Pliq_Nro = 0
            End If
            '--
        
        '----------------------------------
        '----------------------------------MDF
        
          listaperiodos = Split(Proc_Param(UBound(Proc_Param)), "@")
          porperiodos = False
        If UBound(listaperiodos) > 0 Then
          If listaperiodos(1) <> "" Then
             listaperiodos = "0" & listaperiodos(1) & "0"
            porperiodos = True
            Flog.writeline "Se seleccionaron periodos a mano: " & Replace(Replace(listaperiodos, "0,", ""), ",0", "")
          End If
        End If
        
        '----------------------------------MDF
        
        
        
        
        End If
    End If
          
    If GeneraPorLicencia Then
        If TipoLicencia = 0 Then 'Todos los tipos de Licencias
            If Todas Then
                StrSql = "SELECT * from batch_empleado " & _
                         " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro " & _
                         " INNER JOIN emp_lic ON emp_lic.empleado = batch_empleado.ternro " & _
                         " WHERE (bpronro = " & NroProceso & ") AND " & _
                         " (((elfechadesde >= " & ConvFecha(fecha_desde) & ") AND " & _
                         " (elfechadesde <= " & ConvFecha(fecha_hasta) & ")) " & _
                         " OR ((elfechacert >= " & ConvFecha(fecha_desde) & ") AND (elfechacert <= " & ConvFecha(fecha_hasta) & "))) AND " & _
                         " emp_lic.licestnro = 2 " & _
                         " ORDER BY elfechacert DESC, elfechadesde ASC " 'Autorizada
            Else
                StrSql = "SELECT * from batch_empleado " & _
                         " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro " & _
                         " INNER JOIN emp_lic ON emp_lic.empleado = batch_empleado.ternro " & _
                         " WHERE (bpronro = " & NroProceso & ") AND " & _
                         " (tdnro = " & TipoLicencia & ") " & _
                         " AND emp_lic.licestnro = 2 " 'Autorizada
            End If
        Else 'un tipo de licencia en particular
            If Todas Then  'todas las licencias de ese tipo
               If Not porperiodos Then   'mdf - forma standar
                 StrSql = "SELECT * from batch_empleado " & _
                         " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro " & _
                         " INNER JOIN emp_lic ON emp_lic.empleado = batch_empleado.ternro " & _
                         " WHERE (bpronro = " & NroProceso & ") AND " & _
                         " (((elfechadesde >= " & ConvFecha(fecha_desde) & ") AND " & _
                         " (elfechadesde <= " & ConvFecha(fecha_hasta) & ")) " & _
                         " OR ((elfechacert >= " & ConvFecha(fecha_desde) & ") AND (elfechacert <= " & ConvFecha(fecha_hasta) & "))) AND " & _
                         " (tdnro = " & TipoLicencia & ") " & _
                         " AND emp_lic.licestnro = 2 " & _
                         " ORDER BY elfechacert DESC, elfechadesde ASC " 'Autorizada
                         
               Else 'seleccion por periodos
                               
                    If TipoLicencia = 2 Then  'licencias por vacaciones
                    
                      StrSql = "SELECT * from batch_empleado "
                      StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro "
                      StrSql = StrSql & " INNER JOIN emp_lic ON emp_lic.empleado = batch_empleado.ternro "
                      StrSql = StrSql & " inner join lic_vacacion on emp_lic.emp_licnro=lic_vacacion.emp_licnro "
                      StrSql = StrSql & "Where(bpronro = " & NroProceso & ") AND "
                      StrSql = StrSql & " lic_vacacion.vacnro in (" & listaperiodos & ")"
                      StrSql = StrSql & " ORDER BY elfechacert DESC, elfechadesde ASC "
                    Else
                    
                    Flog.writeline "Para filtrar por periodos, la licencia debe ser de vacaciones..."
                    Flog.writeline "Proceso Terminado :("
                    HuboErrores = True
                    GoTo Final
                    End If
               End If
                         
            Else  ' se selecciono una licencia en particular
                StrSql = "SELECT * from batch_empleado " & _
                         " INNER JOIN empleado ON empleado.ternro = batch_empleado.ternro " & _
                         " INNER JOIN emp_lic ON emp_lic.empleado = batch_empleado.ternro " & _
                         " WHERE (bpronro = " & NroProceso & ") AND " & _
                         " (tdnro = " & TipoLicencia & ") AND " & _
                         " (emp_lic.emp_licnro = " & nrolicencia & ") " & _
                         " AND emp_lic.licestnro = 2 " 'Autorizada
            End If
        End If
    Else ' no se genera por licencia
        StrSql = "SELECT * from batch_empleado "
        StrSql = StrSql & " INNER JOIN empleado ON batch_empleado.ternro = empleado.ternro"
        StrSql = StrSql & " WHERE bpronro  =" & NroProceso
    End If
    OpenRecordset StrSql, objReg
    
    If objReg.RecordCount > 0 Then
        CEmpleadosAProc = objReg.RecordCount
        IncPorc = (100 / CEmpleadosAProc)
    Else
       If Not porperiodos Then
        Flog.writeline "No se encontraron Licencias entre el " & fecha_desde & " y el " & fecha_hasta
        Flog.writeline " SQL " & StrSql
       Else
         Flog.writeline "No se encontraron licencias para los periodos: " & Replace(Replace(listaperiodos, "0,", ""), ",0", "")
       End If
    End If
    
    SinError = True
    HuboErrores = False
    PrimeraVez = True
    Ultimo_Empleado = 0
    listapgdto = "0"
    Do While Not objReg.EOF
        MyBeginTrans
        Ternro = objReg!Ternro
        Empleado.Ternro = objReg!Ternro
        Empleado.Legajo = objReg!empleg
        Flog.writeline
        Flog.writeline "Inicio Empleado:" & objReg!empleg
        Total_Dias_A_Generar = TotalGeneral_Dias_A_Generar
        Aux_Generar_Fecha_Desde = Generar_Fecha_Desde
        Flog.writeline "Dias a generar: " & Total_Dias_A_Generar
        Flog.writeline "a partir de: " & Aux_Generar_Fecha_Desde
' ----------------------------------------------------------
        'NroLic = objReg!emp_licnro
        If GeneraPorLicencia Then
            nrolicencia = objReg!emp_licnro
            TipoLicencia = objReg!tdnro
'            If PrimeraVez Then
'                primer_mes = Month(objReg!elfechadesde)
'                primer_ano = Year(objReg!elfechadesde)
'                PrimeraVez = False
'            End If
            If Ultimo_Empleado <> objReg!Ternro Then
                primer_mes = Month(objReg!elfechadesde)
                primer_ano = Year(objReg!elfechadesde)
                PrimeraVez = False
                Ya_Pago = False
                CantidadLicenciasProcesadas = 0
            End If
            Ultimo_Empleado = objReg!Ternro
            
            ' Hasta aca, fecha_desde = fecha desde periodo vacaciones. A partir de aca es
            ' la fecha desde de la licencia del empleado
            fecha_desde = objReg!elfechadesde
        Else
            'Cargo todos los periodos de vacaciones
           If Not porperiodos Then   'mdfffff
                StrSql = "SELECT * FROM vacacion "
                StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(fecha_hasta)
                StrSql = StrSql & " AND  vacfechasta >= " & ConvFecha(fecha_desde)
                StrSql = StrSql & " ORDER BY vacnro"
                OpenRecordset StrSql, rs_Periodos_Vac
           
           Else 'mdf
                StrSql = "SELECT * FROM vacacion inner join vacdiascor on vacdiascor.vacnro= vacacion.vacnro "
                StrSql = StrSql & " WHERE vacacion.vacnro in (" & listaperiodos & ")"
                StrSql = StrSql & " ORDER BY vacacion.vacnro"
                OpenRecordset StrSql, rs_Periodos_Vac
               
           
           End If
           
            ' Hasta aca, fecha_desde = fecha desde periodo vacaciones. A partir de aca es
            ' la fecha "A partir del" que viene del asp
            fecha_desde = Aux_Generar_Fecha_Desde
        End If
        Call Politica(1500) 'Genera pago/dto por Licencias
        MyCommitTrans
' ----------------------------------------------------------
siguiente:
            Progreso = Progreso + IncPorc
            
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
            
        If SinError Then
             ' borro
             StrSql = "DELETE FROM batch_empleado WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
             objConn.Execute StrSql, , adExecuteNoRecords
        Else
             StrSql = "UPDATE batch_empleado SET estado = 'Error' WHERE ternro = " & Ternro & " AND bpronro = " & NroProceso
             objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        objReg.MoveNext
    Loop

'Deshabilito el manejador de errores
On Error GoTo 0

Final:
Flog.writeline "Fin :" & Now
Flog.Close
   
    If HuboErrores Then
        ' actualizo el estado del proceso a Error
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        ' poner el bprcestado en procesado
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    
        ' -----------------------------------------------------------------------------------
        'FGZ - 22/09/2003
        'Copio el proceso en el historico de batch_proceso y lo borro de batch_proceso
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
        If Not IsNull(rs_Batch_Proceso!empnro) Then
            StrSql = StrSql & ",empnro"
            StrSqlDatos = StrSqlDatos & "," & rs_Batch_Proceso!empnro
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
        ' FGZ - 22/09/2003
        ' -----------------------------------------------------------------------------------
    End If
    'MyCommitTrans
    
Fin:
    If objConn.State = adStateOpen Then objConn.Close
    Set objConn = Nothing
    If objReg.State = adStateOpen Then objReg.Close
    Set objReg = Nothing
    Exit Sub
    
ME_Main:
    MyRollbackTrans
    HuboErrores = True
    SinError = False
    'Resume Next
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro & " " & Fecha
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"
    GoTo siguiente
End Sub
