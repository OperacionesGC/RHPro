Attribute VB_Name = "mdlValidarBD"
Option Explicit
'Version: 1.01  'Inicial

'Const Version = 1.01    'Dias Correspondientes Parciales
'Const FechaVersion = "15/10/2005"

'Const Version = 1.02    'Version con otra conexion para el progreso
'Const FechaVersion = "23/11/2005"


'Const Version = 2.01     'Revision general
'Const FechaVersion = "01/12/2005"

'Const Version = 2.02     'se calcula la fecha_hasta como 31/12/anio del periodo de vac
'Const FechaVersion = "25/01/2006"

'Const Version = 2.03     'los dias habiles considerados para hacer la proporcion de dias cuando
                         'la antiguedad es menor a 6 meses
                         'se saca de la configuracion del tipo de vacaciones (Corridos, Habiles de L-V, de L-S, etc) de acuerdo a un tercer parametro en la politica 1501
'Const FechaVersion = "16/02/2006"

'Const Version = 2.04     'se agrega nueva version de la politica 1505
                         'se agrega un nuevo parametro a la politica 1501
                         'se crea la politica 1508, se agrego la logica de la misma
                         'se agrego dos nuevos tipos de parametros:
                         '       18-BaseAntiguedad (Int)
                         '       19-Factor (Double)
                         'se agrego un case a la Version Base Antiguedad que usa Uruguay
'Const FechaVersion = "29/06/2006"

'Const Version = 2.05    'Lisandro Moro
'                        'Se corrigio la forma de calculode las vacaciones
'                        ' en la sub Bus_diasVac, cuando DiasProporcion <> 20
'Const FechaVersion = "13/11/2006"

'---------------------------------------------------------------
'Const Version = "2.06"
'Const FechaVersion = "13/11/2007" 'FGZ
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

'----------------------------------------------------------------------------------------
'Const Version = "2.07"
'Const FechaVersion = "29/01/2008"
' Gustavo Ring - Se agrego redondeo para calcular los dias correspondientes
'                cuando el empleado tiene menos de 6 meses trabajados
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
'Const Version = "2.08"
'Const FechaVersion = "01/02/2008"
' Gustavo Ring - Se cambio el nro de parámetro 22 redondeo por el 16 que ya existia
'----------------------------------------------------------------------------------------

'Const Version = "2.09"
'Const FechaVersion = "24/02/2009"
' Gustavo Ring - Se creo custom para calcular antigüedad para Radiotronica
'----------------------------------------------------------------------------------------

'Const Version = "2.10"
'Const FechaVersion = "11/03/2009"
' Gustavo Ring - Se modifico antigüedad para Radiotronica
'----------------------------------------------------------------------------------------

'Const Version = "2.11"
'Const FechaVersion = "12/03/2009"
' Gustavo Ring - Soporte encriptación de cadena de conexión
'----------------------------------------------------------------------------------------

'Const Version = "2.12"
'Const FechaVersion = "14/04/2009"
' Gustavo Ring - Se modifico para tomar o no los dias feriados como trabajados.
'----------------------------------------------------------------------------------------

'Const Version = "2.13"
'Const FechaVersion = "07/05/2009"
' Gustavo Ring - Se modifico la query de las fases.
'----------------------------------------------------------------------------------------

'Const Version = "2.14"
'Const FechaVersion = "05/06/2009"
''Gustavo Ring - Se utiliza el parametro 11 para la ver cuando se proporciona
''----------------------------------------------------------------------------------------


'Const Version = "2.15"
'Const FechaVersion = "25/06/2009" 'FGZ
''           Nueva Politica 1511 - Vacaciones Acordadas.
''               Esta politica revisa las vacaciones acordadas del empleado
''                   y se queda con lo mas conveniente para el empleado.


'Const Version = "2.16"
'Const FechaVersion = "22/10/2009" 'FGZ
''           Nueva Politica 1512 - Vencimiento de vacaciones.
''               Esta politica calcula la cantidad de dias que vencen del periodo anterior y traspasa lo que se pueda al periodo actual.

'Const Version = "2.17"
'Const FechaVersion = "16/11/2009" 'FGZ
''           Problema con la funcion de validacion de version.

'Const Version = "2.18"
'Const FechaVersion = "09/02/2010" 'MB
''           Politica 1508 se agregó la opcion de base de antiguedad 4 y 5 a una fecha dada con los paramtros 30-dia y 31-mes.
''           la base 4 calcula la fecha como dia/mes/año del periodo + 1 y la base 5 calcula la fecha como dia/mes/año del periodo


'Const Version = "2.19"
'Const FechaVersion = "04/03/2010" 'FGZ
''           Integracion de la version anterior.
''               sincronizacion de numeros de parametros

'------------------------------------------------------------------------------------
'Const Version = "3.00"
'Const FechaVersion = "14/04/2010" 'FGZ
'           Ahora los periodos de vacaciones ahora pueden tener alcance por estructura
'           hubo que agregar 2 parametros mas


'Const Version = "3.01"
'Const FechaVersion = "15/07/2010" 'Margiotta, Emanuel
'''           Nuevo manera para determinar los días correspondientes para el caso de empleados que hayan trabajado menos de la mitad del año.
'''           Se controla los dias efectivamente trabajados en el ultimo año

'Const Version = "3.02"
'Const FechaVersion = "27/08/2010" 'Margiotta, Emanuel
'''           Se agregó el cálculo de dias correspondientes para distintos paises.
'''           Se agregó el cálculo de días corresp. para Uruguay.

'Const Version = "3.03"
'Const FechaVersion = "23/09/2010" 'FGZ
''           Se agregó detalle de log cuando levanta los parametros

'Const Version = "3.04"
'Const FechaVersion = "08/10/2010" 'Margiotta, Emanuel
''           Se comentó una linea en la función de bus_DiasVac_uy la busqueda de antiguedad 2 "Uruguay" porque hacia 2 busquedas seguidas

'Const Version = "3.05"
'Const FechaVersion = "08/10/2010" 'Margiotta, Emanuel
''           Se agregó la validación en la política 1513 cuando no tiene configurada la cantidad de hs diarias

'Const Version = "3.06"
'Const FechaVersion = "04/11/2010" 'Margiotta, Emanuel
''           Se Corrigió la descripción del tipo de dia de vacaciones cuando no lo encuentra en la escala y lo tiene que levantar de la Politica 1501
''           Se agrego a la Pol. 1501 el parametro de redondeo.

'Const Version = "3.07"
'Const FechaVersion = "04/11/2010" 'Margiotta, Emanuel
''           Se cambio en la funcion busqueda de días de vacaciones de uruguay El tipo Base de la antiguedad para que calcule los dias
''           correspondientes al periodo que se esta generando y no al anterior.


'Const Version = "3.08"
'Const FechaVersion = "16/11/2010" 'Margiotta, Emanuel
''           Se saco el calculo de Vencimiento de vacaciones para La Caja


'Const Version = "3.09"
'Const FechaVersion = "19/11/2010" 'FGZ
''           Se cambió el calculo proporcional de dias para Uruguay


'Const Version = "3.10"
'Const FechaVersion = "03/12/2010" 'FGZ
''           Politica 1505. Habia quedado mal los paramtros que utiliza ademas de que no estaba configurable.
''               Los parametros configurables que debe utilizar son
''                   35 - Dia de una fecha
''                   36 - Mes de una fecha

'Const Version = "3.11"
'Const FechaVersion = "12/04/2011" 'Lisandro Moro
''               cas 11821 - Se creo la version colombia

'Const Version = "3.12"
'Const FechaVersion = "12/04/2011" 'Margiotta, Emanuel
''               Se creo la version para Costa Rica - SYKE

'Const Version = "3.13"
'Const FechaVersion = "27/07/2011" 'FGZ
''               Hay un parametro que si viene en NULL aborta. Se controla
''               Se modificó el calculo de antiguedad en el ultimo año

'Const Version = "3.14"
'Const FechaVersion = "11/10/2011" 'EAM
''               Se modificó el calculo de vacaciones de Costa Rica

'Const Version = "3.15"
'Const FechaVersion = "18/10/2011" 'EAM
''               Se agregó la versión 2 de la politica 1513 para que se descuenten los días feriados - Andreani

'Const Version = "3.16"
'Const FechaVersion = "31/10/2011" 'EAM
''               Se corrigio para el procesos planificado de CR- ya que procesa todos los empleado y originariamente los levantaba de batch_empleado y es este caso no existen.

'Const Version = "3.17"
'Const FechaVersion = "06/12/2011" 'EAM
''               Se corrigio la politica 1501, estaba tomando el redondeo siempre con el valor 2
''               Se modifico el comentario del log para CR cuando no tiene Parte de Asignacion Horaria

'Const Version = "3.18"
'Const FechaVersion = "10/02/2012" 'EAM
''               CAS(13972) Se agrego a la política 1501 tipodia2 que sirve para buscar en la escala de vacaiones la cantidad de dias que le corresponde
''               segun el tipo de dia y en tipodia1 se configura el tipo de vacaciones por default.
''               Se agrego el calculo para los empleado con menos de 6 meses de trabajo (sin escala) y relacionado la politica 1513.


'Const Version = "3.19"
'Const FechaVersion = "17/04/2012" 'EAM
'               CAS(15527) Se modifico la función de bus_DiasVac_CR para los empleados con Asignación Horaria que busque los movimientos a partir
'               de la última fecha se proceso, ya que si es por semana completa Ej. del 22/03/12 al 29/03/12 y la primer semana tiene 5 movimientos
'               y la segunda 4 da como resultado que trabaja 9 dias a la semana y es incorrecto.
'               Se seteo la variable NroTPVCorr de la version 3.18 con la variable Columna2 ya que sino quedaba sin valor y daba error.
'               Se modifico el factor de división en el calculo de dias correspondientes por 350

'Const Version = "3.20"
'Const FechaVersion = "17/05/2012" 'Gonzalez Nicolás -  DEMO PORTUGAL
'                                  - Se creó módulo mdlValidarBD, el cual realiza el control de versiones.
'                                  - Se agregó versionado para la política 1514
'                                  - Se agregó función CalcularBeneficioVac_PT() - Calcula los dias PLUS
'                                  - Se agrego para el modelo standard | AnioaProc = Periodo_Anio
'                                  - Se agrego para el modelo standard | auxNroVac = NroVac
'                 Javier Irastorza - Se cambio la manera de determinar los días hábiles de la jornada en Función bus_DiasVac_CR() SYKES - CR
'
'Const Version = "3.21"
'Const FechaVersion = "24/05/2012" 'Gonzalez Nicolás -  CAS-15527 - Sykes -  Correccion dias de Beneficio
''                                  - Se pasa la fecha desde del período de vacaciones a la función CalcularBeneficioVac()

'Const Version = "3.22"
'Const FechaVersion = "21/06/2012" 'Margiotta, Emanuel-  CAS-16247 - Sykes - Se corrigió el ecuación de calculo por el valor 364

'Global Const Version = "3.24"
'Global Const FechaVersion = "07/08/2012"  'Margiotta, Emanuel-  CAS-16247 - Sykes - Se cambio la ecuacion cuando tienen partes de movilidad.
'                                  'Ecuacion: [Días horario del período] * ( 14 / 364 )
'                                  'Se agregaron algunos log y correcciona para cuando la fecha de procesamiento era el dia siguiente a la fecha de corte del período
'                                  'Se cambio la forma en que procesa los empleados con movilidad. Arma sub-intervalos y los analiza. Ademas procesa mas de un Periodo de vacacion.
'                                  'Para los mensuales se modificó para que procese mas un un período.
                                  
'Global Const Version = "3.25"
'Global Const FechaVersion = "13/09/2012"  'Margiotta, Emanuel-  CAS-16247 - Sykes
                                          'Se seteo el parametro FactorDivision en 0 para los empleados que no tienen configurado el grupo de vacaciones.

'Global Const Version = "3.26"
'Global Const FechaVersion = "20/09/2012"  'Margiotta, Emanuel - CAS-16247 - Sykes
'                                        'Se modifico para que la fecha hasta de procesamiento sea el primer domingo hacia atras a la fecha de procesamiento ya que la semana actual es la planificada.

'Global Const Version = "3.27"
'Global Const FechaVersion = "10/10/2012"  'Margiotta, Emanuel - CAS-13764 - H&A
'                                        'Se modifico en la función setear parametros de las politicas para que el parametro st_TipoDia2 controle que pueda venir un valor vacio aperte de null "".
                                        
'Global Const Version = "3.28"
'Global Const FechaVersion = "28/01/2013"  'Margiotta, Emanuel - CAS-18231 - Sykes
'    'Se agrego validación para aquellos empleados que fueron dados de alta un 29 de febrero (bisiesto)
'    'Se modifico la sql para que además de los activos tome los inactivos con la fecha de generacion de días corresp. menores a la fecha de cierre de la fase.

'Global Const Version = "3.29"
'Global Const FechaVersion = "01/02/2013"  'Margiotta, Emanuel - CAS-18231 - Sykes
'    'Corrige un bug para empleado syke
    
'Global Const Version = "3.30"
'Global Const FechaVersion = "19/04/2013"  'Margiotta, Emanuel - CAS-18891 - Telefax-Tata
    'Se agrego nueva política 1516. Descuenta segun la configuración cada ciertas licencias, dias de vacaciones.


'Global Const Version = "3.31"
'Global Const FechaVersion = "26/09/2013"  'Sebastian Stremel - Se creo modelo para el salvador case 7
    

'Global Const Version = "3.32"
'Global Const FechaVersion = "23/09/2013"  'Mauricio Zwenger - CAS-21183 -
    'se agregaron los campos tipvacnrocorr, diasacordcorr y se setea cantdiasCorr=diasacordcorr si corresponde dias acordados

'Global Const Version = "3.33"
'Global Const FechaVersion = "08/10/2013"  'Sebastian Stremel - Se creo modelo para el salvador case 7 - CAS-21472 - Sykes El Salvador - Integración Caso COOP. GIV

'Global Const Version = "3.34"
'Global Const FechaVersion = "18/10/2013"  'Carmen Quintero - CAS-21871 - SYKES - Error Calculo Días Correspondientes Sykes CR
'                                          ' Se modificó el modelo de sykes 4, para que tome en cuenta los empleados que se reincorporan en un mismo periodo de vacacion

'Global Const Version = "3.35"
'Global Const FechaVersion = "24/10/2013"  'Dimatz Rafael - CAS-21739 - Telefax Tata - Error Dias Correspondientes
                                          ' Se creo un RedondearNumero un Case con valor 2 que toma la parte Entera del Numero

'Global Const Version = "3.36"
'Global Const FechaVersion = "12/11/2013"  ' Gonzalez Nicolás - CAS-19425 - H&A - Mapeo GIV multi-pais R4v1
                                          ' Se agregó MdlVac_Paraguay, MdlFuncionesVs y Mdlidioma.
                                          ' Se agregó Política 1515
                                          ' Se agrega validación para política 1515 (Solo para PY)
                                          ' Nuevas Funciones: PeriodoCorrespondienteAlcance() Y ValidaModeloyVersiones() en MdlPoliticasVac
'Global Const Version = "3.37"
'Global Const FechaVersion = "14/11/2013"  ' Gonzalez Nicolás - CAS-21871 - SYKES - Error Calculo Días Correspondientes Sykes CR. Cuando el empleado estuvo de baja se actualiza la fecha en vacdiascor

'Global Const Version = "3.38"
'Global Const FechaVersion = "18/11/2013"  'Margiotta, emanuel - CAS-21472 - SYKES EL salvador - Naiconalización de El Salvador GIV

'Global Const Version = "3.39"
'Global Const FechaVersion = "27/11/2013"  'Margiotta, emanuel - CAS-21472 - SYKES EL salvador - Se corrigio la validación cuando no entra por escala y tiene mas de 6 meses de trabajo.

'Global Const Version = "3.40"
'Global Const FechaVersion = "05/12/2013" ' Gonzalez Nicolás - CAS-21871 - SYKES - Error Calculo Días Correspondientes Sykes CR
                                         ' Se corrigió fecha que se actualizaba en tabla diascord, toma la fecha de alta.
                                         ' Si la pol 1515 esta Activa y no tiene parámetros configurados devuelve False.
                                         ' Se modifico log -> Error cargando configuración de la Política 1515"
                                          
'Global Const Version = "3.41"
'Global Const FechaVersion = "21/01/2014" ' Gonzalez Nicolás - CAS-21871 - SYKES - Error Calculo Días Correspondientes Sykes CR
                                         ' Se corrigió fecha que se actualizaba en tabla diascord, toma la fecha de alta.
'Global Const Version = "3.42"
'Global Const FechaVersion = "21/01/2014" ' Gonzalez Nicolás - CAS-21871 - SYKES - Error Calculo Días Correspondientes Sykes CR
                                         ' Se cierra período con fecha de baja y se crea nuevo con fecha de alta.
'Global Const Version = "3.43"
'Global Const FechaVersion = "21/01/2014" ' Gonzalez Nicolás - CAS-21871 - SYKES - Error Calculo Días Correspondientes Sykes CR
'                                         ' Si TipoVacacionProporcion = 0 se updatea con 1.

'Global Const Version = "3.44"
'Global Const FechaVersion = "11/02/2014" ' Margiotta, Emanuel - CAS-21871 - SYKES - Error Calculo Días Correspondientes Sykes CR [Entrega 8]
                                         ' Se agregó la funcionalidad y controles para que una persona que tiene mas de una fase se cree o no los periodos y actualice las fechas de los periodos
                                         'y procesamiento para que no genere movimientos horarios invalidos.
                                         
'Global Const Version = "3.45"
'Global Const FechaVersion = "13/02/2014" 'Fernandez, Matias - CAS-21597 - SGS - Error Carga Masiva Vacaciones - se fija si los empleados estan filtrados
                                         

'Global Const Version = "3.46"
'Global Const FechaVersion = "13/02/2014" 'Margiotta, Emanuel - CAS-21871 - SYKES - Error Calculo Días Correspondientes Sykes CR
''           Se corrigio para aquellos empleados que se le da de baja una fase y quedan inactivos.

'Global Const Version = "3.47"
'Global Const FechaVersion = "05/03/2014" 'Margiotta, Emanuel - CAS-21871 - SYKES - Error Calculo Días Correspondientes Sykes CR
'           Se corrigio el tipo definido para una variable porque generaba error.
 

'Global Const Version = "3.48"
'Global Const FechaVersion = "08/04/2014" 'Fernandez, Matias - CAS-24956 - SYKES - ERROR EN GENERACION DE DIAS CORRESPONDIENTES-valida no ingresar vacdiascor  por duplicado
'           Se corrigio el tipo definido para una variable porque generaba error.


'Global Const Version = "3.49"
'Global Const FechaVersion = "23/04/2014" 'Fernandez, Matias - CAS-24956 - SYKES - ERROR EN GENERACION DE DIAS CORRESPONDIENTES-se valida que el vacnro no venga vacio

'Global Const Version = "3.50"
'Global Const FechaVersion = "20/05/2014" 'Fernandez, Matias - CAS-25237 - SYKES - VACACIONES DE BENEFICIO- se reacomodaron las escalas, y se descomento el update en la funcion que calcula los beneficios

'Global Const Version = "3.51"
'Global Const FechaVersion = "28/05/2014" 'Ruiz Miriam - CAS-25746 - SYKES - Error Planificacion Proceso Días Correspondientes Sykes CR- Se modifico la toma de empleados

'Global Const Version = "3.52"
'Global Const FechaVersion = "02/06/2014" 'Ruiz Miriam - CAS-25746 - SYKES - Error Planificacion Proceso Días Correspondientes Sykes CR- Se modifico la toma de empleados, ahora no calculaba para empleados individuales

'Global Const Version = "3.53"
'Global Const FechaVersion = "05/06/2014" 'Ruiz Miriam - CAS-25746 - SYKES - Error Planificacion Proceso Días Correspondientes Sykes CR- Se modifico el cálculo de los días de beneficio

'Global Const Version = "3.54"
'Global Const FechaVersion = "10/07/2014" 'Fernandez, Matias -CAS-25237 - SYKES - VACACIONES DE BENEFICIO -CDate en fechas, mas impresion de  parametros antes de llamar a diasbeneficio


'Global Const Version = "3.55"
'Global Const FechaVersion = "10/07/2014" 'Mauricio Zwenger - CAS-26102 - CAS-26102 - Tabacal - Fecha Reconocida Provision - Reajuste Egresos

'Global Const Version = "3.56"
'Global Const FechaVersion = "12/11/2014" 'Sebastian Stremel - CAS-26789 - Santander Uruguay - Días de turismo - Licencia Paga - Licencia 25 años

'Global Const Version = "3.57"
'Global Const FechaVersion = "11/02/2015" 'Miriam Ruiz - CAS-28356 - Monasterio Base AMR- custom días correspondientes AMR - se cambió para AMR el calculo de dias trabajados

'Global Const Version = "3.58"
'Global Const FechaVersion = "04/03/2015" 'Fernandez Matias - CAS-29565 - SYKES - Error en procesamiento GIV reingresos- se controla si los dias correspondientes son de una fase vieja cerrada.


'Global Const Version = "3.59"
'Global Const FechaVersion = "30/03/2015" 'Fernandez Matias - CAS-29565 - SYKES - Error en procesamiento GIV reingresos- se controla si los dias correspondientes son de una fase vieja cerrada.

'Global Const Version = "3.60"
'Global Const FechaVersion = "30/03/2015" 'Fernandez, Matias - CAS-21778 - Sykes El Salvador- QA - Bug Giv
                                         'Cada 6 meses genera dias correspondientes, con la mitad de lo q hay al llegar
                                         'al año


'Global Const Version = "3.61"
'Global Const FechaVersion = "17/04/2015" 'Fernandez, Matias - CAS-21778 - Sykes El Salvador- QA - Bug Giv
                                         'hasta antes del aniversario, imputa los dias al periodo anterior.
                                         
                                         
'Global Const Version = "3.62"
'Global Const FechaVersion = "27/04/2015" 'Miriam Ruiz- CAS-28356 - Monasterio Base AMR- custom días correspondientes AMR
                                         'se corrigió redondeo
                                         
Global Const Version = "3.63"
Global Const FechaVersion = "28/04/2015" 'Fernandez, Matias -CAS-29565 - SYKES - Error en procesamiento GIV
                                         'Si una vacacion no tiene sus dias correspondientes, la crea.
                                         
                                         
'---------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------Version no liberada ------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------


'
'Global Const Version = "3.xx"
'Global Const FechaVersion = "17/09/2012"  'Gonzalez Nicolás - CAS-16809 - Heidt & Asoc - Nacionalizacion Vacaciones - PARAGUAY



                                    'Se corrigio el versionado ya que no salia en el log. Le faltaba declarar como Global las constantes.

'---------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------FIN VERSIONES ------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ValidarVBD(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer, ByVal codPais As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : Gonzalez Nicolás
' Fecha      : 04/05/2012
' Modificado : 17/09/2012 - Se agregó versión 6 PARAGUAY
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True
' CODIGOS DE PAISES
 '0,"" - Argentina
 '1 - Uruguay
 '2 - Chile
 '3 - Colombia
 '4 - Costa Rica
 '5 - Portugal
 '6 - Paraguay

    If Version >= "2.16" Then
        'Revisar los campos
        'vacdiascor.venc
        Texto = "Revisar los campos: vacdiascor.venc"
        
        StrSql = "Select venc from vacdiascor where ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "3.18" Then
        'dias correspondientes
        Texto = ""
        Texto = Texto & " vacdiascor.vdiascorcantcorr,  vacdiascor.tipvacnrocorr"
        StrSql = "Select vacdiascor.vdiascorcantcorr,  vacdiascor.tipvacnrocorr FROM vacdiascor WHERE ternro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If
    
    If Version >= "3.19" Then
        'dias correspondientes
        Select Case codPais
            Case 4: 'COSTA RICA - SYKES
                Texto = ""
                Texto = Texto & " vacdiascor.vdiasfechasta"
                StrSql = "Select vacdiascor.vdiasfechasta FROM vacdiascor WHERE ternro = 1"
                OpenRecordset StrSql, rs
                Texto = ""
                Texto = Texto & " vacacion.ternro"
                StrSql = "Select vacacion.ternro FROM vacacion WHERE ternro = 1"
                OpenRecordset StrSql, rs
                V = True
        End Select
    End If
    
    If Version >= "3.26" Then
        'dias correspondientes
        Select Case codPais
            Case 6: 'PARAGUAY
                Texto = ""
                Texto = Texto & " vacacion.ternro"
                StrSql = "Select vacacion.ternro FROM vacacion WHERE ternro = 1"
                OpenRecordset StrSql, rs
                V = True
        End Select
    End If
    
     If Version >= "3.32" Then
        'dias correspondientes
        Texto = ""
        Texto = Texto & " vacdiasacord.tipvacnrocorr,  vacdiasacord.diasacordcorr"
        StrSql = "SELECT tipvacnro, diasacord, tipvacnrocorr, diasacordcorr FROM vacdiasacord where ternro=1"
        OpenRecordset StrSql, rs

        V = True
    End If

    If Version >= "3.36" Then
        If codPais = 6 Then 'PARAGUAY
        'dias correspondientes
            Texto = "Revisar los campos: "
            Texto = Texto & " vacacion.alcannivel"
            StrSql = "Select vacacion.alcannivel FROM vacacion WHERE alcannivel = 1"
            OpenRecordset StrSql, rs
            V = True
            
            Texto = "Revisar los campos: "
            Texto = Texto & " vacdiascortipo.tdnro,vacdiascortipo.progval"
            StrSql = "Select vacdiascortipo.tdnro,vacdiascortipo.progval FROM vacdiascortipo WHERE tdnro = 1"
            OpenRecordset StrSql, rs
            V = True
            
            Texto = "Revisar Existencia de Tabla vac_alcan. Campos: "
            Texto = Texto & " vac_alcan.vacnro,vac_alcan.vacfecdesde,vac_alcan.vacfechasta,vac_alcan.alcannivel,vac_alcan.origen,vac_alcan.vacestado"
            StrSql = "Select vac_alcan.vacnro,vac_alcan.vacfecdesde,vac_alcan.vacfechasta,vac_alcan.alcannivel,vac_alcan.origen,vac_alcan.vacestado FROM vac_alcan WHERE vacnro = 0"
            OpenRecordset StrSql, rs
            V = True
        End If
    End If

'Case Else:
'    Texto = "version correcta"
'    V = True
'End Select
ValidarVBD = V
    
If V = True Then
    Texto = "version correcta"
    Exit Function
End If




ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function
