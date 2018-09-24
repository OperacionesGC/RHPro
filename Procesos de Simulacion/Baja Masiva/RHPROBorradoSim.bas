Attribute VB_Name = "mdlVersiones"
Option Explicit

'
' ------ NIVELACION DE VERSIONES CON LIQUIDADOR --------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------
Global Const Version = "6.75"
Global Const FechaVersion = "20/02/2017" 'EAM
'       (Sprint 108) PBI 6626 - ITRP - 4747498 - RH Pro - Cambio Legal Ganancias - Alquileres
'               ITRP - 4747498 - RH Pro - Cambio Legal Ganancias - Alquileres
'           Nivelación con versión del liquidador.

'Global Const Version = "6.74"
'Global Const FechaVersion = "10/02/2017" 'EAM - (Sprint 107) PBI 4081 - ITRP - 70415 - GC - Liquidacion - Calculo de Mopre
'       (Sprint 107) PBI 6735 - ITRP - 4779473 - MEDICUS - 79473 Sistema no procesa a partir del cambio en liquidador 6.73
'       (Sprint 107) PBI 6744 - ITRP - 4606019 Error en recálculo de imp unico- RH PRO CHILE [SB]
'           Nivelación con versión del liquidador.

'Global Const Version = "6.73"
'Global Const FechaVersion = "27/01/2017" 'EAM - (Sprint 106) PBI 6174 - ITRP - 4593941 - Cambio legal Ganancias 2017
'           Nivelación con versión del liquidador.

'Global Const Version = "6.72"
'Global Const FechaVersion = "18/01/2017" 'LED - (Sprint 106) PBI 6248 - Performance - Global - Payroll Process
''           Nivelación con versión del liquidador.

'Global Const Version = "6.71"
'Global Const FechaVersion = "28/12/2016" 'FGZ - (Sprint 104) PBI 6006 - ITRP - 4532815 - NGA - Farmacity - Liquidador - msg de error en log
'           Nivelación con versión del liquidador.

'Global Const Version = "6.70"
'Global Const FechaVersion = "15/12/2016" 'EAM - (Sprint 103) - 5784 - ITRP - 93331 - GC - Edenor - Calculo de BAE
'           Nivelación con versión del liquidador.

'Global Const Version = "6.69"
'Global Const FechaVersion = "09/11/2016" 'LED - (Sprint 101) - 5093 - Liquidador - Creación de archivo de log
''           Creacion del archivo de log, si falla la conexion.


'Global Const Version = "6.68"
'Global Const FechaVersion = "09/11/2016" 'LED (Sprint 101) - 5149 - ITRP 73054 - Raffo - MODULO SIMULACION
'           Nivelación con versión del liquidador.

'Global Const Version = "6.67"
'Global Const FechaVersion = "07/10/2016" 'EAM (Sprint 98) - 3409 - CAS-37159 - RHPro Consulting CH - Error en recálculo impuesto unico
'           Nivelación con versión del liquidador.

'Global Const Version = "6.66"
'Global Const FechaVersion = "12/09/2016" 'EAM (Sprint 96) - 3386 - CAS-38734 - MONASTERIO BASE APEX CHILE - Búsqueda novedad de concepto
''           Nivelación con versión del liquidador.

'Global Const Version = "6.65"
'Global Const FechaVersion = "08/09/2016" 'EAM (Sprint 96) - 3531 - CAS-38678 - MONASTERIO APEX CHILE - Mejora licencias integradas
'           Nivelación con versión del liquidador.

'Global Const Version = "6.64"
'Global Const FechaVersion = "20/07/2016" 'EAM (Sprint 93) - 2646 - CAS-33601 - RH Pro (Producto) - Peru - Criterios de Búsquedas  - Corrección CTS
''           Nivelación con versión del liquidador.

'Global Const Version = "6.63"
'Global Const FechaVersion = "06/07/2016"
'' MDF - Srpint 91 pbi 182 - nivelacion de version con el liquidador


'Global Const Version = "6.62"
'Global Const FechaVersion = "01/07/2016" 'EAM (Sprint 91) - 764 - CAS-37441 - GRUPO ROCIO -  ADECUACIONES – Bono por Asistencia
'           Nivelación con versión del liquidador.

'Global Const Version = "6.61"
'Global Const FechaVersion = "28/06/2016" 'EAM (Sprint 91) - 1941 - CAS-33601 - RH Pro (Producto) - Peru - Criterios de Busquedas [Entrega 4]
''           Nivelación con versión del liquidador.

'Global Const Version = "6.60"
'Global Const FechaVersion = "15/06/2016" 'EAM (Sprint 90) - 1320 - Error en tipo de búsqueda 138 - Vacaciones vendidas
''           Nivelación con versión del liquidador.

'Global Const Version = "6.59"
'Global Const FechaVersion = "02/06/2016" ' EAM (Sprint 90) - CAS-36566 - IBT - Error en busqueda de concepto acumulador meses fijos
''           Nivelación con versión del liquidador.

'Global Const Version = "6.58"
'Global Const FechaVersion = "02/06/2016" ' EAM (Sprint 89) - CAS-33601 - RH Pro (Producto) - Peru - Criterios de Busquedas [Entrega 3]
''           Nivelación con versión del liquidador.

'Global Const Version = "6.57"
'Global Const FechaVersion = "18/03/2016" ' EAM (Sprint 88) - RH Pro - Argentina - Cambio legal Nuevo reporte Ganancias RG 3839 AFIP
''           Nivelación con versión del liquidador.

'Global Const Version = "6.56"
'Global Const FechaVersion = "18/03/2016" 'EAM - CAS-36167 - RH Pro (Producto) - NOM - Ganancias - Bug Item 56
''           Nivelación con versión del liquidador.

'Global Const Version = "6.55"
'Global Const FechaVersion = "07/03/2016" 'EAM - CAS-33601 - RH Pro (Producto) - Peru - Criterios de Busquedas [Entrega 2]
''           Nivelación con versión del liquidador.

'Global Const Version = "6.54"
'Global Const FechaVersion = "25/02/2016" 'EAM - CAS-35783 - RH Pro (Producto) - ARG - NOM - Ganancias 2016 Decreto 394
''           Nivelación con versión del liquidador.

'Global Const Version = "6.53"
'Global Const FechaVersion = "16/02/2016 " 'EAM - CAS-32751 - LA CAJA - Custom Seguros ADP [Entrega 2]
''           Nivelación con versión del liquidador.

'Global Const Version = "6.52"
'Global Const FechaVersion = "03/02/2016 " 'MDZ - CAS-34564 - MONASTERIO AMR - Bug en simular
''           Nivelación con versión del Simulador.

'Global Const Version = "6.51"
'Global Const FechaVersion = "08/01/2015 " 'EAM - CAS-29467 - NGA- Citricos - Inconveniente en busqueda de escala
''           Nivelación con versión del liquidador.

'Global Const Version = "6.50"
'Global Const FechaVersion = "29/12/2015 " 'EAM - CAS-31053 - RH Pro (Producto) - NAC. PERU - EPS - Solicitud de Búsqueda para Cálculos de EPS [Entrega 4]
''           Nivelación con versión del liquidador.

'Global Const Version = "6.49"
'Global Const FechaVersion = "21/12/2015 " 'EAM - CAS-32751 - LA CAJA - Custom Seguros ADP
''           Nivelación con versión del liquidador.

'Global Const Version = "6.48"
'Global Const FechaVersion = "16/12/2015 " 'EAM - CAS-31053 - RH Pro (Producto) - NAC. PERU - EPS - Solicitud de Búsqueda para Cálculos de EPS [Entrega 4]
''           Nivelación con versión del liquidador.

'Global Const Version = "6.47"
'Global Const FechaVersion = "14/12/2015 " 'EAM - CAS-33601 - RH Pro (Producto) - Peru - Criterios de Busquedas
''           Nivelación con versión del liquidador.

'Global Const Version = "6.46"
'Global Const FechaVersion = "04/12/2015 " 'EAM - CAS-33657 - NGA BASE FREDDO - Bug en sac proporcional tope 30 mensual
''           Nivelación con versión del liquidador.

'Global Const Version = "6.45"
'Global Const FechaVersion = "01/12/2015 " 'EAM - CAS-31053 - RH Pro (Producto) - NAC. PERU - EPS - Solicitud de Búsqueda para Cálculos de EPS [Entrega 3]
''           Nivelación con versión del liquidador.

'Global Const Version = "6.44"
'Global Const FechaVersion = "26/11/2015 " 'EAM - CAS-34164 - NGA - Modificacion de item 56 y 20 ganancias [Entrega 2]
''                                               CAS-33993 - NGA - Ganancias residentes en el extranjero.
''           Nivelación con versión del liquidador.

'Global Const Version = "6.43"
'Global Const FechaVersion = "24/11/2015 " 'EAM - CAS-34164 - NGA - Modificacion de item 56 y 20 ganancias
''                                               CAS-33993 - NGA - Ganancias residentes en el extranjero.
''           Nivelación con versión del liquidador.

'Global Const Version = "6.42"
'Global Const FechaVersion = "19/11/2015 " 'EAM- CAS-31676 - RH Pro - Modificación calculo aguinaldo IRPF [Entrega 3]
''                                         Se corrigio el query del control de versión de la version 6.41

'Global Const Version = "6.41"
'Global Const FechaVersion = "17/11/2015 " 'EAM- CAS-31676 - RH Pro - Modificación calculo aguinaldo IRPF
''                                               CAS-34041 - MONASTERIO (TODAS LAS BASES Y ENTORNOS) - ERROR EN COPIA SIM
''            Nivelación con versión del liquidador.

'Global Const Version = "6.40"
'Global Const FechaVersion = "21/10/2015 " 'EAM- CAS-30667 - RH Pro (Producto) - LIQ - Ganancias - Items en ajuste anual - Filtro DDJJ personal
''            Nivelación con versión del liquidador.

'Global Const Version = "6.39"
'Global Const FechaVersion = "07/10/2015 " 'EAM- CAS-33430 - CIVA - Bug Venta Vacaciones
'                                          'CAS-33210 - SANTANDER URUGUAY - Busqueda base de calculo paros
''            Nivelación con versión del liquidador.

'Global Const Version = "6.38"
'Global Const FechaVersion = "14/09/2015 " 'LED - CAS-33005 - G.COMPARTIDA - Custom en función del liquidador
''            Nivelación con versión del liquidador.

'Global Const Version = "6.37"
'Global Const FechaVersion = "01/09/2015" 'FGZ - CAS-31053 - RH Pro (Producto) - NAC. PERU – EPS – Solicitud de Búsqueda para  Cálculos de EPS [Entrega 2]
'           nivelacion con version del liquidador

'Global Const Version = "6.36"
'Global Const FechaVersion = "26/08/2015" 'FGZ - CAS-32523 - SANTANDER URUGUAY - LIQ – Bug búsqueda de antigüedad
''           nivelacion con version del liquidador

'Global Const Version = "6.35"
'Global Const FechaVersion = "10/08/2015" 'EAM - CAS-29325 - IMECON - Nueva Búsqueda para dias de vacaciones pendientes de goce [Entrega 4]
''Tipo de Busqueda 140 - Saldo vacaciones PE - Se agrego la funcinalidad de días truncos
''Tipo de Busqueda 120 - CTS - Se agrego la opcion de días truncos y se modifico los dias pendientes

'Global Const Version = "6.34"
'Global Const FechaVersion = "31/07/2015" 'EAM - CAS-29325 - IMECON - Nueva Búsqueda para dias de vacaciones pendientes de goce [Entrega 3]
''Tipo de Busqueda 140 - Saldo vacaciones PE - Se agrego la funcinalidad de días truncos
''Tipo de Busqueda 120 - CTS - Se agrego la opcion de días truncos y se modifico los dias pendientes

'Global Const Version = "6.33"
'Global Const FechaVersion = "17/07/2015" 'FGZ
'   CAS-29325 - IMECON - Nueva Búsqueda para dias de vacaciones pendientes de goce
'           nivelacion con version del liquidador

'Global Const Version = "6.32"
'Global Const FechaVersion = "03/07/2015" 'FGZ
''   CAS-31676 - RH Pro - Modificación calculo aguinaldo IRPF
''           nivelacion con version del liquidador

'Global Const Version = "6.31"
'Global Const FechaVersion = "25/06/2015" 'FGZ
''   CAS-31674 - CDA - Bug en liquidación mensual de Junio
''           nivelacion con version del liquidador

'Global Const Version = "6.30"
'Global Const FechaVersion = "24/06/2015" 'FGZ
''   CAS-31053 - RH Pro (Producto) - NAC. PERU – EPS – Solicitud de Búsqueda para  Cálculos de EPS
''           nivelacion con version del liquidador

'Global Const Version = "6.29"
'Global Const FechaVersion = "02/06/2015" 'FGZ
''   CAS-29187 - SYKES EL SALVADOR - Bug Lic Integradas x fecha de certificado
''           nivelacion con version del liquidador

'Global Const Version = "6.28"
'Global Const FechaVersion = "02/06/2015" 'FGZ
''   CAS-31205 - RH Pro (Producto) - LIQ - Ganancias - Modificación escala interna porcentajes
''           nivelacion con version del liquidador


'Global Const Version = "6.27"
'Global Const FechaVersion = "01/06/2015" 'FGZ
''   CAS-29187 - SYKES EL SALVADOR - Bug Lic Integradas x fecha de certificado
''           nivelacion con version del liquidador

'Global Const Version = "6.26"
'Global Const FechaVersion = "28/05/2015" 'FGZ
''           CAS-31075 - Telefax (Santander URU) - Búsqueda de antiguedad
''           nivelacion con version del liquidador

'Global Const Version = "6.25"
'Global Const FechaVersion = "27/05/2015" 'FGZ
''           CAS-31099 - RH Pro (Producto) - LIQ - Ganancias - Corrección fórmula
''           nivelacion con version del liquidador

'Global Const Version = "6.24"
'Global Const FechaVersion = "22/05/2015" 'FGZ
''           CAS-30979 - RH Pro (Producto) - LIQ - Ganancias - RG 3770 nuevo cambio
''           nivelacion con version del liquidador

'Global Const Version = "6.23"
'Global Const FechaVersion = "21/05/2015" 'FGZ
''           CAS-30979 - RH Pro (Producto) - LIQ - Ganancias - RG 3770 nuevo cambio
''           nivelacion con version del liquidador

'Global Const Version = "6.22"
'Global Const FechaVersion = "19/05/2015" 'FGZ
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
''           nivelacion con version del liquidador

'Global Const Version = "6.21"
'Global Const FechaVersion = "18/05/2015" 'FGZ
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
''           nivelacion con version del liquidador

'Global Const Version = "6.20"
'Global Const FechaVersion = "18/05/2015" 'FGZ
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
''           nivelacion con version del liquidador

'Global Const Version = "6.19"
'Global Const FechaVersion = "13/05/2015" 'FGZ
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
''           nivelacion con version del liquidador

'Global Const Version = "6.18"
'Global Const FechaVersion = "12/05/2015" 'FGZ
''           CAS-30516 - GEST COMPARTIDA (EDENOR) - Custom agregar funcion a liquidador
''           nivelacion con version del liquidador

'Global Const Version = "6.17"
'Global Const FechaVersion = "12/05/2015" 'FGZ
''           CAS-30786 - RH Pro (Producto) - LIQ - Ganancias - Cambio en Fórmula - RG 3770
''           nivelacion con version del liquidador

'Global Const Version = "6.16"
'Global Const FechaVersion = "11/05/2015" 'FGZ
''           CAS-29325 - IMECON - Nueva Búsqueda para dias de vacaciones pendientes de goce
''           nivelacion con version del liquidador

'Global Const Version = "6.15"
'Global Const FechaVersion = "11/05/2015" 'FGZ
''           CAS-29032 - Telefax (Santander URU) - Bug Búsqueda de antiguedad
''           nivelacion con version del liquidador

'Global Const Version = "6.14"
'Global Const FechaVersion = "30/04/2015" 'FGZ
''           CAS-29032 - Telefax (Santander URU) - Bug Búsqueda de antiguedad
''           nivelacion con version del liquidador

'Global Const Version = "6.13"
'Global Const FechaVersion = "23/04/2015" 'FGZ
''           CAS-21778 - Sykes El Salvador - QA – Busqueda Prestamos
''           nivelacion con version del liquidador

'Global Const Version = "6.12"
'Global Const FechaVersion = "21/04/2015" 'FGZ
''           CAS-29945 - SYKES EL SALVADOR - Error búsq. Antig Aniversario
''           nivelacion con version del liquidador

'Global Const Version = "6.11"
'Global Const FechaVersion = "20/04/2015" 'FGZ
''           CAS-29945 - SYKES EL SALVADOR - Error búsq. Antig Aniversario
''           nivelacion con version del liquidador

'Global Const Version = "6.10"
'Global Const FechaVersion = "17/04/2015" 'FGZ
''           CAS-30490 - SANTANDER URUGUAY - Error busqueda tiempo trabajado
''           nivelacion con version del liquidador

'Global Const Version = "6.09"
'Global Const FechaVersion = "16/04/2015" 'FGZ
''           CAS-29032 - Telefax (Santander URU) - Bug Búsqueda de antiguedad
''           nivelacion con version del liquidador

'Global Const Version = "6.08"
'Global Const FechaVersion = "09/04/2015" 'EAM
''           CAS-29945 - SYKES EL SALVADOR - Error búsq. Antig Aniversario
''           nivelacion con version del liquidador

'Global Const Version = "6.07"
'Global Const FechaVersion = "06/04/2015" 'EAM
''           CAS-30295 - RH Pro Producto - LIQ - Ganancias - Bug liquidador 6.06 - Falta de item 56
''           nivelacion con version del liquidador

'Global Const Version = "6.06"
'Global Const FechaVersion = "27/03/2015" 'EAM
''           CAS-29261 - Horwath litoral - AMR - Modificación Búsqueda VNG
''               Tipo de Buqueda 124: 'Dias Corresp - Control Baja.
''                   Se debe proporcionar siempre los dias segun los dias trabajados en el ultimo año.
''   Busqueda 124: EAM- se modifico la query que busca las licencias de vacaciones
''   Busqueda 128: EAM- Se modifico la busqueda de licencias por fecha de certificado. Estaba calculando mal cuando era para  febrero los topes.


'Global Const Version = "6.05"
'Global Const FechaVersion = "23/02/2015" 'FGZ
''           CAS-29317 - H y A - LIQ - Bug en Calculo de Impuesto a las Ganancias.
''       nivelacion con version del liquidador

'Global Const Version = "6.04"
'Global Const FechaVersion = "29/01/2015" 'EAM
''   Se modifico la búsqueda 137 y se creo la búsqueda 139
''       nivelacion con version del liquidador

'Global Const Version = "6.03"
'Global Const FechaVersion = "20/01/2015" 'FGZ
''   CAS-27512 - H&A - LIQ - Ganancias - Item 56 Perc.Compras Exterior mensual
''       nivelacion con version del liquidador

'Global Const Version = "6.02"
'Global Const FechaVersion = "05/01/2015" ' FGZ - nivelacion con version del liquidador
''   CAS-28749 - H&A - LIQ - Ganancias - Bug rango de 15000 a 25000 1er liq del año
''       nivelacion con version del liquidador
   
'Global Const Version = "6.01"
'Global Const FechaVersion = "05/01/2015" ' FGZ - nivelacion con version del liquidador
''   CAS-28749 - H&A - LIQ - Ganancias - Bug rango de 15000 a 25000 1er liq del año
''       nivelacion con version del liquidador
''Ademas
''   Se agregó update de progreso al 100 cuando termina y no hay error

'Global Const Version = "6.00"
'Global Const FechaVersion = "30/12/2014" ' FGZ - nivelacion con version del liquidador
''   CAS-28749 - H&A - LIQ - Ganancias - Bug rango de 15000 a 25000 1er liq del año
''       nivelacion con version del liquidador

'Global Const Version = "5.99"
'Global Const FechaVersion = "16/12/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.98"
'Global Const FechaVersion = "10/12/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.97"
'Global Const FechaVersion = "01/12/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.96"
'Global Const FechaVersion = "11/11/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.95"
'Global Const FechaVersion = "24/10/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.94"
'Global Const FechaVersion = "23/10/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.93"
'Global Const FechaVersion = "22/10/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.92"
'Global Const FechaVersion = "15/10/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.91"
'Global Const FechaVersion = "02/10/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.90"
'Global Const FechaVersion = "30/09/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.89"
'Global Const FechaVersion = "19/09/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.88"
'Global Const FechaVersion = "08/09/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.87"
'Global Const FechaVersion = "15/08/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.86"
'Global Const FechaVersion = "30/07/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.85"
'Global Const FechaVersion = "18/07/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.84"
'Global Const FechaVersion = "16/07/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.83"
'Global Const FechaVersion = "16/07/2014" ' FGZ - nivelacion con version del liquidador


'Global Const Version = "5.82"
'Global Const FechaVersion = "16/07/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.81"
'Global Const FechaVersion = "10/07/2014" ' EAM - nivelacion con version del liquidador

'Global Const Version = "5.80"
'Global Const FechaVersion = "08/07/2014" ' EAM - nivelacion con version del liquidador

'Global Const Version = "5.79"
'Global Const FechaVersion = "26/06/2014" ' EAM - nivelacion con version del liquidador

'Global Const Version = "5.78"
'Global Const FechaVersion = "26/05/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.77"
'Global Const FechaVersion = "19/05/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.76"
'Global Const FechaVersion = "14/05/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.75"
'Global Const FechaVersion = "12/05/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.74"
'Global Const FechaVersion = "30/04/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.73"
'Global Const FechaVersion = "15/04/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.72"
'Global Const FechaVersion = "09/04/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.71"
'Global Const FechaVersion = "09/04/2014" ' FGZ - nivelacion con version del liquidador

'Global Const Version = "5.70"
'Global Const FechaVersion = "08/04/2014" ' FGZ - nivelacion con version del liquidador


'Global Const Version = "5.69"
'Global Const FechaVersion = "04/04/2014" ' FGZ - nivelacion con version del liquidador


'Global Const Version = "5.68"
'Global Const FechaVersion = "31/03/2014" ' FGZ - nivelacion con version del liquidador

' ------ NIVELACION DE VERSIONES CON LIQUIDADOR de 5.46 a 5.68
' ....................

'Global Const Version = "5.45"
'Global Const FechaVersion = "18/09/2013" ' FGZ - no hubo cambios
'                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.44"
'Global Const FechaVersion = "02/09/2013" ' EAM No hubo cambios. Se genera version para nivelar con liquidador

'Global Const Version = "5.43"
'Global Const FechaVersion = "12/08/2013" ' EAM No hubo cambios. Se genera verions para nivelar con liquidador

'Global Const Version = "5.42"
'Global Const FechaVersion = "09/08/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.41"
'Global Const FechaVersion = "05/08/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador


'Global Const Version = "5.40"
'Global Const FechaVersion = "11/07/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.39"
'Global Const FechaVersion = "26/06/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador


'Global Const Version = "5.38"
'Global Const FechaVersion = "07/06/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador


'Global Const Version = "5.37"
'Global Const FechaVersion = "15/03/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.36"
'Global Const FechaVersion = "05/03/2013"   'EAM - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.35"
'Global Const FechaVersion = "26/02/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.34"
'Global Const FechaVersion = "13/02/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.33"
'Global Const FechaVersion = "15/01/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.32"
'Global Const FechaVersion = "08/01/2013"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.31"
'Global Const FechaVersion = "18/12/2012"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.30"
'Global Const FechaVersion = "30/11/2012"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.29"
'Global Const FechaVersion = "23/11/2012"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.28"
'Global Const FechaVersion = "20/11/2012"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.27"
'Global Const FechaVersion = "07/11/2012"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador

'Global Const Version = "5.26"
'Global Const FechaVersion = "05/11/2012"   'FGZ - no hubo cambios
''                                      Nivelacion con ultima version del liquidador y simulador


'Global Const Version = "5.25"
'Global Const FechaVersion = "29/10/2012"   'FGZ - no hubo cambios
''                                      Archivo de log: antes lo creaba en la reaiz del directorio de logs, ahora lo crea en la carpeta de usuario
''                                      Nivelacion con ultima version del liquidador y simulador


'Global Const Version = "5.24"
'Global Const FechaVersion = "24/10/2012"   'FGZ - no hubo cambios
''                                      Archivo de log: antes lo creaba en la reaiz del directorio de logs, ahora lo crea en la carpeta de usuario
''                                      Nivelacion con ultima version del liquidador y simulador
'==================================================================================================================
'

'Global Const Version = "1.01"
'Global Const FechaVersion = "02/08/2012"
''Autor = Lisandor Moro
'' Modificacion:Se extendio el timeout

'Global Const Version = "1.00"
'Global Const FechaVersion = "10/07/2012"
'Autor = Sebastian Stremel

'==================================================================================================================


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
Case 371: 'Borrado de Simulador
    If Version >= "5.24" Then
'        'Tabla nueva
'
'        'CREATE TABLE [dbo].[novretro](
'        '      [nretronro] [int] IDENTITY(1,1) NOT NULL,
'        '      [ternro] [int] NOT NULL,
'        '      [concnro] [int] NOT NULL,
'        '      [nretromonto] [decimal](19, 4) NULL,
'        '      [nretrocant] [int] NULL,
'        '      [pliqnro] [int] NULL,
'        '      [simpronro] [int] NULL,
'        '      [pronro] [int] NULL,
'        '      [pronropago] [int] NULL
'        ') ON [PRIMARY]
'
'        Texto = "Revisar que exista tabla novretro y su estructura sea correcta."
'
'        StrSql = "Select nretronro,ternro,concnro,nretromonto,nretrocant,pliqnro,simpronro,pronro,pronropago FROM novretro WHERE ternro = 1"
'        OpenRecordset StrSql, rs
        
        V = True
    End If

    If Version >= "6.41" Then
        'Campo Nuevo
        'venc,vacnro
        'ALTER TABLE [vacvendidos] ADD [venc] [int] NOT  NULL
        'ALTER TABLE [vacvendidos] ADD [vacnro] [int] NOT  NULL
                
        Texto = "Revisar tabla sim_vacvendidos"
        StrSql = "select venc,vacnro from sim_vacvendidos WHERE vacvendidosnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If


'Cambio en busqueda de embargos bus_embargos
Case Else:
    Texto = "version correcta"
    V = True
End Select


    ValidarV = V
Exit Function

ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function


