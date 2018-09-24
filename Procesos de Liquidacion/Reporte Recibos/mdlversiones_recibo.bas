Attribute VB_Name = "mdlversiones_recibo"
Option Explicit

'Global Const Version = "2.01"
'Global Const FechaModificacion = "28/03/2006"
'Global Const UltimaModificacion = " " '

'Global Const Version = "2.02"
'Global Const FechaModificacion = "17/04/2006"

'Global Const UltimaModificacion = " " '

'Global Const Version = "2.03"
'Global Const FechaModificacion = "19/04/2006"
'Global Const UltimaModificacion = " " '

'Global Const Version = "2.04"
'Global Const FechaModificacion = "20/04/2006"
'Global Const UltimaModificacion = " " 'La 1 columna del confrep se puede configurar como un concepto o un acumulador (Sueldo).

'Global Const Version = "2.05"
'Global Const FechaModificacion = "21/04/2006"
'Global Const UltimaModificacion = " " 'La 1 columna del confrep se puede configurar como un concepto o un acumulador (Sueldo).

'Global Const Version = "2.06"
'Global Const FechaModificacion = "24/04/2006"
'Global Const UltimaModificacion = " " 'Se agrego un nuevo recibo de Sueldo 42.

'Global Const Version = "2.07"
'Global Const FechaModificacion = "02/05/2006"
'Global Const UltimaModificacion = " " 'Se agrego un nuevo recibo de Sueldo 43 - BNP Paribas Argentina.

'Global Const Version = "2.08"
'Global Const FechaModificacion = "10/05/2006"
'Global Const UltimaModificacion = " " 'Se agrego un nuevo recibo de Sueldo 44 - Liggett

'Global Const Version = "2.09"
'Global Const FechaModificacion = "11/05/2006"
'Global Const UltimaModificacion = " " 'Se agrego un nuevo recibo de Sueldo 45 - REDEPA.

'Global Const Version = "2.10"
'Global Const FechaModificacion = "19/05/2006"
'Global Const UltimaModificacion = " " 'Se agrego un nuevo recibo de Sueldo 46 - Distribuidora Metropolitana S.R.L

'Global Const Version = "2.11"
'Global Const FechaModificacion = "29/05/2006"
'Global Const UltimaModificacion = " " 'Se agrego un nuevo recibo de Sueldo 48 - Codere

'Global Const Version = "2.12"
'Global Const FechaModificacion = "31/05/2006"
'Global Const UltimaModificacion = " " 'Se agrego un nuevo recibo de Sueldo 49 - Tetra Nueva version

'Global Const Version = "2.13"
'Global Const FechaModificacion = "01/06/2006"
'Global Const UltimaModificacion = " " 'Se modifico el recibo 41 paolini: fecha alta reconocidad y alta

'Global Const Version = "2.14"
'Global Const FechaModificacion = "03/06/2006"
'Global Const UltimaModificacion = " " 'Se agrego un nuevo recibo de Sueldo 50 - Divino - Uruguay - nueva version

'Global Const Version = "2.15"
'Global Const FechaModificacion = "07/06/2006"
'Global Const UltimaModificacion = " " 'Se agrego un nuevo recibo de Sueldo 51 - Darmex

'Global Const Version = "2.16"
'Global Const FechaModificacion = "14/06/2006"
'Global Const UltimaModificacion = " " 'Retoques al recibo de Sueldo 13 - Halliburton

'Global Const Version = "2.17"
'Global Const FechaModificacion = "20/06/2006"
'Global Const UltimaModificacion = " " 'Retoques al recibo de Sueldo 12 - IHSA

'Global Const Version = "2.17"
'Global Const FechaModificacion = "20/06/2006"
'Global Const UltimaModificacion = " " 'Retoques al recibo de Sueldo 12 - IHSA, Agregado de logs

'Global Const Version = "2.18"
'Global Const FechaModificacion = "04/07/2006"
'Global Const UltimaModificacion = " " 'FAF - Se agrego un nuevo recibo de Sueldo 52 - Aadi-Capif

'Global Const Version = "2.19"
'Global Const FechaModificacion = "05/07/2006"
'Global Const UltimaModificacion = " " 'NF - Modificaciones a recibo de Sueldo 51 - Darmex

'Global Const Version = "2.20"
'Global Const FechaModificacion = "06/07/2006"
'Global Const UltimaModificacion = " " 'FAF - Modificaciones a recibo de Sueldo 39 - Marsh

'Global Const Version = "2.21"
'Global Const FechaModificacion = "11/07/2006"
'Global Const UltimaModificacion = " " 'FAF - Modificaciones a recibo de Sueldo 42 - Megatone
                                      ' Se agrego mid(direccion, 1, 100) y mid(localidad, 1, 100) de todos los modelos
                                      ' Se agrego la ultima sql ejecutada en el log de todos los modelos
'Global Const Version = "2.22"
'Global Const FechaModificacion = "14/07/2006"
'Global Const UltimaModificacion = " " 'FAF - Modificaciones al recibo de Sueldo 39 - Marsh -
                                      ' La fecha de baja sale de la ultima fase marcada como real.
'Global Const Version = "2.23"
'Global Const FechaModificacion = "17/07/2006"
'Global Const UltimaModificacion = " " 'FAF - Modificaciones al recibo de Sueldo 40 - BIA -
'                                      ' Se cambio el tipo de estructura para Departamento. De Categoria a Departamento.

'Global Const Version = "2.24"
'Global Const FechaModificacion = "17/07/2006"
'Global Const UltimaModificacion = " " 'FGZ - Modificaciones en sub main:
                                      ' Cint() por Clng()
                                      ' mensajes de log con cada parametro

'Global Const Version = "2.25"
'Global Const FechaModificacion = "01/08/2006"
'Global Const UltimaModificacion = " " 'LA. - Modificaciones al recibo de Sueldo 50 - Divino - Uruguay - Se le saco la coma a los numeros decimales, se hicieron chequeos de variables y vhora se calcula con info de col44 en lugar de 43

'Global Const Version = "2.26"
'Global Const FechaModificacion = "04/08/2006"
'Global Const UltimaModificacion = " " 'LA. - Modificaciones al recibo de Sueldo 50 - Divino - Uruguay - Los Tipos de Documento  RUC y BPS se pasan a traves del confrep y se saacaron los tipos de documentos MISS y BSE

'Global Const Version = "2.27"
'Global Const FechaModificacion = "09/08/2006"
'Global Const UltimaModificacion = " " 'FAF - Modificaciones al recibo de Sueldo 52 - Se cambio la fecha de pago por la fecha planeada

'Global Const Version = "2.28"
'Global Const FechaModificacion = "14/08/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro en el modelo 37 telearte se busco la desc del periodo

'Global Const Version = "2.29"
'Global Const FechaModificacion = "17/08/2006"
'Global Const UltimaModificacion = " " 'FGZ - Modificaciones al recibo de Sueldo 20 - DTT" -
'                                      ' ahora controlo por nulo o vacio - If Not EsNulo(rsConsult!ctabnro) Then

'Global Const Version = "2.30"
'Global Const FechaModificacion = "17/08/2006"
'Global Const UltimaModificacion = " " 'HDS - Se agregaron Flog.writeline para verificar contenido
                                      'de la variable CuentaBancaria

'Global Const Version = "2.31"
'Global Const FechaModificacion = "18/08/2006"
'Global Const UltimaModificacion = " " 'FAF - Se agregaron 3 columnas en el confrep 90 91 92
                                      
'Global Const Version = "2.32"
'Global Const FechaModificacion = "23/08/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se creo el modelo 53 para mercado Central

'Global Const Version = "2.33"
'Global Const FechaModificacion = "15/09/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se creo el modelo 54 para Union de Paris

'Global Const Version = "2.34"
'Global Const FechaModificacion = "13/09/2006"
'Global Const UltimaModificacion = " " 'Leticia A. - Se modifico Recibo 50 - Divino - Uruguay - se especificaron las cols 42-43-44 para tickets-v. hora - sueldo

'Global Const Version = "2.35"
'Global Const FechaModificacion = "26/09/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Recibo 54 - Buscar la direccion del empleado

'Global Const Version = "2.36"
'Global Const FechaModificacion = "27/09/2006"
'Global Const UltimaModificacion = " " 'FAF - Se creo el modelo 55 para CCU
'                                      'FAF - Se creo el modelo 56 para Finca

'Global Const Version = "2.37"
'Global Const FechaModificacion = "02/10/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Modelo 01 - Agregue en direccion de la empresa piso y dpto
                                      
'Global Const Version = "2.38"
'Global Const FechaModificacion = "05/10/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - modelo 54 - Agregue en direccion de la empresa piso y dpto
                                      
'Global Const Version = "2.39"
'Global Const FechaModificacion = "18/10/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - modelo 19 - Agregue en direccion de la empresa piso y dpto
                                      
'Global Const Version = "2.40"
'Global Const FechaModificacion = "19/10/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se creo Modelo de 57 para Latin3

'Global Const Version = "2.41"
'Global Const FechaModificacion = "24/10/2006"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se creo Modelo de 58 para horvath
                                      
'Global Const Version = "2.42"
'Global Const FechaModificacion = "26/10/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se creo Modelo de 59 para Praxair a partir del 50
                                      
'Global Const Version = "2.43"
'Global Const FechaModificacion = "03/11/2006"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se creo Modelo de 60 para gedas
                                      
'Global Const Version = "2.44"
'Global Const FechaModificacion = "05/11/2006"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se creo Modelo de 61 para AGD
                                      
'Global Const Version = "2.45"
'Global Const FechaModificacion = "17/11/2006"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se modifico el modelo de 56. Calculo del Sueldo Basico.
                                      
'Global Const Version = "2.46"
'Global Const FechaModificacion = "20/11/2006"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se modifico el modelo 56. No salia el Banco de Pago.
                                      
'Global Const Version = "2.47"
'Global Const FechaModificacion = "23/11/2006"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se agrego el modelo 62 para A.C.A.R.A.
                                      
'Global Const Version = "2.48"
'Global Const FechaModificacion = "23/11/2006"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se modifico el modelo 46. Se agrego OS elegida y el Nro de recibo (cliqnro).
                                      
'Global Const Version = "2.49"
'Global Const FechaModificacion = "12/12/2006"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se agrego el modelo 63 para EKI.

'Global Const Version = "2.50"
'Global Const FechaModificacion = "15/01/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se agrego el modelo 64 para Gorina.

'Global Const Version = "2.51"
'Global Const FechaModificacion = "14/02/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se agrego el modelo 65 para APEX.

'Global Const Version = "2.52"
'Global Const FechaModificacion = "06/02/2006"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se agrego el modelo 66 para Teleperformance Chile.

'Global Const Version = "2.53"
'Global Const FechaModificacion = "21/03/2007"
'Global Const UltimaModificacion = " " 'Martin Ferraro - cambios en el modelo 66 de Teleperformance Chile.

'Global Const Version = "2.54"
'Global Const FechaModificacion = "22/03/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se agrego el modelo 67 para Praxair Argentina.

'Global Const Version = "2.55"
'Global Const FechaModificacion = "28/03/2007"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se agrego la busqueda de la forma de pago en el modelo 62 de ACARA.

'Global Const Version = "2.56"
'Global Const FechaModificacion = "29/03/2007"
'Global Const UltimaModificacion = " " 'N. Trillo - Se modificó el modelo 60 de GEDAS.

'Global Const Version = "2.57"
'Global Const FechaModificacion = "03/04/2007"
'Global Const UltimaModificacion = " " 'Martin Ferraro - en el modelo 62 del recibo de acara se cambio la forma de pago.

'Global Const Version = "2.58"
'Global Const FechaModificacion = "13/04/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se agrego el modelo 68 para PHA.

'Global Const Version = "2.59"
'Global Const FechaModificacion = "17/04/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se agrego el modelo 69 para San Martin de Tabacal.


'Global Const Version = "2.60"
'Global Const FechaModificacion = "27/04/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se modifico el recibo 59 de praxair Uruguay.
                                                     'Se Agrego Grupo y SubGrupo de actividad.
'Global Const Version = "2.61"
'Global Const FechaModificacion = "18/05/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se modifico el recibo 61 de AGD.
                                      'Se agregaron columnas al confrep. Se cambio fecha alta.
                                      
'Global Const Version = "2.62"
'Global Const FechaModificacion = "31/05/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se agrego el recibo 70 para Provincia Seguros.

'Global Const Version = "2.63"
'Global Const FechaModificacion = "04/06/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se modifico el proceso 18 de Estrada.
                                      'Se agrego fecha de alta reconocida
                                      
'Global Const Version = "2.64"
'Global Const FechaModificacion = "12/06/2007"
'Global Const UltimaModificacion = " " 'G. Bauer - N. Trillo - Se agrego la consulta para que vaya a buscar la firma de la empresa y salga en el recibo modelo 41 para Paolini.

'Global Const Version = "2.65"
'Global Const FechaModificacion = "05/07/2007"
'Global Const UltimaModificacion = " " 'G. Bauer -se cambio la definicion de cabliqNumero de entero a long en el recibo numero 19 de IMSA.

'Global Const Version = "2.66"
'Global Const FechaModificacion = "19/07/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se creo el modelo 71 para Alta plastica.
'
'Global Const Version = "2.67"
'Global Const FechaModificacion = "24/07/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se creo el modelo 72 para Zarcam.



'Global Const Version = "2.68"
'Global Const FechaModificacion = "11/09/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se creo el modelo 73 para Papelbril.
'                                                     'Se modifico el modelo 61 de AGD. Se cambio los campos: Condicion,Lugar de pago,Planta

'Global Const Version = "2.69"
'Global Const FechaModificacion = "21/09/2007"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se modifico la busqueda de la forma bancaria para el modelo 2
  
'Global Const Version = "2.70"
'Global Const FechaModificacion = "27/09/2007"
'Global Const UltimaModificacion = " " ' 27-09-2007 - Diego Rosso - Categoria tomar del tipo de estructura 3.

'Global Const Version = "2.71"
'Global Const FechaModificacion = "28/09/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se modifico el modelo 61 de AGD. Se cambio la consulta de la forma de pago
                                                     'Se cambio la descripcion de Condicion. Se agrego Codigo Postal en empresa.

'Global Const Version = "2.72"
'Global Const FechaModificacion = "05/10/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se modifico el modelo 61 de AGD. Se cambio la forma de pago (OTRA VEZ!)
                                                     
'Global Const Version = "2.73"
'Global Const FechaModificacion = "12/10/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se modifico el modelo 55 de CCU. Se cambio el tipo de domicilio de la direccion de la sucursal. De 10 por 5
                                                       'Se modifico el modelo 56 de Finca. Se agrego la busqueda del Contrato al que pertenece el empleado.
                                                       'Se modifico el modelo 56 de Finca. Se cambio el tipo de domicilio de la direccion de la sucursal. De 10 por 5

'Global Const Version = "2.74"
'Global Const FechaModificacion = "23/10/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se modifico el modelo 55 de CCU. La fecha de alta es de la primer fase.
                                                       'La antiguedad sale del concepto configurado en ela columna 44 del confrep

'Global Const Version = "2.75"
'Global Const FechaModificacion = "26/10/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se modifico el modelo 58 de Horwath.
                                       'Se cambio la forma de pago

'Global Const Version = "2.76"
'Global Const FechaModificacion = "30/10/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se modifico el modelo 56 de FLC (Finca La Celina). La fecha de alta es de la primer fase.
                                                       'La antiguedad sale del concepto configurado en la columna 44 del confrep

'Global Const Version = "2.77"
'Global Const FechaModificacion = "02/11/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se creo el modelo 74 para Praxair Chile.

'Global Const Version = "2.78"
'Global Const FechaModificacion = "06/11/2007"
'Global Const UltimaModificacion = " " 'Gustavo Ring - En el modelo 69 (Tabacal) se cambio el codigo externo del sector por la descripción abreviada del mismo.

'Global Const Version = "2.79"
'Global Const FechaModificacion = "09/11/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se modifico el modelo 56 de FLC (Finca La Celina). El básico sale del concepto o acumulador definico en la columna 1

'Global Const Version = "2.80"
'Global Const FechaModificacion = "30/11/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se modifico el modelo 55 de CCU (Cicsa). El básico sale del concepto o acumulador definico en la columna 1

'Global Const Version = "2.81"
'Global Const FechaModificacion = "15/01/2008"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se creo el modelo 75 para Arlei.

'Global Const Version = "2.82"
'Global Const FechaModificacion = "18/01/2008"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se creo el modelo 76 para Santana.

'Global Const Version = "2.83"
'Global Const FechaModificacion = "24/01/2008"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se creo el modelo 77 para Repsa.

'Global Const Version = "2.84"
'Global Const FechaModificacion = "17/04/2008"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Modificaciones varias del modelo 35

'Global Const Version = "2.85"
'Global Const FechaModificacion = "28/04/2008"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se declaro la varible generoRecibo como global. Esta
                                      'esta variable se encargaba de incrementar el orden si el se generaba el recibo.
                                      'El control lo hacia en el main. Pero hay veces que el proceso de generacion aborta
                                      'y la variable generoRecibo ya habia sido actualizada dando errores en el orden lo
                                      'cual implicaba errores en la barra de navegacion de recibos.
                                      'Se corrigio el uso de la misma en el recibo 35

'Global Const Version = "2.86"
'Global Const FechaModificacion = "07/05/2008"
'Global Const UltimaModificacion = " " 'Martin Ferraro - modelo 35, si el basico es cero entonces buscarlo en la remuneracion

'Global Const Version = "2.87"
'Global Const FechaModificacion = "03/06/2008"
'Global Const UltimaModificacion = " " 'Martin Ferraro - modelo 35, si el basico es cero entonces buscarlo en la remuneracion

'Global Const Version = "2.88"
'Global Const FechaModificacion = "26/08/2008"
'Global Const UltimaModificacion = " " 'Martin Ferraro - se creo el modelo 78 para AMIA

'Global Const Version = "2.89"
'Global Const FechaModificacion = "26/08/2008"
'Global Const UltimaModificacion = " Martin Ferraro - Cambios para el modelo 78"

'Global Const Version = "2.90"
'Global Const FechaModificacion = "09/10/2008"
'Global Const UltimaModificacion = " Martin Ferraro - Cambios para el modelo 70 porque buscaba mal firma"

'Global Const Version = "2.91"
'Global Const FechaModificacion = "22/10/2008"
'Global Const UltimaModificacion = " Martin Ferraro - Cambios para el modelo 70 - Fecha de alta = a la del tablero del empleado"

'Global Const Version = "2.92"
'Global Const FechaModificacion = "13/11/2008"
'Global Const UltimaModificacion = " Stankunas Cesar - Se creo el modelo 79 para Vittal"

'Global Const Version = "2.93"
'Global Const FechaModificacion = "19/11/2008"
'Global Const UltimaModificacion = " Diego Nuñez - Se generó el modelo 80 para la generación de recibos de sueldo de DTT II - RIM"

'Global Const Version = "2.94"
'Global Const FechaModificacion = "04/12/2008"
'Global Const UltimaModificacion = " Diego Nuñez - Se modificó el modelo 71 para la generación de recibos de sueldo de Alta Plástica - Fecha de alta reconocida"

'Global Const Version = "2.95"
'Global Const FechaModificacion = "22/12/2008"
'Global Const UltimaModificacion = "Martin Ferraro - Cambios en el modelo 23"

'Global Const Version = "2.96"
'Global Const FechaModificacion = "13/01/2009"
'Global Const UltimaModificacion = "Lisandro Moro - Se creo el modelo 81 - Santillana"

'Global Const Version = "2.97"
'Global Const FechaModificacion = "20/01/2009"
'Global Const UltimaModificacion = "Fernando Favre - Se modificó el modelo 18 - Angel Estrada"

'Global Const Version = "2.98"
'Global Const FechaModificacion = "20/01/2009"
'Global Const UltimaModificacion = "Hatsembiller Octavio - Se creo el modelo 82 - Action Line"

'Global Const Version = "2.99"
'Global Const FechaModificacion = "20/01/2009"
'Global Const UltimaModificacion = "Diego Muñez - Se modificó la Subrutina generarDatosRecibo02 - Teimaiken"

'Global Const Version = "3.00"
'Global Const FechaModificacion = "21/05/2009"
'Global Const UltimaModificacion = "Stankunas Cesar - Se corrigió la Subrutina generarDatosRecibo02 - Temaiken"


'Global Const Version = "3.00"
'Global Const FechaModificacion = "03/06/2009"
'Global Const UltimaModificacion = "Hatsembiller octavio - Se corrigió la Subrutina generarDatosRecibo02 - Temaiken"


'Global Const Version = "3.01"
'Global Const FechaModificacion = "03/06/2009"
'Global Const UltimaModificacion = "Hatsembiller Octavio - Se creo el modelo 83 - Administradora Sanatorial"

'Global Const Version = "3.02"
'Global Const FechaModificacion = "03/06/2009"
'Global Const UltimaModificacion = "Hatsembiller Octavio - Se creo el modelo 84 - Action Line"

'Global Const Version = "3.03"
'Global Const FechaModificacion = "03/06/2009"
'Global Const UltimaModificacion = "Hatsembiller Octavio - Se creo el modelo 85 - TPS"
' Se modifico los datos que componen el domicilio de la empresa

'Global Const Version = "3.04"
'Global Const FechaModificacion = "31/07/2009"
'Global Const UltimaModificacion = "Martin Ferraro - Se acomodo el log para Encriptacion de string connection"

'Global Const Version = "3.05"
'Global Const FechaModificacion = "30/09/2009"
'Global Const UltimaModificacion = "Javier Irastorza - Se creo el modelo 86 - SuperCanal"

'Global Const Version = "3.06"
'Global Const FechaModificacion = "08/09/2009"
'Global Const UltimaModificacion = "Hatsembiller Octavio - Corrección de Obra Social Admistradora sanatorial"

'Global Const Version = "3.07"
'Global Const FechaModificacion = "08/09/2009"
'Global Const UltimaModificacion = "Hatsembiller Octavio - Horas caluculadas"


'Global Const Version = "3.08"
'Global Const FechaModificacion = "08/09/2009"
'Global Const UltimaModificacion = "Hatsembiller Octavio -  Se creo el modelo 12 - Grupo Cargo"

'Global Const Version = "3.09"
'Global Const FechaModificacion = "21/12/2009"
'Global Const UltimaModificacion = "Elizabeth Gisela Oviedo- Modificación Subrutina generarDatosRecibo08 - Norton - Guarde Luar y Ubic. Geo"

'Global Const Version = "3.10"
'Global Const FechaModificacion = "12/02/2010"
'Global Const UltimaModificacion = "Cesar Stankunas - Se corrigió el modelo 83 (la antigüedad reconocida del empleado)- Administradora Sanatorial"

'Global Const Version = "3.11"
'Global Const FechaModificacion = "22/03/2010"
'Global Const UltimaModificacion = "Cesar Stankunas - Se creó el modelo 88 - Medicus"

'Global Const Version = "3.12"
'Global Const FechaModificacion = "30/03/2010"
'Global Const UltimaModificacion = "Cesar Stankunas - Se creó el modelo 89 - Price"

'Global Const Version = "3.13"
'Global Const FechaModificacion = "12/04/2010"
'Global Const UltimaModificacion = "Fernando Favre - Se creó el modelo 90 - Radio Continental a partir del modelo 2"

'Global Const Version = "3.14"
'Global Const FechaModificacion = "20/04/2010"
'Global Const UltimaModificacion = "Cesar Stankunas - Se modificó el modelo 89 (Price) - Cambio en la forma de buscar la anigüedad"

'Global Const Version = "3.15"
'Global Const FechaModificacion = "12/05/2010"
'Global Const UltimaModificacion = "Cesar Stankunas - Se modificó el modelo 89 (Price) - Se modificó la forma de comparar la fecha de alta con la fecha actual"

'Global Const Version = "3.16"
'Global Const FechaModificacion = "18/05/2010"
'Global Const UltimaModificacion = "Stankunas Cesar - Se creó el modelo 91 - Grupo Cargo a partir del modelo 35"

'Global Const Version = "3.17"
'Global Const FechaModificacion = "28/05/2010"
'Global Const UltimaModificacion = "Lisandro Moro - Se creó el modelo 92 - Multivoice"

'Global Const Version = "3.18"
'Global Const FechaModificacion = "28/05/2010"
'Global Const UltimaModificacion = "Stankunas Cesar - Se modificó el modelo 83 (ASM)"

'Global Const Version = "3.19"
'Global Const FechaModificacion = "16/07/2010"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se modificó el modelo 39 (MARSH)"

'Global Const Version = "3.20"
'Global Const FechaModificacion = "19/07/2010"
'Global Const UltimaModificacion = "Stankunas Cesar - Se modificó el modelo 88 (Medicus)"
'                                  'Se quitó la condición para los empleados con estructura 154 (Trabajo Eventual) en
'                                  'la búsqueda de Contrato

'Global Const Version = "3.21"
'Global Const FechaModificacion = "28/07/2010"
'Global Const UltimaModificacion = "Stankunas Cesar - Se modificó el modelo 91 (Grupo Cargo)"
                                  'Se agregó una columna configurable que guarda el monto de un acumulador para mostrarlo en el recibo.

'Global Const Version = "3.22"
'Global Const FechaModificacion = "29/07/2010"
'Global Const UltimaModificacion = "Fernando Favre - Se creó el modelo 93 (Golden Peanut)"

'Global Const Version = "3.23"
'Global Const FechaModificacion = "30/07/2010"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se creó el modelo 94 - Teletech"

'Global Const Version = "3.24"
'Global Const FechaModificacion = "11/08/2010"
'Global Const UltimaModificacion = "Stankunas Cesar - Se modificó el modelo 83 - ASM"

'Global Const Version = "3.25"
'Global Const FechaModificacion = "09/09/2010"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se modificó el modelo 94 - Teletech"

'Global Const Version = "3.26"
'Global Const FechaModificacion = "23/09/2010"
'Global Const UltimaModificacion = "Stankunas Cesar - Se modificó el modelo 89 - Price"

'Global Const Version = "3.27"
'Global Const FechaModificacion = "18/10/2010"
'Global Const UltimaModificacion = "Stankunas Cesar - Se creó el 95 - Tremac"

'Global Const Version = "3.28"
'Global Const FechaModificacion = "19/10/2010"
'Global Const UltimaModificacion = "Stankunas Cesar - Se creó el 96 - Monresa"

'Global Const Version = "3.29"
'Global Const FechaModificacion = "22/10/2010"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se creó el modelo 97 - Multivoice Colombia"

'Global Const Version = "3.30"
'Global Const FechaModificacion = "02/11/2010"
'Global Const UltimaModificacion = "Manuel López - Nivelación, ya que se había perdido un cambio en el modelo 93."

'Global Const Version = "3.31"
'Global Const FechaModificacion = "12/11/2010"
'Global Const UltimaModificacion = "Stankunas Cesar - Se modificó el modelo 89 - Price"

'Global Const Version = "3.32"
'Global Const FechaModificacion = "09/12/2010"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se modificó el modelo 94 - Teletech"

'Global Const Version = "3.33"
'Global Const FechaModificacion = "30/12/2010"
'Global Const UltimaModificacion = "Fernando Favre - Se creó el modelo 98 - Freddo"

'Global Const Version = "3.34"
'Global Const FechaModificacion = "10/01/2011"
'Global Const UltimaModificacion = "Verónica Bogado - Modelo 99 para Fundición San Cayetano"

'Global Const Version = "3.35"
'Global Const FechaModificacion = "26/01/2011"
'Global Const UltimaModificacion = "Fernando Favre - Se modifico modelo 98. El 'centro de resultado' se muestra la estructura Centro Costo. Antes se mostraba la categoria"

'Global Const Version = "3.36"
'Global Const FechaModificacion = "03/03/2011"
'Global Const UltimaModificacion = "Verónica Bogado - Se agregó el modelo 100 para Sidersa."

'Global Const Version = "3.37"
'Global Const FechaModificacion = "01/03/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se agregó Obra Social en Bapro."

'Global Const Version = "3.38"
'Global Const FechaModificacion = "02/03/2011"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se modificó el modelo 86 (Supercanal) para registrar la antiguedad del empleado en el campo auxchar de la tabla rep_recibo"

'Global Const Version = "3.39"
'Global Const FechaModificacion = "18/03/2011"
'Global Const UltimaModificacion = "Stankunas Cesar - Se modificó el modelo 96 (Monresa) - Se cambiaron los campos de Grupo y Subgrupo de Actividad"

'Global Const Version = "3.40"
'Global Const FechaModificacion = "15/03/2011"
'Global Const UltimaModificacion = "Fernando Favre - Se modifico el modelo 1 - Standard (Bapro)."

'Global Const Version = "3.41"
'Global Const FechaModificacion = "06/04/2011"
'Global Const UltimaModificacion = "Stankunas Cesar - Se modifico el modelo 98 - Freddo"

'Global Const Version = "3.42"
'Global Const FechaModificacion = "06/04/2011"
'Global Const UltimaModificacion = "Fernando Favre - Se modifico el modelo 40 - BIA. Antiguedad reconocida de la estructura 44 a la 39"

'Global Const Version = "3.43"
'Global Const FechaModificacion = "06/04/2011"
'Global Const UltimaModificacion = "Verónica Bogado - Se eliminó el truncamiento de categoría para elmodelo 100"

'Global Const Version = "3.44"
'Global Const FechaModificacion = "18/04/2011"
'Global Const UltimaModificacion = "Verónica Bogado - Se genera el modelo 101 para Mundo Maipú"

'Global Const Version = "3.45"
'Global Const FechaModificacion = "29/04/2011"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se genera el modelo 102 para Sykes"

'Global Const Version = "3.46"
'Global Const FechaModificacion = "13/05/2011"
'Global Const UltimaModificacion = "Verónica Bogado - Se agregó una tercera opción al campo configurable, col44"

'Global Const Version = "3.47"
'Global Const FechaModificacion = "20/05/2011"
'Global Const UltimaModificacion = "Stankunas Cesar - Se agregó el modelo 103 para MAE" 'Se hace en un Módulo aparte (repRecibos2) ya que en este no hay espacio

'Global Const Version = "3.48"
'Global Const FechaModificacion = "30/05/2011"
'Global Const UltimaModificacion = "Verónica Bogado - Se agregó el modelo 104 para droguería Suizo Argentina" 'En el modulo secundario creado para el modelo anterior

'Global Const Version = "3.49"
'Global Const FechaModificacion = "13/06/2011"
'Global Const UltimaModificacion = "Fernando Favre - Se agregó el modelo 105 para Merck Sharp Down" 'En repRecibo2

'Global Const Version = "3.50"
'Global Const FechaModificacion = "21/06/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se agrego el modelo 106 para Farmografica 'En repRecibo2"

'Global Const Version = "3.51"
'Global Const FechaModificacion = "28/06/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico el modelo 62, estaba comentado la parte del logo, ahora muestra el logo en Vision - Hospital Britanico"

'Global Const Version = "3.52"
'Global Const FechaModificacion = "05/07/2011"
'Global Const UltimaModificacion = "Stankunas Cesar - Se agregó el modelo 107 para INDAP (Chile)"

'Global Const Version = "3.53"
'Global Const FechaModificacion = "01/07/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico en el modelo 1, la variable numero de cuenta y se le puso el valor 0"

'Global Const Version = "3.54"
'Global Const FechaModificacion = "08/07/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico el modelo 101 Mundo Maipu porque no guardaba un campo en auxchar4"

'Global Const Version = "3.55"
'Global Const FechaModificacion = "13/07/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico modelo 106 para que muestre Basico - Antiguedad - A Cuentas Futuras - Bruto"

'Global Const Version = "3.56"
'Global Const FechaModificacion = "21/07/2011"
'Global Const UltimaModificacion = "FGZ - Se modificó el modelo 107 para INDAP (Chile) "
'                                   Ademas
'                                   Se hicieron varios cambios generales
'                                      Se creo un modulo de versiones para el proceso
'                                      Se cambiaron de lugar algunas definiciones de variables globales que se habian definido en reprecibos2.bas
'                                      La activicion del manejador de errores CE estaba mal colocada
'                                         si hay un error antes de crear los archivos de log ==> daba error el manejador de error)
'
'

'Global Const Version = "3.57"
'Global Const FechaModificacion = "13/07/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se creo el Modelo 108 para General Mills"

'Global Const Version = "3.58"
'Global Const FechaModificacion = "28/07/2011"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se creo el Modelo 109 para Merk SD"

'Global Const Version = "3.59"
'Global Const FechaModificacion = "28/07/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modificaron campos para el modelo 62 Hospital Britanico"

'Global Const Version = "3.60"
'Global Const FechaModificacion = "05/08/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se saco el campo auxdeci1, guardaba el nrodecuenta, el asp no lo usa"

'Global Const Version = "3.61"
'Global Const FechaModificacion = "05/08/2011"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se modificó el modelo 97 - Multivoice Colombia"
'                                                         Se quito la condición de si el sueldo del empleado era cero para
'                                                         traer la del acumulador

'Global Const Version = "3.62"
'Global Const FechaModificacion = "09/08/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se creo modelo 110 - Coop. Seguros"

'Global Const Version = "3.63"
'Global Const FechaModificacion = "12/08/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico el modelo 106 Farmografica para que guarde los campos correspondientes"

'Global Const Version = "3.64"
'Global Const FechaModificacion = "15/08/2011"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se modifico el modelo 102 Sykes CR. "
                                  'Ahora se trae el titulo del Sueldo desde el confrep.

'Global Const Version = "3.65"
'Global Const FechaModificacion = "16/08/2011"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se modifico el modelo 109. Recibo adicional para MSD (Merck Sharp & Dohme Arg). "
                                  
'Global Const Version = "3.66"
'Global Const FechaModificacion = "17/08/2011"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se modifico el modelo 105 (Recibo MSD) y el 109 (Recibo adicional para MSD) . "
                                   'Seincorporó un nuevo campo llamado "modelo" a la tabla "rep_recibo" para poder diferenciar el
                                   'modelo con el que se generó dicho recibo.
                                   
'Global Const Version = "3.67"
'Global Const FechaModificacion = "17/08/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico el modelo 106 Farmografica "
                                   'Se corigieron campos para que salgan los datos correspondientes
                                   
'Global Const Version = "3.68"
'Global Const FechaModificacion = "17/08/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico el modelo 108 General Mills "
                                   'Se corigieron campos para que salgan los datos correspondientes
                                  
'Global Const Version = "3.69"
'Global Const FechaModificacion = "17/08/2011"
'Global Const UltimaModificacion = "Brzozowski Juan Pablo - Se modifico el modelo 105 (Recibo MSD)"
           'Se agrego en el filtro donde busca los conceptos, la condicion que el concepto no sea
           'del tipo especificado en la columna 33 del confrep (caso contrario al recibo adic. generado por el modelo 105 MSD)"
                                                                
'Global Const Version = "3.70"
'Global Const FechaModificacion = "23/08/2011"
'Global Const UltimaModificacion = "Sebastian Stremel - Se modifico modelo 76 - se mustra la fecha de alta reconocida "
                                  
'Global Const Version = "3.71"
'Global Const FechaModificacion = "25/08/2011"
'Global Const UltimaModificacion = "FGZ - Se modifico modelo 102 - Sykes "

'Global Const Version = "3.72"
'Global Const FechaModificacion = "01/09/2011"
'Global Const UltimaModificacion = "Carmen Quintero - Se creo modelo 111 - para Portugal"

'Global Const Version = "3.73"
'Global Const FechaModificacion = "12/09/2011"
'Global Const UltimaModificacion = "Gonzalez Nicolás - Se creo modelo 112 - Para Horwath Rosario"

'Global Const Version = "3.74"
'Global Const FechaModificacion = "13/09/2011"
'Global Const UltimaModificacion = "Gonzalez Nicolás - Se creo modelo 113 - Para Met Roma"

'Global Const Version = "3.75"
'Global Const FechaModificacion = "19/09/2011"
'Global Const UltimaModificacion = "Sebastian Stremel - Se creo modelo 114 - Para Deloitte"

'Global Const Version = "3.76"
'Global Const FechaModificacion = "29/09/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico para HB Residentes y Estudiantes "

'Global Const Version = "3.77"
'Global Const FechaModificacion = "05/10/2011"
'Global Const UltimaModificacion = "Dimatz Rafael - Se modifico el modelo 110 para que se puedan configurar 3 Firmas"

'Global Const Version = "3.78"
'Global Const FechaModificacion = "07/10/2011"
'Global Const UltimaModificacion = "Gonzalez Nicolás - Se modifico el modelo 113 - Met Roma | Se agregaron estructura sector y regimen jubilatorio"

'Global Const Version = "3.79"
'Global Const FechaModificacion = "13/10/2011"
'Global Const UltimaModificacion = "Gonzalez Nicolás - Se modifico el modelo 113 - Met Roma | Se agregaron Comentarios para el Log y se corrigio la forma en que traia la fecha de alta y de baja."

'Global Const Version = "3.80"
'Global Const FechaModificacion = "13/10/2011"
'Global Const UltimaModificacion = "Gonzalez Nicolás - Se modifico el modelo 113 - Met Roma | Se corrigió error en recordset de fases"

'Global Const Version = "3.81"
'Global Const FechaModificacion = "19/10/2011"
'Global Const UltimaModificacion = " Matias Dallegro - Se modifico el modelo 27 - Indra que no filtre por tidnro entre 1 y  menor a 5 sino que traiga tidnro menor que tenga el tercero "

'Global Const Version = "3.82"
'Global Const FechaModificacion = "20/10/2011"
'Global Const UltimaModificacion = "FGZ - Se modifico el modelo 113 - Met Roma | Se corrigió error en recordset de fases"

'Global Const Version = "3.83"
'Global Const FechaModificacion = "10/11/2011"
'Global Const UltimaModificacion = " Sebastian Stremel - Se corrigieron errores en el modelo 114 de deloitte chile"

'Global Const Version = "3.84"
'Global Const FechaModificacion = "23/11/2011"
'Global Const UltimaModificacion = " Brzozowski Juan Pablo - Se creo el modelo 115 para generar el recibo Mensual y de Vacaciones para APEX S.A Paraguay. "

'Global Const Version = "3.85"
'Global Const FechaModificacion = "23/11/2011"
'Global Const UltimaModificacion = " Brzozowski Juan Pablo - Se creo el modelo 116 para generar el recibo final por Termino de Contrato para APEX S.A Paraguay"

'Global Const Version = "3.86"
'Global Const FechaModificacion = "29/11/2011"
'Global Const UltimaModificacion = " Fernando Favre - Se creo el modelo 117 para Cardif (Gestion Compartida) a partir del modelo 71"

'Global Const Version = "3.87"
'Global Const FechaModificacion = "30/11/2011"
'Global Const UltimaModificacion = " Carlos Masson - Se creo el modelo 118 para PKF a partir del modelo 89, se agrego funcionalidad para editar la cantidad de digitos a mostrar en el legajo"

'Global Const Version = "3.88"
'Global Const FechaModificacion = "05/12/2011"
'Global Const UltimaModificacion = " Fernando Favre - Se modifico el modelo 45 (Redepa) para que muestre la descripcion del Centro Costo"

'Global Const Version = "3.89"
'Global Const FechaModificacion = "06/12/2011"
'Global Const UltimaModificacion = " Fernando Favre - La lectura del parametro 90 daba error."

'Global Const Version = "3.90"
'Global Const FechaModificacion = "06/12/2011"
'Global Const UltimaModificacion = " Carmen Quintero - Se modificó el modelo 114, para que insert los siguientes campos AFP, COTIZACION, ISAPRE-PLAN, UF, %, MONTO en la tabla rep_recibo"

'Global Const Version = "3.91"
'Global Const FechaModificacion = "16/12/2011"
'Global Const UltimaModificacion = " Fernando Favre - Se modificó el modelo 81, que calcula el Sueldo Base/V. Hora por confrep columna 42"
 
'Global Const Version = "3.92"
'Global Const FechaModificacion = "20/12/2011"
'Global Const UltimaModificacion = " Carmen Quintero - Se modificó el modelo 114, se cambió el valor a mostrar en la columna de COTIZACION"

'Global Const Version = "3.93"
'Global Const FechaModificacion = "22/12/2011"
'Global Const UltimaModificacion = " Carmen Quintero - Se modificó en el modelo 114, la manera de seleccionar la forma de pago"

'Global Const Version = "3.94"
'Global Const FechaModificacion = "27/12/2011"
'Global Const UltimaModificacion = " Carlos Masson - Se agrega el modelo 119 adaptado a Medicus"

'Global Const Version = "3.95"
'Global Const FechaModificacion = "24/01/2012"
'Global Const UltimaModificacion = " Carmen Quintero - Se modificó en el modelo 114, para que no guarde en el campo Puesto el Sector"

'Global Const Version = "3.96"
'Global Const FechaModificacion = "30/01/2012"
'Global Const UltimaModificacion = " Carlos Masson - Se agrega el modelo 121 para OSDOP"

'Global Const Version = "3.97"
'Global Const FechaModificacion = "07/02/2012"
'Global Const UltimaModificacion = " Brzozowski Juan Pablo - Se agrega el modelo 120 para Kraft (Uruguay)"

'Global Const Version = "3.98"
'Global Const FechaModificacion = "15/02/2012"
'Global Const UltimaModificacion = " Dimatz, Rafael - Se agrego el modelo 122 para Mimo Vestiditos"
'                                 Gonzalez Nicolás - Se modificó el valor del campo FREGUESIA (LOCALIDAD), antes guardaba la dirección.

'Global Const Version = "3.99"
'Global Const FechaModificacion = "12/03/2012"
'Global Const UltimaModificacion = " Gonzalez Nicolás - Modelo 111 Portugal - Se volvio atrás cambio anterior y se concateno @ entre direccion y localidad "


'Global Const Version = "4.00"
'Global Const FechaModificacion = "13/03/2012"
'Global Const UltimaModificacion = " Masson, Carlos - Modelo 123 Previsora - Se agregó el modelo 123 basado en el 119 de Medicus"

'Global Const Version = "4.01"
'Global Const FechaModificacion = "14/03/2012"
'Global Const UltimaModificacion = " Dimatz, Rafael - Modelo 122 Mimo Vestiditos - Se dejo mas espacio entre la calle del empleado y el numero"

'Global Const Version = "4.02"
'Global Const FechaModificacion = "20/03/2012"
'Global Const UltimaModificacion = " Gonzalez Nicolás - Modelo 111 Portugal - Se agregaron nuevos AC para las columnas 96,97,98 "

'Global Const Version = "4.03"
'Global Const FechaModificacion = "22/03/2012"
'Global Const UltimaModificacion = " Gonzalez Nicolás - Modelo 111 Portugal - "

'Global Const Version = "4.04"
'Global Const FechaModificacion = "26/03/2012"
'Global Const UltimaModificacion = " Gonzalez Nicolás - Modelo 111 Portugal - Se agregó columna 99"

'Global Const Version = "4.05"
'Global Const FechaModificacion = "03/04/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - Modelo 122 Mimo Vestiditos - Se cambio la estructura Contrato 18 por Leyenda Recibo 45"

'Global Const Version = "4.06"
'Global Const FechaModificacion = "03/04/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - Modelo 122 Mimo Vestiditos - Se saco de la direccion la provincia"


'Global Const Version = "4.07"
'Global Const FechaModificacion = "25/04/2012"
'Global Const UltimaModificacion = " Gonzalez Nicolás - Modelo 111 - DEMO PORTUGAL - Se cambio tenro= 2  a tenro= 3"

'Global Const Version = "4.08"
'Global Const FechaModificacion = "16/05/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - Modelo 69 - Tabacal - Se agrego Obra Social en el campo centrocosto"

'Global Const Version = "4.09"
'Global Const FechaModificacion = "24/05/2012"
'Global Const UltimaModificacion = " Gonzalez Nicolás - Modelo 119 - Medicus - Se Cambio valor de la fecha de ingreso"

'Global Const Version = "4.10"
'Global Const FechaModificacion = "04/06/2012"
'Global Const UltimaModificacion = " Dimatz, Rafael - Modelo 124 - Clinica San Camilo - Se creo el modelo para Clinica San Camilo"

'Global Const Version = "4.11"
'Global Const FechaModificacion = "18/06/2012"
'Global Const UltimaModificacion = " Dimatz, Rafael - Modelo 110 - Coop. Seguros - Se cambio la fecha de alta del empleado para que traiga la mayor y no la de Alta Reconocida"

'Global Const Version = "4.12"
'Global Const FechaModificacion = "21/06/2012"
'Global Const UltimaModificacion = " Dimatz, Rafael - Modelo 110 - Coop. Seguros - Se puso alias en la fecha de alta y se comento MError"

'Global Const Version = "4.13"
'Global Const FechaModificacion = "27/06/2012"
'Global Const UltimaModificacion = " Carlos Masson - Modelo 125 - Se creo el modelo 125 para Paradigma basado en el modelo 124"

'Global Const Version = "4.14"
'Global Const FechaModificacion = "02/07/2012"
'Global Const UltimaModificacion = " Deluchi Ezequiel - Se modifico el modelo 119 (Medicus) para que muestre antiguedad segun lo pedido en el cas 16313"

'Global Const Version = "4.15"
'Global Const FechaModificacion = "02/07/2012"
'Global Const UltimaModificacion = " Deluchi Ezequiel - Correcion en el modelo 119 (Medicus) del calculo de antiguedad - cas 16313"

'Global Const Version = "4.16"
'Global Const FechaModificacion = "13/07/2012"
'Global Const UltimaModificacion = " Deluchi Ezequiel - Correcion en el modelo 119 (Medicus) del calculo de antiguedad - cas 16313 - Se busca la antiguedad en el campo monto "

'Global Const Version = "4.17"
'Global Const FechaModificacion = "18/07/2012"
'Global Const UltimaModificacion = " Gonzalez Nicolás - Nuevo modelo 126 (Perú) - CAS-16441 - H&A - PERU - BOLETA DE PAGO"

'Global Const Version = "4.18"
'Global Const FechaModificacion = "23/07/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - modificaciones al  modelo 126 (Perú) - CAS-16441 - H&A - PERU - BOLETA DE PAGO"
'                                   Gonzalez Nicolás  - Corrección en documento de empresa - DEMO PORTUGAL

'Global Const Version = "4.19"
'Global Const FechaModificacion = "23/07/2012"
'Global Const UltimaModificacion = " Gonzalez Nicolás - modificaciones al  modelo 11 (Portugal) - CAS-13843 - H&A - Nacionalizacion Portugal"

'Global Const Version = "4.20"
'Global Const FechaModificacion = "23/07/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - Se creo nuevo modelo 127 copia del 55 con 4 campos nuevos - CAS-16363- CCU- Custom- Recibo de Haberes"

'Global Const Version = "4.21"
'Global Const FechaModificacion = "15/08/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - 16514 - Se creo modelo nuevo 128 para Apex Honduras"

'Global Const Version = "4.22"
'Global Const FechaModificacion = "15/08/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - 13996 - Se cambio el campo descripcion para que guarde el dato correspondiente para MIMO"

'Global Const Version = "4.23"
'Global Const FechaModificacion = "16/08/2012"
'Global Const UltimaModificacion = " Gonzalez Nicolás - CAS-13843 - H&A - Nacionalizacion Portugal - Se busca en columna 99 el concnro utilizando confval"

'Global Const Version = "4.24"
'Global Const FechaModificacion = "27/08/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - 13719 - Configuracion de Firma de Recibos Coop. Seguros"

'Global Const Version = "4.25"
'Global Const FechaModificacion = "12/09/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - 16690 - Se agrego modelo nuevo 129 para EDESTE"

'Global Const Version = "4.26"
'Global Const FechaModificacion = "17/09/2012"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-16313 - MEDICUS - Recibo de Sueldo Modificacion v2, cambio en el calculo de antiguedad modelo 119 suma todos los historicos de un acumulador "
                                   
'Global Const Version = "4.27"
'Global Const FechaModificacion = "18/09/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - CAS-16919 - NGA - Kraft Chile - Recibo de Sueldo - se crea nuevo modelo 130 para kraft chile, copia del modelo 114 pero con un cambio en la busqueda de las ctas bancarias "

'Global Const Version = "4.28"
'Global Const FechaModificacion = "19/09/2012"
'Global Const UltimaModificacion = "Gonzalez Nicolás - CAS-16865 - Coop. Mutual Seguros - Estimar Modificacion Recibo Sueldo - Modelo 110 - Se busca la Obra Social del empleado"

'Global Const Version = "4.29"
'Global Const FechaModificacion = "28/09/2012"
'Global Const UltimaModificacion = "Masson Carlos - CAS-14945 - OSDOP - Modelo 121 en la fecha de ingreso mostrar la de la ultima Fase"

'Global Const Version = "4.30"
'Global Const FechaModificacion = "02/10/2012"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-16313 - MEDICUS - Recibo de Sueldo Modificacion v2, cambio en el calculo de antiguedad modelo 119 suma todos los historicos de un acumulador (acu_mes) "

'Global Const Version = "4.31"
'Global Const FechaModificacion = "03/10/2012"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-16313 - MEDICUS - Recibo de Sueldo Modificacion v2, validacion sobre nulo de lo hecho en la version 4.30  "

'Global Const Version = "4.32"
'Global Const FechaModificacion = "10/10/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 17154 - AGD - Se cambio Ternro por Long en Sub bus_Antiguedad "

'Global Const Version = "4.33"
'Global Const FechaModificacion = "16/10/2012"
'Global Const UltimaModificacion = " Carlos Masson - CAS 17216 - Express Beer - Se agregó el modelo 131 para Express Beer "

'Global Const Version = "4.34"
'Global Const FechaModificacion = "18/10/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 17154 - AGD - Se cambio la definicion del ternro por long en Public Sub diastrab ByVal Ternro As Long "

'Global Const Version = "4.35"
'Global Const FechaModificacion = "19/10/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - se agrego modelo 132 para TATA"

'Global Const Version = "4.36"
'Global Const FechaModificacion = "02/11/2012"
'Global Const UltimaModificacion = " Carlos Masson - se agrego modelo 133 para Sedamil"

'Global Const Version = "4.37"
'Global Const FechaModificacion = "06/11/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - se agrego modelo 134 para HB VISION OUTSOURCES"

'Global Const Version = "4.38"
'Global Const FechaModificacion = "06/11/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - se corrigio modelo 132 cuando busca el documento bse de la empresa - CAS-16722 - Telefax - Tata - Recibo de Sueldo"

'Global Const Version = "4.39"
'Global Const FechaModificacion = "27/11/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - Se creo modelo 135 para Paraguay - CAS 17545 - H&A - Nacionalizacion Py - Recibo de Sueldo"
'Global Const UltimaModificacion = " guarda todos los conceptos y acumuladores imprimibles"

'Global Const Version = "4.40"
'Global Const FechaModificacion = "05/12/2012"
'Global Const UltimaModificacion = " Carlos Masson - Se creo modelo 136 para Conexia Colombia - CAS - 17409 – Conexia SAS – Nuevo reporte de Recibo de Sueldo"

'Global Const Version = "4.41"
'Global Const FechaModificacion = "20/12/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - se modifico la longitud del campo categoria en el modelo 132 TATA - "
'Global Const UltimaModificacion = " CAS-16722 - Telefax - Tata - Recibo de Sueldo"

'Global Const Version = "4.42"
'Global Const FechaModificacion = "20/12/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - se modifico el modelo 135 para que muestre los acumuladores "
'Global Const UltimaModificacion = " configurados en el confrep y todos los conceptos imprimibles "
'Global Const UltimaModificacion = " CAS-17545 - H&A - Nacionalizacion Py - Recibo de Sueldo"

'Global Const Version = "4.43"
'Global Const FechaModificacion = "21/12/2012"
'Global Const UltimaModificacion = " Dimatz Rafael - Se creo modelo nuevo 138 para Diaz"

'Global Const Version = "4.44"
'Global Const FechaModificacion = "28/12/2012"
'Global Const UltimaModificacion = " Carlos Masson - Se modificó el modelo 125 agregando el modelo de liquidación para poder mostrarlo en el recibo"

'Global Const Version = "4.45"
'Global Const FechaModificacion = "02/01/2013"
'Global Const UltimaModificacion = " Gonzalez Nicolás - Se creo modelo 137 para NGA"

'Global Const Version = "4.46"
'Global Const FechaModificacion = "07/01/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - Se corrige error en variable ImpIRPF, se la declara como double y se hace replace para el insert."
'                                  'CAS-16722 - Telefax - Tata - Recibo de Sueldo

'Global Const Version = "4.47"
'Global Const FechaModificacion = "07/01/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 16839 - Sedamil Modelo 133 - Se agrego el campo Zona, es tomado del confrep, se agrego FormaPago, Banco y Cuenta. Configurar columna 136 con un valor en confval "

'Global Const Version = "4.48"
'Global Const FechaModificacion = "09/01/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - CAS-17545 - H&A - Nacionalizacion Py - Recibo de Sueldo - se descomento la variable columnasConfrep, por lo tanto se muestran los acumuladores que se configuren en las columnas del confrep"

'Global Const Version = "4.49"
'Global Const FechaModificacion = "10/01/2013"
'Global Const UltimaModificacion = " SDimatz Rafael - CAS 16839 - Sedamil Modelo 133 - Se corrigio para que en Forma de Pago muestre la Cuenta Bancaria y si tiene tipo Forma de Pago 11 muestre CC"

'Global Const Version = "4.50"
'Global Const FechaModificacion = "23/01/2013"
'Global Const UltimaModificacion = " Gonzalez Nicolás -  CAS-17836- NORTHGATE- Recibo de Sueldo Configurable + PDF - Se modifico modelo 137. se creo variable nueva para asignar un string e insertar en auxchar4 (contenido de la col44)"

'Global Const Version = "4.51"
'Global Const FechaModificacion = "23/01/2013"
'Global Const UltimaModificacion = " Sebastian Stremel -  Se corrigio cuando levanto valores del confrep, si es concepto lo tomo de confval2, sino es AC de confval - CAS-16722 - Recibo de sueldo - Telefax Tata-(CAS-15298) "

'Global Const Version = "4.52"
'Global Const FechaModificacion = "28/01/2013"
'Global Const UltimaModificacion = " Dimatz Rafael -  Se cambio en el Recibo de Diaz modelo 138 la estructura de la modalidad por la 53 - CAS-17005 "

'Global Const Version = "4.53"
'Global Const FechaModificacion = "05/02/2013"
'Global Const UltimaModificacion = " Fernando Favre -  Se creo el modelo 139 a partir del modelo 56 - CAS-18362"

'Global Const Version = "4.54"
'Global Const FechaModificacion = "13/02/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - Se busca el nro patronal de la empresa, el mismo se obtiene del confrep de la columna 200 configurada como tipo DOC"
'                                 CAS-17545 - H&A - Nacionalizacion Py - Recibo de Sueldo - Modelo 135

'Global Const Version = "4.55"
'Global Const FechaModificacion = "14/02/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - el nro patronal de la empresa se obtiene ahora de his_estructura, el tidnro se obtiene del confrep de la columna 200 configurada como tipo DOC"
'                                 CAS-17545 - H&A - Nacionalizacion Py - Recibo de Sueldo - Modelo 135

'Global Const Version = "4.56"
'Global Const FechaModificacion = "26/03/2013"
'Global Const UltimaModificacion = " Fernando Favre - La fecha de alta (modelo 139) se calcula de la fase marcada con Fecha De Alta Reconocida."
'                                 CAS - 18765 - CDA - Bug en la fecha de ingreso y antiguedad del recibo - modelo 139

'Global Const Version = "4.57"
'Global Const FechaModificacion = "03/04/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 18330 - Se creo modelo 140 Recibo Ecuador"

'Global Const Version = "4.58"
'Global Const FechaModificacion = "04/04/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS 18893 - Horwath Litoral - AMR - Se creo modelo 141 Recibo AMR"

'Global Const Version = "4.59"
'Global Const FechaModificacion = "08/04/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - CAS 16441 - CAS-16441 - H&A - PERU - Modificación Boleta de Pago"
                                   'CAS-16441 - H&A - PERU - Boleta de Pago - se realizaron modificaciones varias en modelo 126

'Global Const Version = "4.60"
'Global Const FechaModificacion = "09/04/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - CAS-16441 -  H&A -  PERU - DISTRIBUCION DE UTILIDADES - Se creo el modelo 142"
                                   
'Global Const Version = "4.61"
'Global Const FechaModificacion = "29/04/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19208 - Se creo el modelo 143 para Chacomer"

'Global Const Version = "4.62"
'Global Const FechaModificacion = "07/05/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19044 - Se creo el modelo 144 para Venezuela"

'Global Const Version = "4.63"
'Global Const FechaModificacion = "15/05/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19044 - Se modifico para poder configurar en el confrep un Concepto o Acumulador para el Sueldo"

'Global Const Version = "4.64"
'Global Const FechaModificacion = "17/05/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19208 - Se modifico direccion del empleado"

'Global Const Version = "4.65"
'Global Const FechaModificacion = "21/05/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - CAS-16441 - H&A - PERU - BOLETA DE PAGO [Entrega 2] - se buscan los datos del logo de la empresa - se busca las fechas de las vacaciones para el periodo de la boleta de pago modelo 126"

'Global Const Version = "4.66"
'Global Const FechaModificacion = "23/05/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - Se crea un nuevo modelo y se genera el nuevo recibo - CAS-19634 - PayrollPy - Axion - Recibo de Sueldo -"

'Global Const Version = "4.67"
'Global Const FechaModificacion = "27/05/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19208 - Se modifico modelo 143 - Chacomer - Se modifico Grenecia y Sector. Se agrego guaraníes en la leyenda. Se agrego espacio en la firma. Se corrigio Codigo de la Sucursal"

'Global Const Version = "4.68"
'Global Const FechaModificacion = "04/06/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19627 - Se corrigio del Modelo 92 Multivoice la Forma de Pago"

'Global Const Version = "4.69"
'Global Const FechaModificacion = "05/06/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19398 - Se creo modelo nuevo 146 - COLOMBIA NVS - CAS 19398"

'Global Const Version = "4.70"
'Global Const FechaModificacion = "05/06/2013"
'Global Const UltimaModificacion = " Ana Annese - CAS-11049 - Recibo De Haberes - Cargo"

'Global Const Version = "4.71"
'Global Const FechaModificacion = "07/06/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19208 - Se Modifico la Localidad y Direccion para que traiga la de la Sucursal Modelo 143 Chacomer"

'Global Const Version = "4.72"
'Global Const FechaModificacion = "13/06/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19044 - Se modificaron Query. Ademas se corrigio para que pueda traer el valor de la segunda columna del confrep para el Sueldo"

'Global Const Version = "4.73"
'Global Const FechaModificacion = "17/06/2013"
'Global Const UltimaModificacion = " Sebastian Stremel  - CAS 16441 - Se busca el representante legal de la empresa, y se hace configurable por confrep el ruc a mostrar."

'Global Const Version = "4.74"
'Global Const FechaModificacion = "25/06/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19976 - Se creo un modelo nuevo para Novartis Modelo 147"

'Global Const Version = "4.75"
'Global Const FechaModificacion = "03/07/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - CAS 16441 DISTRIBUCION DE UTILIDADES - se corrigio errores cuando levantaba el confrep, es continuacion de la version 4.73 "

'Global Const Version = "4.76"
'Global Const FechaModificacion = "04/07/2013"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-19829 - HORWATH LITORAL - GEMPLAST - Recibo Blue - Nuevo modelo 148"

'Global Const Version = "4.77"
'Global Const FechaModificacion = "08/07/2013"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-19940 - EXPRESSBEER - CUSTOM RECIBO DE SUELDO - Modificacion del modelo 131 se agrego sueldo basico configurable por confrep (COM,ACL,ACM)"

'Global Const Version = "4.78"
'Global Const FechaModificacion = "10/07/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-18991 - Raffo - Adecuaciones LIQ - Recibo de Sueldo - Nuevo modelo 149 basado en en modelo 125"

'Global Const Version = "4.79"
'Global Const FechaModificacion = "11/07/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 20087 - Horwath Literal Gemplast - Crear Nuevo Recibo de Sueldo Modelo 150 "

'Global Const Version = "4.80"
'Global Const FechaModificacion = "15/07/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 17053 - Brasil - Crear Nuevo Recibo de Sueldo Modelo 151 "

'Global Const Version = "4.81"
'Global Const FechaModificacion = "16/07/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 19976 - Novartis - Se modifico para que Salga la Leyenda del Parametro "

'Global Const Version = "4.82"
'Global Const FechaModificacion = "22/07/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-18991 - Raffo - Adecuaciones LIQ - Recibo de Sueldo - Correcciones en modelo 149"

'Global Const Version = "4.83"
'Global Const FechaModificacion = "23/07/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-20224 - CDA - RECIBO DE HABERES CHILE Y PERU - Nuevo modelo 152 Basado en el modelo 130"

'Global Const Version = "4.84"
'Global Const FechaModificacion = "30/07/2012"
'Global Const UltimaModificacion = " Sebastian Stremel - Modelo 126 se busco el sueldo que puede salir de uno u otro concepto - CAS-16441 - H&A - PERU - BOLETA DE PAGO"

'Global Const Version = "4.85"
'Global Const FechaModificacion = "02/08/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-18991 - Raffo - Adecuaciones LIQ - Recibo de Sueldo - Se modifica modelo 149, Se agrega centro de costo"

'Global Const Version = "4.86"
'Global Const FechaModificacion = "06/08/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-20224 - CDA - RECIBO DE HABERES CHILE Y PERU - Se modifica modelo 152, el tipodoc para el cuit es 201 y se agregó el modelo 153 para Perú basado en el 126"

'Global Const Version = "4.87"
'Global Const FechaModificacion = "07/08/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-18330 - Ecuador - Se corrigio el forma de pago y se agrego DatosCBO para ue muestre la Descripcion Abreviada de la Estructura configurada en el confrep"

'Global Const Version = "4.88"
'Global Const FechaModificacion = "15/08/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-20224 - CDA - RECIBO DE HABERES CHILE Y PERU - Se modifica modelo 153 para que el sueldo neto se tome de la columna 304 y de mutre la descripcion de la estructura Area en vez del código externo"

'Global Const Version = "4.89"
'Global Const FechaModificacion = "15/08/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-18991 - Raffo - Adecuaciones LIQ - Recibo de Sueldo - Se modifica modelo 149, Se agrega el piso en la dirección de la empresa"

'Global Const Version = "4.90"
'Global Const FechaModificacion = "22/08/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 14302 - Plusmar - Se creo Modelo Nuevo Nro 154 "

'Global Const Version = "4.91"
'Global Const FechaModificacion = "29/08/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-20791 - CDA - Error en Antigüedad de Recibo Haberes Argentina - Se modifica la fecha que se toma como referencia para calcular la antiguedad (modelo 139)"

'Global Const Version = "4.92"
'Global Const FechaModificacion = "30/08/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 14302 - Plusmar - Se cambio en la query rsconsult2 por rsconsult2 "

'Global Const Version = "4.93"
'Global Const FechaModificacion = "05/09/2013"
'Global Const UltimaModificacion = " Sebastian Stremel -  Se busca el valor del carnet AFP del tipo de documento configurado en la columna 243 del confrep"
                                    'CAS-16441 -  NOVARTIS - BOLETA DE PAGO [Entrega 3] "

'Global Const Version = "4.94"
'Global Const FechaModificacion = "06/09/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 14302 - Plusmar - Se agrego en la query Forma de Pago el Tipo 10 Caja de Ahorro "

'Global Const Version = "4.95"
'Global Const FechaModificacion = "06/09/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 21214 - Sedamil - Se corrigio la query de Antiguedad "

'Global Const Version = "4.96"
'Global Const FechaModificacion = "12/09/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 20087 - Horwath Litoral - Gemplast - Se agregaron los datos del encabezado, Nombre Empresa, Direccion, Cuit, Logo"

'Global Const Version = "4.97"
'Global Const FechaModificacion = "24/09/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-20702 - CDA - Recibo de haberes Mexico - Se agrega el modelo 155"

'Global Const Version = "4.98"
'Global Const FechaModificacion = "26/09/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-21214 - Sedamil - Se saca del modelo 133 en Antiguedad la parte final que dice desde cuando esta Activo"

'Global Const Version = "4.99"
'Global Const FechaModificacion = "26/09/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-21519 - AMR - Se agrego la Forma de Pago"

'Global Const Version = "5.00"
'Global Const FechaModificacion = "04/10/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-21142 - CDA - Recibo de Haberes España - Se agrega el modelo 156"

'Global Const Version = "5.01"
'Global Const FechaModificacion = "04/10/2013"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-21426 - SGS - Modificaciones Recibo de Haberes modelo 71 - Nuevo Modelo 157"

'Global Const Version = "5.02"
'Global Const FechaModificacion = "10/10/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-20702 - CDA - Recibo de haberes Mexico - Se modifica el modelo 155 haciendo configurables los documentos RFC y Reg. Pat."

'Global Const Version = "5.03"
'Global Const FechaModificacion = "22/10/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS 21442 - Recibo BDO - Se creo Recibo para BDO GE Modelo 158"

'Global Const Version = "5.04"
'Global Const FechaModificacion = "30/10/2013"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-21897 - BDO - Custom de Recibo de Haberes - Se creo Modelo 159"

'Global Const Version = "5.05"
'Global Const FechaModificacion = "31/10/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-21817 - Colombia Zoetis - Se creo Modelo 160"

'Global Const Version = "5.06"
'Global Const FechaModificacion = "05/11/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-22135 - Chile Zoetis - Se creo Modelo 161"

'Global Const Version = "5.07"
'Global Const FechaModificacion = "05/11/2013"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-21426 - SGS - Modificaciones Recibo de Haberes modelo 71, se agrego contrato al modelo 157"

'Global Const Version = "5.08"
'Global Const FechaModificacion = "06/11/2013"
'Global Const UltimaModificacion = " Sebastian Stremel - Se inicializan variables del confrep en la columna 142 para que no se rompa el proceso - CAS-15298 - CAS-16441 -  H&A -  PERU - DISTRIBUCION DE UTILIDADES [Entrega 6] (CAS-15298)"

'Global Const Version = "5.09"
'Global Const FechaModificacion = "07/11/2013"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-21897 - BDO - Custom de Recibo de Haberes, se cambio fecha de alta y forma de pago"

'Global Const Version = "5.10"
'Global Const FechaModificacion = "11/11/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-22135 - Chile Zoetis - Se modifico la query para la direccion de la empresa, se saco la coma"

'Global Const Version = "5.11"
'Global Const FechaModificacion = "14/11/2013"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-21897 - BDO - Custom de Recibo de Haberes, se corrigio forma de pago y se cambio forma de calcular fecha de alta"

'Global Const Version = "5.12"
'Global Const FechaModificacion = "15/11/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-22135 - Chile Zoetis - Se modifico la query de la Forma de Pago y el Banco"

'Global Const Version = "5.13"
'Global Const FechaModificacion = "20/11/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-22200 - Deloitte - Se modifico la query de Fecha de Alta"

'Global Const Version = "5.14"
'Global Const FechaModificacion = "28/11/2013"
'Global Const UltimaModificacion = " Deluchi Ezequiel - CAS-21897 - BDO - Custom de Recibo de Haberes, cambio en el calculo de fecha de ingreso y egreso "

'Global Const Version = "5.15"
'Global Const FechaModificacion = "29/11/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-21817 - Recibo Colombia Zoetis, se reemplazo en el sueldo la coma por un punto. Se agrego Cuenta en la tabla rep_recibo "

'Global Const Version = "5.16"
'Global Const FechaModificacion = "05/12/2013"
'Global Const UltimaModificacion = " Mauricio Zwenger - CAS-22280  - Modelo 65 - Se cambio la forma de calcular la fecha de alta, ahora la saca de la fase "

'Global Const Version = "5.17"
'Global Const FechaModificacion = "16/12/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-22821  - Modelo 1 - Se modifico la consulta de Lugar de Pago y Forma de Pago, no traia bien los resultados "

'Global Const Version = "5.18"
'Global Const FechaModificacion = "17/12/2013"
'Global Const UltimaModificacion = " Dimatz Rafael - CAS-21426  - Modelo 157 - Se modifico la Fecha de alta, para que sea la Fecha de la Fase Activa y se modifico la Estructura de Gerencia por la 35 "

'Global Const Version = "5.19"
'Global Const FechaModificacion = "20/12/2013"
'Global Const UltimaModificacion = " Carlos Masson - CAS-20702 - CDA - Recibo de haberes Mexico - Se modifica el modelo 155: se pasa el Reg. Pat. al campo auxchar6"

'Global Const Version = "5.20"
'Global Const FechaModificacion = "20/12/2013"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 22135 - Zoetis Chile - Modelo 161 Se agrego query de confrep Columna 335 que resuelve el valor que muestra Plan Salud"

'Global Const Version = "5.21"
'Global Const FechaModificacion = "02/01/2014"
'Global Const UltimaModificacion = " Carlos Masson - CAS-21142 - CDA - Recibo de Haberes España - modelo 156: se modifica la consulta de domicilio"

'Global Const Version = "5.22"
'Global Const FechaModificacion = "15/01/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 23113 - Zarcam - Modelo 72 - Se corrigio Fecha de Ingreso para que muestre la Fecha de la ultima Fase activa"

'Global Const Version = "5.23"
'Global Const FechaModificacion = "15/01/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 23377 - HIRSCH - Modelo 139 - Se cambio el SECTOR por el AREA"

'Global Const Version = "5.24"
'Global Const FechaModificacion = "21/01/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 22772 - IBT - Modelo 126 - Se agrego Firma de Empresa"

'Global Const Version = "5.25"
'Global Const FechaModificacion = "22/01/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 23473 - SGS - Modelo 157 - Se agrego la Fecha Reconocida de la Fase"

'Global Const Version = "5.26"
'Global Const FechaModificacion = "31/01/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 23377 - HIRSCH - Modelo 163 - Se cambio el Nro de Estructura 120 por el Nro 108"

'Global Const Version = "5.27"
'Global Const FechaModificacion = "12/02/2014"
'Global Const UltimaModificacion = "LED - CAS-23403 - TETRAPAK - CUSTOM RECIBOS - Modelo 49 - Se obtiene el dni del empleado para la exportacion a pdf"

'Global Const Version = "5.28"
'Global Const FechaModificacion = "14/02/2014"
'Global Const UltimaModificacion = "LM - CAS-23269 - SEDAMIL - CUSTOM FECHA EN RECIBO DE HABERES [Entrega 3] se cambio la fecja del proceso por la fecha del pedido de pago"

'Global Const Version = "5.29"
'Global Const FechaModificacion = "24/02/2014"
'Global Const UltimaModificacion = "Gonzalez Nicolás - CAS-23702 - Sykes El salvador - Boleta de pago (quincenal)"
                                  ' Se creó módulo repRecibo3 para los nuevos modelos a partir del 500

'Global Const Version = "5.30"
'Global Const FechaModificacion = "25/02/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS-23377 - HIRSCH - Recibo se configura el Area por confrep en la columna 336"

'Global Const Version = "5.31"
'Global Const FechaModificacion = "27/02/2014"
'Global Const UltimaModificacion = " CAS-23704 - SYKES El salvador -  Comprobante de pago (mensual)"
                                  ' Nuevo modelo 501 - Comprobante de Pago

'Global Const Version = "5.32"
'Global Const FechaModificacion = "27/02/2014"
'Global Const UltimaModificacion = "Gonzalez Nicolás - CAS-23702 - Sykes El salvador - Boleta de pago (quincenal)"
                                  ' Modelo 500: Se agregaron Mid() en los insert para los campos de tipo varchar
'Global Const Version = "5.33"
'Global Const FechaModificacion = "27/02/2014"
'Global Const UltimaModificacion = "Gonzalez Nicolás - CAS-23701 - SYKES El Salvador -  Recibo de liquidación (final)"
                                  'Nuevo Modelo 502 - Recibo de Liquidación (Final)


'Global Const Version = "5.34"
'Global Const FechaModificacion = "06/03/2014"
'Global Const UltimaModificacion = "Mauricio Zwenger - CAS-23790 - PLA SA -  Recibo de haberes"
                                  'Nuevo Modelo 164 - Recibo de haberes
                                  

'Global Const Version = "5.35"
'Global Const FechaModificacion = "12/03/2014"
'Global Const UltimaModificacion = "Mauricio Zwenger - CAS-23857 - UTDT - Custom Recibo de sueldo"
                                  'Nuevo Modelo 165 - Recibo de haberes
                                  
'Global Const Version = "5.36"
'Global Const FechaModificacion = "14/03/2014"
'Global Const UltimaModificacion = "Mauricio Zwenger - CAS-23790 - PLA SA -  Recibo de haberes"
                                  'se configura por confrep (columnas 341 y 342 de reporte 60) las estructuras correspondientes a primera y segunda quincena

'Global Const Version = "5.37"
'Global Const FechaModificacion = "17/03/2014"
'Global Const UltimaModificacion = " CAS-23704 - SYKES El salvador -  Comprobante de pago (mensual)"
                                  ' Modelo 501 - Se cambió inner join por left en query que buscaba empreporta
                                  
'Global Const Version = "5.38"
'Global Const FechaModificacion = "21/03/2014"
'Global Const UltimaModificacion = " Carlos Masson - CAS-24323 - MEGATLON - Recibo de Haberes"
                                  ' Se creó el modelo 166 para Megatlon

'Global Const Version = "5.39"
'Global Const FechaModificacion = "25/03/2014"
'Global Const UltimaModificacion = "Mauricio Zwenger - CAS-23790 - PLA SA -  Recibo de haberes"
                                    'se corrigio validacion de estructuras quincenales

'Global Const Version = "5.40"
'Global Const FechaModificacion = "28/03/2014"
'Global Const UltimaModificacion = "Sebastian Stremel - CAS-24504 - ZOETIS CHILE - CUSTOM RECIBO DE SUELDO"
                                    'Se configura el tipo de documento para el rut de la empresa (modelo 161)
                                    'Se guardan todos los conceptos tanto imprimibles como no imprimibles.

'Global Const Version = "5.41"
'Global Const FechaModificacion = "31/03/2014"
'Global Const UltimaModificacion = "Mauricio Zwenger - CAS-23857 - UTDT - Custom Recibo de sueldo"
                                  ' se corrigieron errores en obtencion de estructuras configurables

'Global Const Version = "5.42"
'Global Const FechaModificacion = "03/04/2014"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-24119 - HORWATH LITORAL - AMR - Modificacion Recibo de Sueldo"
                                  ' Se modificó el modelo 141 para que el valor de la ganancia neta
                                  'se obtenga igual que en el reporte de ganancias

'Global Const Version = "5.43"
'Global Const FechaModificacion = "07/04/2014"
'Global Const UltimaModificacion = "Miriam Ruiz - CAS-24538 - CCU - ERROR EN ANTIGUEDAD EN RECIBO"
                                  ' Se corrigió el cálculo de la antiguedad en el modelo 127
                                  
'Global Const Version = "5.44"
'Global Const FechaModificacion = "07/04/2014"
'Global Const UltimaModificacion = "Borrelli Facundo - CAS-23269 - SEDAMIL - CUSTOM FECHA EN RECIBO DE HABERES [Entrega 5]"
                                  'Se corrige para que se muestre la fecha del pedido de pago en el recibo de haberes
                                  'Modelo 133
                              
'Global Const Version = "5.45"
'Global Const FechaModificacion = "11/04/2014"
'Global Const UltimaModificacion = "Sebastian Stremel  - CAS-24504 - ZOETIS CHILE - CUSTOM RECIBO DE SUELDO [Entrega 3]"
                                  'Se muestran los primeros 100 caracteres de la cta bancaria y la descripcion del banco
                                  
'Global Const Version = "5.46"
'Global Const FechaModificacion = "28/04/2014"
'Global Const UltimaModificacion = "Miriam Ruiz  - CAS-24235 - ZOETIS CHILE - BUG EN RECIBO DE SUELDO (CAS-15298)"
                                  'Se agregó control cuando no está configurada la columna 335 del modelo 161
                              
'Global Const Version = "5.47"
'Global Const FechaModificacion = "30/04/2014"
'Global Const UltimaModificacion = "Sebastian Stremel - CAS-24504 - ZOETIS CHILE - CUSTOM RECIBO DE SUELDO [Entrega 3]"
                                  'Modelo 161, si el concepto es el 999998 le cambio el tipo para que salga en los descuento.

'Global Const Version = "5.48"
'Global Const FechaModificacion = "13/05/2014"
'Global Const UltimaModificacion = "Facundo Eggle - CAS-25276 - 5CA - Modificaciones en recibo de haberes - Modelo 167"

'Global Const Version = "5.49"
'Global Const FechaModificacion = "20/05/2014"
'Global Const UltimaModificacion = "LED - CAS-24624 - V.O - PAYROLL - MODIFICACION EN RECIBO - Modelo 115, se obtiene el documento asociado a la empresa configurado en el confrep"

'Global Const Version = "5.50"
'Global Const FechaModificacion = "21/05/2014"
'Global Const UltimaModificacion = "Facundo Eggle - CAS-25506 - PUIG - Reporte de Recibo de Haberes - modificación - Modelo 168"

'Global Const Version = "5.51"
'Global Const FechaModificacion = "21/05/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS-24978 - Bolivia - Se crearon 3 Modelos para Bolivia 503 - 504 - 505"

'Global Const Version = "5.52"
'Global Const FechaModificacion = "23/05/2014"
'Global Const UltimaModificacion = "Facundo Eggle - CAS-25506 - PUIG - Reporte de Recibo de Haberes - modificación - Modelo 168"
                                  'Se eliminan las búsquedas de "neto en pesos", "sueldo basico" y "neto en bonos"

'Global Const Version = "5.53"
'Global Const FechaModificacion = "05/06/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS-22772 - Peru - Se agregaron campos para mostrar en el Recibo PDF - Modelo 126"

'Global Const Version = "5.54"
'Global Const FechaModificacion = "06/06/2014"
'Global Const UltimaModificacion = "Facundo Eggle - CAS-25848 - Cambio en Recibo de Haberes - modelo 157"

'Global Const Version = "5.55"
'Global Const FechaModificacion = "10/06/2014"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-25915 - SEDAMIL - ERROR EN FECHA DE PAGO DE RECIBO"
                                  'Se corrige para que se muestre la fecha del pedido de pago en el recibo de haberes del Modelo 133

'Global Const Version = "5.56"
'Global Const FechaModificacion = "11/06/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS-22772 - IBT - Se corrigio para que guarde el Doc113 como alfanumerico - Modelo 126"
                                  
'Global Const Version = "5.57"
'Global Const FechaModificacion = "17/06/2014"
'Global Const UltimaModificacion = "Borrelli Facundo - CAS-25673 - VISION - ERROR EN ANTIGUEDAD DE RECIBO"
                                  ' Se corrigió el cálculo de la antiguedad en el modelo 56
                                  
'Global Const Version = "5.58"
'Global Const FechaModificacion = "26/06/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS-24978 - Recibo de Bolivia - Se agrego el Modelo del Proceso y la Fecha Ultimo Anticipo"

'Global Const Version = "5.59"
'Global Const FechaModificacion = "04/07/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS-24978 - Recibo de Bolivia - Se agrego para que salgan los conceptos del confrep en el encabezado"

'Global Const Version = "5.60"
'Global Const FechaModificacion = "11/07/2014"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-26315 - Punto Farma - Modificación de Recibo de Haberes"

'Global Const Version = "5.61"
'Global Const FechaModificacion = "15/07/2014"
'Global Const UltimaModificacion = "CCarlos Masson - CAS-25918 - Clarín - Modificación recibo de haberes"

'Global Const Version = "5.62"
'Global Const FechaModificacion = "25/07/2014"
'Global Const UltimaModificacion = "LED - CAS-25416 - NORTHGATE ARINSO - CAMBIO DE NUMERO EN LEYENDA - Se agrego tipo de codigo configurable col 357 reporte, para informar en el pdf numero de resolucion"

'Global Const Version = "5.63"
'Global Const FechaModificacion = "28/07/2014"
'Global Const UltimaModificacion = "LED - CAS-26033 - ANDREANI - Modificaciones Recibos de Sueldos - Se agrego tipo de documento configurable col 358 reporte, para informar en el recibo numero de resolucion"

'Global Const Version = "5.64"
'Global Const FechaModificacion = "20/08/2014"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-23704 - SYKES El salvador -  Comprobante de pago (mensual) [Entrega 3] - Se consideró el mes de liquidacion en la consulta donde se insertan los registro en la tabla rep_recibo_det"


'Global Const Version = "5.65"
'Global Const FechaModificacion = "03/09/2014"
'Global Const UltimaModificacion = "Se agregaron los modelos 171, 172, 173 y 174" ' Mauricio Zwenger - CAS-26571 - VSO - Interpack - Recibos de Sueldo

'Global Const Version = "5.66"
'Global Const FechaModificacion = "05/09/2014"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-23704 - SYKES El salvador -  Comprobante de pago (mensual) [Entrega 5] - Mejoras varias al modelo 501"

'Global Const Version = "5.67"
'Global Const FechaModificacion = "05/09/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 21913 - Modelo 65 Apex - Se agrego en la tabla de recibo que inserte el Nro de Estructura de la Empresa "

'Global Const Version = "5.68"
'Global Const FechaModificacion = "08/09/2014"
'Global Const UltimaModificacion = "Maurcio Zwenger - CAS-26571 - Modificaciones varias, modelos 171, 172, 173 y 174.  "

'Global Const Version = "5.69"
'Global Const FechaModificacion = "11/09/2014"
'Global Const UltimaModificacion = "Carlos Masson - CAS-26727 - Ingenio la Esperanza - Modificación Recibo de Haberes. Nuevo modelo 175, basado en el 55 de CCU  "

'Global Const Version = "5.70"
'Global Const FechaModificacion = "11/09/2014"
'Global Const UltimaModificacion = "Maurcio Zwenger - CAS-26571 - Modificaciones varias, modelos 171, 172, 173 y 174.  "

'Global Const Version = "5.71"
'Global Const FechaModificacion = "12/09/2014"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-23704 - SYKES El salvador - Comprobante de pago (mensual) (CAS-15298) - Se modificó el modelo 501 para que muestre los datos del empleado, aun cuando no tenga Reporta A asignado. "

'Global Const Version = "5.72"
'Global Const FechaModificacion = "15/09/2014"
'Global Const UltimaModificacion = "LED - CAS-26500 - BDO - CUSTOM RECIBO KRAFT URUGUAY [Entrega 2] - Cambio en el modelo 120, se informa la primer cuenta activa del empleado como forma de pago "

'Global Const Version = "5.73"
'Global Const FechaModificacion = "23/09/2014"
'Global Const UltimaModificacion = "Sebastian Stremel - CAS-23881 - Telefax - Santander URU- Recibo de Sueldos - Se creo el modelo 176 para Santander UY"

'Global Const Version = "5.74"
'Global Const FechaModificacion = "24/09/2014"
'Global Const UltimaModificacion = "Miriam Ruiz  - CAS-17053 - Nac Brasil - Bug recibo de pagamento - Se comentaron los close de los recordset en el modelo 151"

'Global Const Version = "5.75"
'Global Const FechaModificacion = "24/09/2014"
'Global Const UltimaModificacion = "Sebastian Stremel  - CAS-25915 - SEDAMIL - CUSTOM LUGAR DE PAGO EN RECIBO DE HABERES - Se busca el Lugar de Pago, tipo de estructura 20"

'Global Const Version = "5.76"
'Global Const FechaModificacion = "02/10/2014"
'Global Const UltimaModificacion = "Carlos Masson - CAS-20702 - CDA - Recibo Mexico [Entrega 7] (CAS-15298) - Cuandofaltan configurar las columnas del confrep no sigue"

'Global Const Version = "5.77"
'Global Const FechaModificacion = "02/10/2014"
'Global Const UltimaModificacion = "Maurcio Zwenger - CAS-26571 - Modificaciones varias, modelos 171, 172, 173 y 174. se cambio direccion de empresa por direccion de sucursal "

'Global Const Version = "5.78"
'Global Const FechaModificacion = "15/10/2014"
'Global Const UltimaModificacion = "Sebastian Stremel - CAS-23881 - Telefax - Santander URU- Recibo de Sueldos [Entrega 2]- se corrigio la localidad y el valor neto "

'Global Const Version = "5.79"
'Global Const FechaModificacion = "16/10/2014"
'Global Const UltimaModificacion = "LED - CAS-26500 - BDO - CUSTOM RECIBO KRAFT URUGUAY [Entrega 3]- Correcion en el modelo 120, valor seguro de vida se paso a formato sql para la insercion del valor"

'Global Const Version = "5.80"
'Global Const FechaModificacion = "16/10/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS-17053 - Brasil - Se Corrigio para que traiga el Tipo de Documento de la Empresa configurable por Confrep"

'Global Const Version = "5.81"
'Global Const FechaModificacion = "28/10/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 24978 - Se corrigio para que Calcule bien los meses Calculo Promedio - Busca solo los Mensuales"

'Global Const Version = "5.82"
'Global Const FechaModificacion = "29/10/2014"
'Global Const UltimaModificacion = "Fernando Favre - CAS-27387 - FESTO - Adecuación de Recibo - Se creo el modelo 177 a partir del modelo 117"

'Global Const Version = "5.83"
'Global Const FechaModificacion = "03/11/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 24978 - Se corrigio el Modelo 505 para calcular los meses de los 3 meses anteriores al liquidado"

'Global Const Version = "5.84"
'Global Const FechaModificacion = "04/11/2014"
'Global Const UltimaModificacion = "Mauricio Zwenger - CAS-27164 - NGA - Se creo el modelo 178 para San Miguel (NGA) basado el el modelo 55 (CCU)"

'Global Const Version = "5.85"
'Global Const FechaModificacion = "06/11/2014"
'Global Const UltimaModificacion = "Miriam Ruiz- CAS-26972 - H&A - Bug en modelo estándar del recibo de sueldo - se corrigió el modelo 1 para que salgan correctamente el banco y la cuenta bancaria y la fecha de ingreso"

'Global Const Version = "5.86"
'Global Const FechaModificacion = "06/11/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 24978 - Se corrigio Modelo 504 el calculo promedio de 3 meses, la Fecha de Ingreso y la Fecha Ultimo Anticipo Indemnizacion"

'Global Const Version = "5.87"
'Global Const FechaModificacion = "06/11/2014"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 24978 - Se modifico la query Centro de Costo, para que traiga el Departamento"

'Global Const Version = "5.88"
'Global Const FechaModificacion = "10/11/2014"
'Global Const UltimaModificacion = "Fernando Favre - CAS-27387 - FESTO - Adecuación de Recibo - rechazo. Modelo 177, se agrego buscar el tipo de estructura 20 - Lugar de Pago"

'Global Const Version = "5.89"
'Global Const FechaModificacion = "13/11/2014"
'Global Const UltimaModificacion = "Fernando Favre - CAS-27387 - FESTO - Adecuación de Recibo - rechazo 4. Modelo 177, se agrego buscar el tipo de estructura 20 - Lugar de Pago"

'Global Const Version = "5.90"
'Global Const FechaModificacion = "03/12/2014"
'Global Const UltimaModificacion = "Fernando Favre - CAS-26727 - Ingenio la Esperanza - Modificación Recibo de Haberes. Modelo 175, se agrego buscar una estructura y su fecha activa configurada por confrep"

'Global Const Version = "5.91"
'Global Const FechaModificacion = "29/12/2014"
'Global Const UltimaModificacion = "Sebastian Stremel - Se busco el campo basico en el co o ac configurado en la columna 375, ademas se busca la estructura lugar de pago - CAS-27164 - NGA - Críticos - Libro de Sueldos y Recibo de haberes San Miguel [Entrega 3]"

'Global Const Version = "5.92"
'Global Const FechaModificacion = "30/12/2014"
'Global Const UltimaModificacion = "Fernando Favre - CAS-28162 - IMECON QA -  Recibo de Sueldo"

'Global Const Version = "5.93"
'Global Const FechaModificacion = "16/01/2015"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-29080 - TPS - Bug en firma del recibo de haberes"
' Se modificó la consulta que obtiene la firma valida que tiene la empresa del modelo 85

'Global Const Version = "5.94"
'Global Const FechaModificacion = "27/01/2015"
'Global Const UltimaModificacion = "Miriam Ruiz - CAS-27164 - NGA - Críticos - Libro de Sueldos y Recibo de haberes San Miguel [Entrega 4]"
' Se agregó el modelo de liquidacion al recibo modelo=178

'Global Const Version = "5.95"
'Global Const FechaModificacion = "10/02/2015"
'Global Const UltimaModificacion = "Mauricio Zwenger - CAS-29011 - Se agregó firma de responsable RRHH y Nro de resolucion para Recibo Electronico"

'Global Const Version = "5.96"
'Global Const FechaModificacion = "24/02/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 22772 - Se modifico para que salga la descripcion del periodo y cantidad de horas trabajadas"
' Se modifico en el Modelo 126

'Global Const Version = "5.97"
'Global Const FechaModificacion = "02/03/2015"
'Global Const UltimaModificacion = "Carlos Masson - Solicitud de fuentes: CAS-20702 - CDA - Recibo de haberes Mexico [Entrega 5] (CAS-15298) - Se modifico el modelo 155 para que muestre en el log cuando no se encontrron configuradas algunas de las columnas requeridas"

'Global Const Version = "5.98"
'Global Const FechaModificacion = "15/04/2015"
'Global Const UltimaModificacion = "Borrelli Facundo - CAS-30420 - MONASTERIO BASE Gemplast - Bug en recibos de sueldo"
'Se modifico el Modelo 150 para que se muestre siempre la fecha de ingreso independiente del estado del empleado (activo/inactivo)

'Global Const Version = "5.99"
'Global Const FechaModificacion = "17/04/2015"
'Global Const UltimaModificacion = "Sebastian Stremel - Se creo el modelo 180 copia del modelo 35 de andreani para ASMSA - CAS-30200 - ASMSA - Custom modificación de recibo electrónico"
'

'Global Const Version = "6.00"
'Global Const FechaModificacion = "23/04/2015"
'Global Const UltimaModificacion = "Sebastian Stremel - Se modifico el modelo 01 estandar, se guarda en el campo auxchar3 que no se usaba el codigo de resolucion de la empresa - CAS-30003 - TSFOT - CUSTOM RECIBOS DIGITALES"

'Global Const Version = "6.01"
'Global Const FechaModificacion = "27/04/2015"
'Global Const UltimaModificacion = "Borrelli Facundo - CAS-30509 - SYKES - Error en estructuras de comprobante de pago"
'Para el Modelo 102 se modificaron las consultas para obtener la ubicacion y el departamento.

'Global Const Version = "6.02"
'Global Const FechaModificacion = "28/04/2015"
'Global Const UltimaModificacion = "Sebastian Stremel - CAS-30285 - ING. LA ESPERANZA - Custom antiguedad en recibo"
'Para el modelo 175 se cambia la forma de buscar la antiguedad, ahora informa lo mismo que se visualiza en el tablero.

'Global Const Version = "6.03"
'Global Const FechaModificacion = "12/05/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 30797 - Se creo el Modelo 181 para Apex El Salvador. Basado en el Modelo 102"
'Este Modelo esta basado en el Modelo 102 Sykes El Salvador

'Global Const Version = "6.04"
'Global Const FechaModificacion = "26/05/2015"
'Global Const UltimaModificacion = "Sebastian Stremel - Se buscan los dias trabajados configurados en el CO/AC columna 61 del modelo 175 Ing. La Esperanza"
'                                  CAS-30837 - ILE - Modificacion Recibo

'Global Const Version = "6.05"
'Global Const FechaModificacion = "08/06/2015"
'Global Const UltimaModificacion = "LED - se agrego descripcion del modelo de liquidacion a los modelos 171, 172, 173 y 174"
'                                 CAS-30693 - VSO - Interpack - Custom Recibos (CAS-26571) - Solicitud de ajustes

'Global Const Version = "6.06"
'Global Const FechaModificacion = "02/07/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - Se agrego la Fecha Antig Rec en el Modelo 178"
'                                 CAS-31632 - NGA - CITRICOS - Custom Recibo de Sueldo

'Global Const Version = "6.07"
'Global Const FechaModificacion = "03/07/2015"
'Global Const UltimaModificacion = "" 'Gonzalez Nicolás - CAS-23704 - SYKES El salvador -  Comprobante de pago (mensual) - 501 : Se agregó pliqanio a la query

'Global Const Version = "6.08"
'Global Const FechaModificacion = "13/07/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 30615 - Se agrego el Modelo 182 para el Cliente Dalkia, basado en el Modelo 20, solo que se le agrega Obra Social al Modelo 182"

'Global Const Version = "6.09"
'Global Const FechaModificacion = "14/07/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 30615 - Se agrego la llamada al Modelo 182 para el Cliente Dalkia"

'Global Const Version = "6.10"
'Global Const FechaModificacion = "15/07/2015"
'Global Const UltimaModificacion = "Fernandez, Matias - CAS-30746 - Prudential - Bug Pie Recibo- Modelo 39, cuando el cbu es 0 toma el numero de cuenta."

'Global Const Version = "6.11"
'Global Const FechaModificacion = "28/07/2015"
'Global Const UltimaModificacion = "Stremel Sebastian - Se agrego el modelo 183 - CAS-30003 - TSFOT - CUSTOM RECIBOS DIGITALES [Entrega 2]"

'Global Const Version = "6.12"
'Global Const FechaModificacion = "31/07/2015"
'Global Const UltimaModificacion = "Stremel Sebastian - Se modifico el modelo 141, se agrego la forma de liquidacion y la localidad de la empresa - CAS-32275 - MONASTERIO BASE AMR - Custom campo de recibo"

'Global Const Version = "6.13"
'Global Const FechaModificacion = "05/05/2015"
'Global Const UltimaModificacion = "Fernandez, Matias - CAS-32445 - NGA BASE CITRICOS - Bug en antigüedad de recibo, modificacion sobre la antiguedad para modelo 178"
                                  
'Global Const Version = "6.14"
'Global Const FechaModificacion = "12/08/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 31984 - Recibo La Caja - Modelo 184"


'Global Const Version = "6.15"
'Global Const FechaModificacion = "21/08/2015"
'Global Const UltimaModificacion = "Borrelli Facundo - CAS-32691 - NGA BASE CITRICOS - Bug en fechas de recibos"
'                               Se corrige la consulta para obtener la fecha de inicio y la fecha de ingreso del empleado.

'Global Const Version = "6.16"
'Global Const FechaModificacion = "24/08/2015"
'Global Const UltimaModificacion = "Miriam Ruiz - CAS-32521 - TECHNISYS - Modificación de Recibo de Haberes- se agregó al modelo 55 el sector"

'Global Const Version = "6.17"
'Global Const FechaModificacion = "25/08/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 31984 - Se modifico la Calificacion Personal y la Estructura Remuneracion para que se calcule en el Proceso"

'Global Const Version = "6.18"
'Global Const FechaModificacion = "28/08/2015"
'Global Const UltimaModificacion = "Miriam Ruiz- CAS-32521 - TECHNISYS - Modificación de Recibo de Haberes [Entrega 2]"
'                                   Se corrigió el cálculo de la antiguedad

'Global Const Version = "6.19"
'Global Const FechaModificacion = "01/09/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 31984 - Recibo La Caja Nro Modelo 184, se corrigio para que inserte bien el dato DESTINO"

'Global Const Version = "6.20"
'Global Const FechaModificacion = "07/09/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 31984 - Recibo La Caja Nro Modelo 184, se corrigio para que traiga bien las Estrcuturas Confidenciales y Generales"


'Global Const Version = "6.21"
'Global Const FechaModificacion = "08/09/2015"
'Global Const UltimaModificacion = "Mauricio Zwenger - CAS-31756 - GE - incorporación de un nuevo campo al recibo de sueldo de la empresa GE"
                                   ' se agregó campo SueldoConfirmado a columna auxdeci1 de modelo 158

'Global Const Version = "6.22"
'Global Const FechaModificacion = "11/09/2015"
'Global Const UltimaModificacion = "Borrelli Facundo - CAS-33027 - SANTANDER URUGUAY -  Error redondeo neto recibo"
                                  'Se cambio a Double el tipo de la variable valorTotal y sueldo.

'Global Const Version = "6.23"
'Global Const FechaModificacion = "28/09/2015"
'Global Const UltimaModificacion = "Miriam Ruiz - CAS-33196 - MIMO - Bug en recibo de autogestión"
                                  'Modelo 122 Mimo se corriguió el cuit de la empresa y el centro de costo.
                                  
'Global Const Version = "6.24"
'Global Const FechaModificacion = "29/09/2015"
'Global Const UltimaModificacion = "Stremel Sebastian - CAS-28162 - IMECON QA - Recibo de Sueldo [Entrega 2]"
                                  'Modelo 179 Correccion en RUC, direccion de la empresa.
                                  
'Global Const Version = "6.25"
'Global Const FechaModificacion = "07/10/2015"
'Global Const UltimaModificacion = "Stremel Sebastian - CAS-28162 - IMECON QA - Recibo de Sueldo [Entrega 4]"
                                  'Modelo 179 Correccion en periodo vacacional y mejora del proceso
                                  
'Global Const Version = "6.26"
'Global Const FechaModificacion = "13/10/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 31984 - Se modifico el Recibo Modelo Nro 184 (La Caja) para que muestre bien las Estructuras"
                                  
'Global Const Version = "6.27"
'Global Const FechaModificacion = "15/10/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 32670 - Se creo Modelo Nuevo Nro 185 para Uruguay"

'Global Const Version = "6.28"
'Global Const FechaModificacion = "29/10/2015"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-32922 - SEDAMIL - Recibo en PDF"
                                'Modelo 133 - Se modificó para busque el nro de resolucion configurada en la empresa, a la cual pertenece el empleado

'Global Const Version = "6.29"
'Global Const FechaModificacion = "09/11/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 32670 - Recibo Uruguay Modelo 185 - Se configuro el Confrep"

'Global Const Version = "6.30"
'Global Const FechaModificacion = "10/11/2015"
'Global Const UltimaModificacion = "Miriam Ruiz - CAS-28162 - IMECON QA - Recibo de Sueldo [Entrega 5] - Se modificó el modelo 179"
                               
'Global Const Version = "6.31"
'Global Const FechaModificacion = "13/11/2015"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS 32670 - Se corrigio la Query de Nro de Transaccion"

'Global Const Version = "6.32"
'Global Const FechaModificacion = "16/11/2015"
'Global Const UltimaModificacion = "Miriam Ruiz - CAS-28162 - IMECON QA - Recibo de Sueldo [Entrega 5] - Se modificó el modelo 179"

'Global Const Version = "6.33"
'Global Const FechaModificacion = "17/11/2015"
'Global Const UltimaModificacion = "Carmen Quintero - CAS-33605 - TATA - Modificacion de recibo - Se modificó el modelo 132, para que almacene la fecha de alta de reconocida"

'Global Const Version = "6.34"
'Global Const FechaModificacion = "25/11/2015"
'Global Const UltimaModificacion = "Miriam Ruiz- CAS-33601 - RH Pro (Producto) - Peru - Domicilio Boleta de Pago - se modificó la direccion de la empresa para el modelo 179 de Perú"

'Global Const Version = "6.35"
'Global Const FechaModificacion = "25/11/2015"
'Global Const UltimaModificacion = "Carmen Quintero- CAS-33605 - TATA - Modificacion de recibo [Entrega 2] - Se modificó origen de fecha de ingreso (alta ultima fase) y fecha de pago (Fecha planeada)."
                                  ' Se informa la fecha de baja y la causa de baja para el modelo 132


'Global Const Version = "6.36"
'Global Const FechaModificacion = "10/12/2015"
'Global Const UltimaModificacion = "Miriam Ruiz- CAS-34458 - BDO - Bug en campo de Recibo - Se creó el modelo 186 de agco."

'Global Const Version = "6.37"
'Global Const FechaModificacion = "21/12/2015"
'Global Const UltimaModificacion = "Fernandez, Matias-CAS-34143 – Santander Uruguay – Bug en Recibos de Sueldo - modelo  176 - fasrecofec = -1 en lugar de fecalta"

'Global Const Version = "6.38"
'Global Const FechaModificacion = "29/12/2015"
'Global Const UltimaModificacion = "Stremel Sebastian - CAS-33601 - RH Pro (Producto) - Peru - Boleta de Pago - Modificacion en modelo 179, se busca la estructura moneda"

'Global Const Version = "6.39"
'Global Const FechaModificacion = "21/01/2016"
'Global Const UltimaModificacion = "Dimatz Rafael CAS-31984 - LA CAJA - Custom Recibo de Haberes - Se corrigio para que salga el Nro de CBU o Nro de Cuenta segun el Banco"

'Global Const Version = "6.40"
'Global Const FechaModificacion = "04/02/2016"
'Global Const UltimaModificacion = "Ruiz Miriam - CAS-33601 - RH Pro ( Producto ) - Peru - Nueva Boleta de Pago - Se creo el modelo 187 para todo peru"

'se agregan campos a la tabla rep_recibo

'rep_recibo.add(auxdeci10).decimal(19,4).null;
'rep_recibo.add(auxdeci11).decimal(19,4).null;
'rep_recibo.add(auxdeci12).decimal(19,4).null;
'rep_recibo.add(auxdeci13).decimal(19,4).null;
'rep_recibo.add(auxdeci14).decimal(19,4).null;
'rep_recibo.add(auxdeci15).decimal(19,4).null;


'Global Const Version = "6.41"
'Global Const FechaModificacion = "12/02/2016"
'Global Const UltimaModificacion = "Ruiz Miriam - CAS-35556 - IHSA - Modificación de firma en recibo - Se creo el modelo 188 para IHSA"

'Global Const Version = "6.42"
'Global Const FechaModificacion = "18/02/2016"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS-35622 - PERSONAL - Nuevo Modelo de Recibo - Se creo nuevo Modelo 189 para Personal"

'Global Const Version = "6.43"
'Global Const FechaModificacion = "22/02/2016"
'Global Const UltimaModificacion = "Dimatz Rafael - CAS-35622 - PERSONAL - Nuevo Modelo de Recibo - Se corrigio la query de los documentos de empleados"

'Global Const Version = "6.44"
'Global Const FechaModificacion = "15/02/2016"
'Global Const UltimaModificacion = "LED - CAS-34811 - Monresa - Adec Recibo Digital y ESS - Se modifico el modelo 185 se guardar string en los campos auxchar"

'Global Const Version = "6.45"
'Global Const FechaModificacion = "26/02/2016"
'Global Const UltimaModificacion = "Borrelli Facundo - CAS-35449 - SANTANDER URUGUAY - Errores en recibo de sueldo [Entrega 2]"
                                'Se modifico el modelo 176, para mostrar en la fecha de ingreso la fecha de ultima fase activa

'Global Const Version = "6.46"
'Global Const FechaModificacion = "03/03/2016"
'Global Const UltimaModificacion = "Ruiz Miriam - CAS-33601 - RH Pro (Producto) - Peru - Nueva Boleta de Pago [Entrega 4]"
                                'se cambió la forma que busca la firma

'Global Const Version = "6.47"
'Global Const FechaModificacion = "26/04/2016"
'Global Const UltimaModificacion = "FMD - CAS-31625 - MONASTERIO BASE AMR - Adecuación en el recibo de haberes"
                                'Se cambio la leyenda Puesto por Categoria y el valor por la estructura categoria del empleado


'Global Const Version = "6.48"
'Global Const FechaModificacion = "27/04/2016"
'Global Const UltimaModificacion = "Gonzalez Nicolás - CAS-31984 - LA CAJA - Custom Recibo de Haberes"
                                  'Se busca Fecha de alta reconocida

'Global Const Version = "6.49"
'Global Const FechaModificacion = "28/04/2016"
'Global Const UltimaModificacion = "Gonzalez Nicolás - CAS-31984 - LA CAJA - Custom Recibo de Haberes"
                                  'Se busca Fecha de alta reconocida de la fase inactiva
                                  
Global Const Version = "6.50"
Global Const FechaModificacion = "29/04/2016"
Global Const UltimaModificacion = "FMD - CAS-33196 - MIMO - Bug en recibo de haberes"
                                  'Cambio para que se muestre la antiguedad en vez del CUIL en el recibo del aplicativo
'----------------------------------------------------------

