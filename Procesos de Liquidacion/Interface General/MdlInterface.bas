Attribute VB_Name = "MdlInterface"
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------
'*********************************************************** AVISO *********************************************************************
'**************************************************** VERSIONES NO LIBERADAS **************************************************
'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------
'Global Const Version = "6.0X"
'Global Const FechaModificacion = "XX/XX/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-31503 - RH Pro (Producto) - Interfase de dias correspondientes R4
'                                 Se permite la carga de dias con decimales - modelo 2004


'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------
'*********************************************************** AVISO *********************************************************************
'*********************************** LA DESCRIPCION DE LA ULTIMA VERSION SIGUE A CONTINUACION ******************************************
'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------

Global Const Version = "6.38"
Global Const FechaModificacion = "04/05/2016"
Global Const UltimaModificacion = "FMD - CAS-37038 - Mejora interface 921 - IBT GROUP - En la interfaz 921, se quito obligatoriedad en idiomas"

'Global Const Version = "6.37"
'Global Const FechaModificacion = "29/04/2016"
'Global Const UltimaModificacion = "FMD - CAS-36340 - MONASTERIO BASE BLD - Bug en interfaz 620 - En la interfaz 620 se corrigio bug para datos N/A y Log"

'Global Const Version = "6.36"
'Global Const FechaModificacion = "28/04/2016"
'Global Const UltimaModificacion = "FMD - CAS-36340 - MONASTERIO BASE BLD - Bug en interfaz 620 - En la interfaz 620 se corrigio bug para datos N/A introducido en version anterior"

'Global Const Version = "6.35"
'Global Const FechaModificacion = "27/04/2016"
'Global Const UltimaModificacion = "FMD - CAS-36340 - MONASTERIO BASE BLD - Bug en interfaz 620 - Se controlo el separador en los items interface 620"

'Global Const Version = "6.34"
'Global Const FechaModificacion = "'11/02/2016"
'Global Const UltimaModificacion = "" 'Dimatz Rafael CAS-35413 - SOLAR - Modelo de domicilio - Se agrego Longitud y Latitud al Modelo 668 Paraguay

'Global Const Version = "6.33"
'Global Const FechaModificacion = "'29/01/2016"
'Global Const UltimaModificacion = "" 'Fernandez, Matias -CAS-32751 - LA CAJA - Custom Seguros ADP-Se controla los tipos de seguro y de beneficiario

'Global Const Version = "6.32"
'Global Const FechaModificacion = "'19/01/2016"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS-34846 - UTDT - Bug en interfaz 230 - Correccion en acreditacion de dias a periodos
                                     'Correspondientes.

'Global Const Version = "6.31"
'Global Const FechaModificacion = "'12/01/2016"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-33601 - RH Pro (Producto) - Peru - Interfaz modelo 922
                                     'Se controla que exista el tipo de domicilio para el mod 922.

'Global Const Version = "6.30"
'Global Const FechaModificacion = "'22/12/2015"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-32751 - LA CAJA - Custom Seguros ADP
'                                     'Se agreg� el modelo 404 Importaci�n de Seguros (Custom)
                                     
'Global Const Version = "6.29"
'Global Const FechaModificacion = "'16/12/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz- CAS-33789 - RHPro - Bug en interfaz 630 [Entrega 3]
'                                     'Se modifico la query que busca las obras sociales

'Global Const Version = "6.28"
'Global Const FechaModificacion = "'11/12/2015"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-33601 - RH Pro (Producto) - Peru - Interfaz modelo 922
                                     'Se modific� funci�n TraerTelefonosMask  Controla para tel�fonos enmascarados en campo telnro

'Global Const Version = "6.27"
'Global Const FechaModificacion = "10/12/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-28226 - Raffo - Interface Jefe y Jefe del Jefe (Roles de Evaluaci�n)
                                     'Nuevo modelo 2009 - Roles de eventos

'Global Const Version = "6.26"
'Global Const FechaModificacion = "'10/12/2015"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-33601 - RH Pro (Producto) - Peru - Interfaz modelo 922
                                     '922 : C�d. telef�nicos para PERU


'Global Const Version = "6.25"
'Global Const FechaModificacion = "'02/12/2015"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-33601 - RH Pro (Producto) - Peru - Interfaz modelo 922
                                     '922 : Se marcan como principales los tel�fonos de tipo 1 (S�lo v�lido para PERU)

'Global Const Version = "6.24"
'Global Const FechaModificacion = "'19/11/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-33789 - RHPro - Bug en interfaz 630
                                     'se corrigi� para que no insertara dos planes duplicados
             

'Global Const Version = "6.23"
'Global Const FechaModificacion = "'19/11/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-33601 - RH Pro (Producto) - Peru - Bug Estudios Formales
                                     'Se modific� la interfaz 296 para que no permita cargar dos carreras con el mismo nombre
                                

'Global Const Version = "6.22"
'Global Const FechaModificacion = "'13/11/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-33789 - RHPro - Bug en interfaz 630
                                     'Se modific� la interfaz 630 para que no permita cargar un plan de obra social si el empleado no tiene una obra social asignada
                                     

'Global Const Version = "6.21"
'Global Const FechaModificacion = "'10/11/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-33972 - BDO Per� - Bug Asignaci�n Nro. Empleados
                                     'Se agreg� empresa a la interface 672

'Global Const Version = "6.20"
'Global Const FechaModificacion = "'16/10/2015"
'Global Const UltimaModificacion = "" 'Dimatz Rafael - CAS 32670 - Recibo Digital Uruguay - Nro de Transaccion de Pedido de Pago por Empleado
                                     'Se crea una nueva Interface 2007

'Global Const Version = "6.19"
'Global Const FechaModificacion = "09/10/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-28350 - Salto Grande - Custom ADP - Organigrama Funcional [Entrega 3]
                                     'Cambio en el modelo 400, se permite cargar nodos sin padre ubicados en cualquier nivel.

'Global Const Version = "6.18"
'Global Const FechaModificacion = "08/10/2015"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-31268 - IBT - CUSTOM PARTE DE CAMBIO DE TURNO [Entrega 3]
                                     'Se agrego validacion en el modelo 402 Importaci�n de Partes de Cambio de Turno (Custom)

'Global Const Version = "6.17"
'Global Const FechaModificacion = "28/09/2015"
'Global Const UltimaModificacion = "" 'Borrelli Facundo - Borrelli Facundo - CAS-32386 - Telefax (Santander URU) - Bug Interface 354 [Entrega 2](CAS-15298)
                                     'Se modifica la interfaz 354 para que no se generen errores al utilizar los formatos 2 y 6.

'Global Const Version = "6.16"
'Global Const FechaModificacion = "04/09/2015"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-31268 - IBT - CUSTOM PARTE DE CAMBIO DE TURNO [Entrega 2]
                                     'Modificaciones varias al modelo 402 Importaci�n de Partes de Cambio de Turno (Custom)
                                     
'Global Const Version = "6.15"
'Global Const FechaModificacion = "31/08/2015"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-32562 - RH Pro (Producto) - H&A - Modelos para Tel�fonos
                                     'Modelo 1006 - Se agreg� carga del c�digo telef�nico (el formato es din�mico)
                                     'Modelo 668 y 922: Se amplio Telefono1 y Telefono2 a 60

'Global Const Version = "6.14"
'Global Const FechaModificacion = "28/08/2015"
'Global Const UltimaModificacion = "" ' Miriam Ruiz - CAS-32730 - SANTANDER URUGUAY - Migracion de licencias en dias habiles
                                     'se modifica model 2003 - se agrega el control sobre dias corridos

'Global Const Version = "6.13"
'Global Const FechaModificacion = "21/08/2015"
'Global Const UltimaModificacion = "" ' LED - CAS-28350 - Salto Grande - Custom ADP - Organigrama Funcional
                                     'nuevo modelo 400 - Importaci�n de dependencia de estructuras
                                     
'Global Const Version = "6.12"
'Global Const FechaModificacion = "14/08/2015"
'Global Const UltimaModificacion = "" ' Sebastian Stremel - CAS-30813 - Salto Grande - ADP - Mejoras de funcionalidad (CAS-15298) [Entrega 3] - Se movio el modelo 676 y 677 para CTM a las interfaces Custom


'Global Const Version = "6.11"
'Global Const FechaModificacion = "13/08/2015"
'Global Const UltimaModificacion = "" 'MDZ - CAS-29842 - GE - Agrandar campo Domicilio
'               Se amplio campo Calle a 250 caracteres  en todos los modelos que lo utilizan


'Global Const Version = "6.10"
'Global Const FechaModificacion = "12/08/2015"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-31268 - IBT - CUSTOM PARTE DE CAMBIO DE TURNO
'                     Se agreg� el modelo 402 Importaci�n de Partes de Cambio de Turno (Custom)


'Global Const Version = "6.09"
'Global Const FechaModificacion = "11/08/2015"
'Global Const UltimaModificacion = "" 'Borrelli Facundo - CAS-32386 - Telefax (Santander URU) - Bug Interface 354
'                    Se agrega la posibilidad de cargar la fecha hasta como N/A, para cuando no se sabe cuando finaliza la vigencia

'Global Const Version = "6.08"
'Global Const FechaModificacion = "10/08/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-32204 - G.COMPARTIDA - Bug en Interfaz 312
'                    Se agreg� en el modelo 312 el control sobre el n�mero de tarjeta

'Global Const Version = "6.07"
'Global Const FechaModificacion = "05/08/2015"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-30758 - ACARA - SO - ABM de visitas medicas y nueva interfaz de importaci�n
'                     Se agrego modelo 403 Importacion de visitas medicas

'Global Const Version = "6.06"
'Global Const FechaModificacion = "24/07/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-28245 - MEGATLON -  DSITRIBUCION CONTABLE POR CONCEPTO [Entrega 2]
'                     Correcion en el modelo 288 se chequea el tipo de las estructuras con las del modelo de asiento junto con la descripcion.

'Global Const Version = "6.05"
'Global Const FechaModificacion = "17/07/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-28350 - Salto Grande - Custom ADP � Importaci�n masiva de archivos
'                                 Nuevo modelo de interfaz 2005 - Importaci�n masiva de archivos

'Global Const Version = "6.04"
'Global Const FechaModificacion = "02/07/2015"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-22072 - Raffo - Adecuaciones GDD - Migraci�n historico People Review
                                     'Se agreg� el modelo 401 Importaci�n de Hist�rico de People Review (Custom)
'Global Const Version = "6.03"
'Global Const FechaModificacion = "29/06/2015"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-17053 - Nac Brasil � Nacionalizaci�n Brasil - Reporte ESOCIAL
                                     'Se agreg� Tipo de Logradouro para modelo de BRASIL : Interfaz 668 y 922

'Global Const Version = "6.02"
'Global Const FechaModificacion = "22/06/2015"
'Global Const UltimaModificacion = "" ' LED - CAS-31503 - RH Pro (Producto)
'                                 Interfase de dias correspondientes R4 - Nuevo modelo 2004 (Importacion Dias Correspondientes GIV R4)

'Global Const Version = "6.01"
'Global Const FechaModificacion = "08/06/2015"
'Global Const UltimaModificacion = "" ' Borrelli Facundo - CAS-31286 - COMPA�IA DE ALIMENTOS - Bug en interfaz 245
'                                   Se corrigio error de Typemismatch al subir la interfaz con Formato 1, adem�s se corrigieron
'                                   los mensajes en el log, cuando se generan errores al importar datos.

'Global Const Version = "6.00"
'Global Const FechaModificacion = "03/06/2015"
'Global Const UltimaModificacion = "" ' FGZ - CAS-31298 - Telefax (Santander URU) - Modificacion Interfase 233
'                                   Se agreg� al modelo 2003 - Importaci�n de Licencias con control de d�as habiles
'
'   Ademas
'                                  CAS-30739 - SYKES EL SALVADOR - LIQ - Bug interfase 354
'                                   Modelo 354 (se cambi� validacion de fechas).


'Global Const Version = "5.99"
'Global Const FechaModificacion = "26/05/2015"
'Global Const UltimaModificacion = "" ' Miriam Ruiz - CCAS-30722 - RH Pro - Libro Registro de Horas Suplementarias [Entrega 2]
''                                      Se corrigi� insert de la interface 2001,guardaba solo la hora desde
''                                      Se corrigi� la interface 2002 , insertaba mal la hora en decimal

'Global Const Version = "5.98"
'Global Const FechaModificacion = "22/05/2015"
'Global Const UltimaModificacion = "" ' Sebastian Stremel - CAS-30813 - Salto Grande - ADP - Mejoras de funcionalidad - Se agrego el modelo 676 y 677 para CTM



'Global Const Version = "5.97"
'Global Const FechaModificacion = "21/05/2015"
'Global Const UltimaModificacion = "" ' Miriam Ruiz - CAS-30722 - RH Pro - Libro Registro de Horas Suplementarias - entrega 2
'                                   Se agreg� al modelo 2002 - Interface de Horario Cumplido - el campo 'horestado' al insert


'Global Const Version = "5.96"
'Global Const FechaModificacion = "15/05/2015"
'Global Const UltimaModificacion = "" ' Miriam Ruiz - CAS-30722 - RH Pro - Libro Registro de Horas Suplementarias
'                                   Se agreg� el modelo 2002 - Interface de Horario Cumplido

'Global Const Version = "5.95"
'Global Const FechaModificacion = "22/04/2015"
'Global Const UltimaModificacion = "" ' Borrelli Facundo - CAS-21778 - Sykes El Salvador - QA � Bug Interface 354 [Entrega 3]
'Se mueve de lugar la validacion de fechas, al final de las validaciones.

'Global Const Version = "5.94"
'Global Const FechaModificacion = "16/04/2015"
'Global Const UltimaModificacion = "" ' Borrelli Facundo - CAS-21778 - Sykes El Salvador - QA � Bug Interface 354 [Entrega 2]
'Se modifica la validacion para que fecha desde sea menor a la fecha hasta.


'Global Const Version = "5.93"
'Global Const FechaModificacion = "13/04/2015"
'Global Const UltimaModificacion = "" ' Sebastian Stremel - CAS-30411 - Salto Grande -  Creaci�n de Interfase importaci�n de  Nick Names
'Se crea el modelo 675 de migracion inicial


'Global Const Version = "5.92"
'Global Const FechaModificacion = "13/04/2015"
'Global Const UltimaModificacion = "" ' Borrelli Facundo - CAS-21778 - Sykes El Salvador - QA � Bug Interface 354
'Se valida que la fecha desde sea menor a la fecha hasta.

'Global Const Version = "5.91"
'Global Const FechaModificacion = "09/04/2015"
'Global Const UltimaModificacion = "" ' LED - CAS-25180 - Uni�n Personal - Desarrollo de nueva interfaz para EyP [Entrega 4] - Se agrego nivel de estudio
'       Modelo 392 - cambio en el campo de personal a cargo ahora es alfanumerico, y el nivel educacional deja de ser una estructura y se controla con el nivel de estudio.


'Global Const Version = "5.90"
'Global Const FechaModificacion = "30/03/2015"
'Global Const UltimaModificacion = "" ' EAM - CAS-16645 - PERU - TSS- Nacionalizacion- Adecuacion de GIV
'       Modelo 298 - Se sacaron validaciones para cargar las licencias de vacaciones
'       MDZ - se convierten los dias de licencia al tipo indicado, si se indica el tipo a convertir
'---------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------
'Global Const Version = "5.89"
'Global Const FechaModificacion = "27/03/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-16441 - H&A - PERU - Nacionalizacion Per� - Modificaci�n interfaz 672
'                                     ' Se modific�0 la interfaz 672 se agregaron dos campos si el contrato tiene per�odo a prueba
'                                     ' y fecha de fin de prueba, en caso de que la misma venga vac�a se calculan 90 d�as a paratir de la fecha de alta del empleado


'Global Const Version = "5.88"
'Global Const FechaModificacion = "16/03/2015"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-26755 - MEDICUS - CUSTOM INTERFAZ DE NOVEDADES HORARIAS [Entrega 4]
                                     ' Se reemplaza el separador decimal por un punto para que no rompa el insert en modelo 394 y 396


'Global Const Version = "5.87"
'Global Const FechaModificacion = "12/03/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-16441 - H&A - PERU - Nacionalizacion Per� - Modificaci�n interfaz 672 [Entrega 3]
                                     ' Se agreg� control sobre el tipo de estructura cuando ven�a N/A
                                     


'Global Const Version = "5.86"
'Global Const FechaModificacion = "06/03/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-28245 - MEGATLON -  DSITRIBUCION CONTABLE POR CONCEPTO
                                     ' Se modifico la interfaz 288 - se nivelo el funcionamiento con el ABM desde el sistema
                                     ' se permite cargar novedades que solo difieran en las estructuras de la distribucion

'Global Const Version = "5.85"
'Global Const FechaModificacion = "03/03/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-28350 - Salto Grande - Custom ADP - Reporte de Familiares-Inteface
                                     ' Se modific{o la interfaz 398 para que no sea obligatorio el campo fecha de vencimiento
                                     ' quedan como obligatorios solo los siguientes campos:
                                     'Legajo, nombre, apellido, fecha de nacimiento, pa�s de nacimiento, nacionalidad, estado civil, Sexo, parentesco, Tipo Documento, Nro documento y A CARGO CTM

'Global Const Version = "5.84"
'Global Const FechaModificacion = "02/03/2015"
'Global Const UltimaModificacion = "" 'seba - CAS-26755 - MEDICUS - CUSTOM INTERFAZ DE NOVEDADES HORARIAS [Entrega 3]
                                     ' correcion en la insercion en gti_novedadhoraria de fechaprocesamiento.


'Global Const Version = "5.83"
'Global Const FechaModificacion = "20/02/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-26755 - MEDICUS - CUSTOM INTERFAZ DE PARTES [Entrega 4] - modificacion modelo 396, el campo cantidad horas permite decimales
                                     ' y se agrego columna autorizable.
                                     
'Global Const Version = "5.82"
'Global Const FechaModificacion = "19/02/2015"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS-29429 - Salto Grande - QA- bug importaci�n interfaz 387 - si el documento no
                                     ' es unico, se puede cargar a mas de un empleado



'Global Const Version = "5.81"
'Global Const FechaModificacion = "12/02/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-16441 - H&A - PERU - Nacionalizacion Per� - Modificaci�n interfaz 672 [Entrega 3]
'                                     Se cambi� a fecha de fin de per�odo de prueba
'                                     a 89 d�as despues de la fecha de fin del contrato

'Global Const Version = "5.80"
'Global Const FechaModificacion = "11/02/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-25180 - Uni�n Personal - Desarrollo de nueva interfaz para EyP [Entrega 3]
'                                     Correciones en el modelo 392 - Se agregaron logs en archivo de errores y de lineas procesadas
'                                     la columna nivel educacional (13) es una estructura configurable  y la columna personal a cargo (18) es un texto libre numerico

'Global Const Version = "5.79"
'Global Const FechaModificacion = "03/02/2015"
'Global Const UltimaModificacion = "" ' Miriam Ruiz - CAS-16441 - H&A - PERU - Nacionalizacion Per� - Modificaci�n interfaz 672 -
'                                     Si el pais es per� cuando se da el alta del contrato calcula la fecha de fin de per�odo de prueba
'                                     como 90 dias despues de la fecha de fin del contrato

'Global Const Version = "5.78"
'Global Const FechaModificacion = "29/01/2015"
'Global Const UltimaModificacion = "" ' Miriam Ruiz - CAS-16441 - H&A - PERU - Nacionalizacion Per� - Modificaci�n interfaz 672 -
'                                     Se agregaron los campos de estudio y fecha de fin de prueba

'Global Const Version = "5.77"
'Global Const FechaModificacion = "22/01/2015"
'Global Const UltimaModificacion = "" ' Fernandez, Matias - CAS-27978 - NGA - San Miguel - interfaz 253 - correccion condicional de quincena


'Global Const Version = "5.76"
'Global Const FechaModificacion = "22/01/2015"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-26755 - MEDICUS - BUG EN INTERFAZ PEDIDOS DE VACACIONES -
                                     ' - Se cualificaron los campos de la consulta ubicada en la funcion sup_ped_lic utilizada en el modelo 230.
'Global Const Version = "5.75"
'Global Const FechaModificacion = "19/01/2015"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-25641 - VSO - PRAXAIR BOLIVIA - Interface Empleados -
'                                     ' - Se controlan caracteres invalidos en nombre y apellido. Se agreg� terape2 y ternom2 al insert de empleado.
'Global Const Version = "5.74"
'Global Const FechaModificacion = "19/01/2015"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-28350 - Salto Grande - Custom ADP - Reporte de Familiares-Inteface - Se agrega el modelo 398


'Global Const Version = "5.73"
'Global Const FechaModificacion = "13/01/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-25180 - Uni�n Personal - Desarrollo de nueva interfaz para EyP
'                                   ' Cambio en el modelo 392, se quitaron campo y se agregaron otros (ver formularios del caso).

'Global Const Version = "5.72"
'Global Const FechaModificacion = "09/01/2015"
'Global Const UltimaModificacion = "" 'Dimatz Rafael - CAS 16441 - Peru - Se modifico, ademas de verificar la cuenta si son iguales, verifique si los Tipos son iguales
'                                   ' Modelo de Interfaz 671

'Global Const Version = "5.71"
'Global Const FechaModificacion = "08/01/2015"
'Global Const UltimaModificacion = "" 'LED - CAS-28695 - NGA - Citricos - Interface I-Time
'                                   ' Nuevo modelo de interfaz de importacion de novedades de liq custom (397)


'Global Const Version = "5.70"
'Global Const FechaModificacion = "02/01/2015"
'Global Const UltimaModificacion = "" 'Ruiz Miriam - CAS-28707 - NGA - Iron Mountain � Inconveniente en columnas interfaz 317
'                                   ' se agreg� fecha de vigencia desde y hasta a la interfaz 317


'Global Const Version = "5.69"
'Global Const FechaModificacion = "30/12/2014"
'Global Const UltimaModificacion = "" 'Dimatz Rafael - CAS-22751 - Peru - Agregar Pais por Nro de Codigo


'Global Const Version = "5.68"
'Global Const FechaModificacion = "29/12/2014"
'Global Const UltimaModificacion = "" 'FGZ - CAS-26028 - H&A - Mejora en indicadores individuales
                                     'se modific� la validacion de cantidad de campos del modelo 2000.



'Global Const Version = "5.67"
'Global Const FechaModificacion = "23/12/2014"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS-27978 - NGA - San Miguel - interfaz 253 [Entrega 2] (CAS-15298)
'                                     'se acomodaron diferentes condicionales y mensajes de errores.
                        

'Global Const Version = "5.66"
'Global Const FechaModificacion = "23/12/2014"
'Global Const UltimaModificacion = "" ' Gonzalez Nicol�s - CAS-27929 - UTDT - Planilla Horaria - Nuevo modelo 2001 - Interfaz de Planilla Horaria



'Global Const Version = "5.65"
'Global Const FechaModificacion = "23/12/2014"
'Global Const UltimaModificacion = "" ' Dimatz Rafael - CAS 16441 - Se inserta en tercero los apellidos y nombres, se agrego split por @ para apellido y nombre
'       Modelo 299 - Interface Familiar


'Global Const Version = "5.64"
'Global Const FechaModificacion = "19/12/2014"
'Global Const UltimaModificacion = "" ' Dimatz Rafael - CAS-22751 - VILLA MARIA - Adecuacion Empleos y Postulantes - Modificaci�n Interfaz 921
'       Modelo 921 - Se corrigieron las funciones que traen Localidad, distrito, departamento, pais, institucion, carrera, titulo


'Global Const Version = "5.63"
'Global Const FechaModificacion = "19/12/2014"
'Global Const UltimaModificacion = "" 'Ruiz Miriam - CAS-26028 - H&A - Mejora en indicadores individuales - Se agreg� la interface 2000
                                     'para cargar indicadores individuales


'Global Const Version = "5.62"
'Global Const FechaModificacion = "11/12/2014"
'Global Const UltimaModificacion = "" 'Fernandez,Matias - CAS-27978 - NGA - San Miguel - interfaz 253 - se corrigio el condicional en mes de
                                     'finalizacion


'Global Const Version = "5.61"
'Global Const FechaModificacion = "05/12/2014"
'Global Const UltimaModificacion = "" ' LED - CAS-26582 - GE - GPIT Nueva funcionalidad de interfaz [Entrega 2]
''       Modelo 293 - Cuando la interfaz encuentra una novedad sin vigenvia y se quiere reemplazar se pisa sin cerrar fechas
''       Modelo 335 - Cuando la interfaz encuentra una novedad exactamente igual, actualiza la fecha de procesamiento

'Global Const Version = "5.60"
'Global Const FechaModificacion = "05/12/2014"
'Global Const UltimaModificacion = "" 'Dimatz Rafael - Modelo 921 - CAS 22751 - Se Modifico la Fecha Desde de Estudios Formales


'Global Const Version = "5.59"
'Global Const FechaModificacion = "05/12/2014"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - modelo 275, se permite guardar hasta 500 caracteres en el campo tareas desempe�adas - CAS-28074 - IHSA - Planificador Empleos y Postulantes
                                     


'Global Const Version = "5.58"
'Global Const FechaModificacion = "04/12/2014"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s -CAS-16441 - H&A - PERU - Nacionalizacion Per� - Modificaci�n de la interfaz de familiares y domicilios
                                     ' Modelo 299 - Interfaz de Familiares (Sin domicilios)
                                     ' Modelo 922 - Interfaz de Domicilios (por documento)

'Global Const Version = "5.57"
'Global Const FechaModificacion = "01/12/2014"
'Global Const UltimaModificacion = "" ' EAM - sse Agreg� un nuevo modelo para R4 de Licencia GIV
                                    ' Modelo 298 - EAM- 27/11/2014 - Nuevo modelo para la carga de licencias de vacaciones R4 Cas "CAS-16645 - PERU - TSS- Nacionalizacion- Adecuacion de GIV"

'Global Const Version = "5.56"
'Global Const FechaModificacion = "01/12/2014"
'Global Const UltimaModificacion = "" ' LED - se quito horas autorizable del modelo 396
'                                    ' CAS-26755 - MEDICUS - CUSTOM INTERFAZ DE PARTES [Entrega 2]

'Global Const Version = "5.55"
'Global Const FechaModificacion = "27/11/2014"
'Global Const UltimaModificacion = "" ' Gonzalez Nicol�s - Modelo de Honduras interfaz 668 y 1006

'Global Const Version = "5.54"
'Global Const FechaModificacion = "25/11/2014"
'Global Const UltimaModificacion = "" ' Dimatz Rafael - CAS 22751 - Villa Maria - Se agrego funcion TraerCodInstitucionDescripcion
'       Modelo 921 - Se cambio para comparar la institucion con la Descripcion y se agrego 90 caracteres a institucion y carrera

'Global Const Version = "5.53"
'Global Const FechaModificacion = ""
'Global Const UltimaModificacion = "" ' LED - CAS-26755 - MEDICUS - CUSTOM INTERFAZ DE PARTES
'       Modelo 396 - importacion de partes de autorizacion de horas sin firmas, funciona igual a la sykes modelo 336 pero sin firmas y con mantiene, pisa o suma

'Global Const Version = "5.52"
'Global Const FechaModificacion = "20/11/2014"
'Global Const UltimaModificacion = "" ' CAS-26755 - MEDICUS - CUSTOM INTERFAZ DE NOVEDADES HORARIAS
                                     ' Sebastian Stemel - CAS-26755 - MEDICUS - CUSTOM INTERFAZ DE NOVEDADES HORARIAS
                                     ' Se creo el modelo 394 que importa novedades horarias.


'Global Const Version = "5.51"
'Global Const FechaModificacion = "19/11/2014"
'Global Const UltimaModificacion = "" ' CAS-16441 - H&A - PERU - PVS T-REGISTRO - Cambio Legal [Entrega 3]
                                     ' Modificacion en modelo 297, error en el a�o de la carrera.
                                     ' si el usario tiene perfil(3) de rrhh marca la carrera como el estudio rrhh aprobado
                                     ' Sebastian Stremel


'Global Const Version = "5.50"
'Global Const FechaModificacion = "18/11/2014"
'Global Const UltimaModificacion = "" ' CAS-25708 - VILLA MARIA - DOMICILIO PERU [Entrega 3]
                                     ' para las localidades, provincias y departamentos si el dato existe se modifica el cod externo y el ubigeo
                                     ' Sebastian Stremel


'Global Const Version = "5.49"
'Global Const FechaModificacion = "17/11/2014"
'Global Const UltimaModificacion = "" ' Fernandez, Matias - CAS-27978 - NGA - San Miguel - interfaz 253
                                     'se realizaron varios controles sobre las variables, y el string se paso a un
                                     ' arreglo por medio de la funcion split


'Global Const Version = "5.48"
'Global Const FechaModificacion = "13/11/2014"
'Global Const UltimaModificacion = "" ' Sebastian Stremel - se modifico las funciones donde cortaba el codext de las estructuras a 20 caracteres, para que permita 30 caracteres.
                                     ' CAS-25884 - ALTHAUS - ADP - Espacios Cod Externo Estructura

'Global Const Version = "5.47"
'Global Const FechaModificacion = "12/11/2014"
'Global Const UltimaModificacion = "" ' Sebastian Stremel - Se creo el modelo 395 - Importacion de dias de turismo
                                     ' CAS-26789 - Santander Uruguay - D�as de turismo - Licencia Paga - Licencia 25 a�os


'Global Const Version = "5.46"
'Global Const FechaModificacion = "12/11/2014"
'Global Const UltimaModificacion = "" ' Miriam Ruiz  - CAS-27476 - Heidt & Asociados - Interfaz 615 y 620 de Impuesto a las Ganancias
'           Modelo 620 - se modifica para que pise si la fecha es la misma



'Global Const Version = "5.45"
'Global Const FechaModificacion = "31/10/2014"
'Global Const UltimaModificacion = "" ' LED - CAS-26582 - GE - GPIT Nueva funcionalidad de interfaz [Entrega 2]
'       Modelo 293 - Cuando la interfaz encuentra una novedad sin vigenvia y se quiere reemplazar se pisa sin cerrar fechas
'       Modelo 335 - Cuando la interfaz encuentra una novedad exactamente igual, actualiza la fecha de procesamiento

'Global Const Version = "5.44"
'Global Const FechaModificacion = "27/10/2014"
'Global Const UltimaModificacion = "" ' LED - CAS-27511 - Sykes El Salvador - Retroactivo Nocturnidad [Entrega 2]
       'Correcion en el insert de la novedad con la fecha de procesamiento

'Global Const Version = "5.43"
'Global Const FechaModificacion = "27/10/2014"
'Global Const UltimaModificacion = "" ' Carmen Quintero - CAS-27632 - NGA - Error en descripcion de convenio
       'Se modifico la funcion EliminarCHInvalidosII por agregar el caracter "�" como valido

'Global Const Version = "5.42"
'Global Const FechaModificacion = "24/10/2014"
'Global Const UltimaModificacion = "" 'CAS-27511 -  Sykes El Salvador - Retroactivo Nocturnidad
       'Se modifico el modelo 335 para sykes - SV. Se agrego el campo fecha de procesamiento para el calculo de retroactividad.

'Global Const Version = "5.41"
'Global Const FechaModificacion = "21/10/2014"
'Global Const UltimaModificacion = "" 'CAS-26582 - GE - Nuevo formato interfaz Novedades
'       'Cambio en el modelo 293, se reestructuro la interfaz y se agregaron campos

'Global Const Version = "5.40"
'Global Const FechaModificacion = "20/10/2014"
'Global Const UltimaModificacion = "" 'CAS-27184 - DELOITTE & CO.S.R.L. (I y III)   - Error en Interfaz Import. Distrib. Contable.
       'Se permite el caracter "&" en el modelo de interfaz 317


'Global Const Version = "5.39"
'Global Const FechaModificacion = "15/10/2014"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-26942 - SYKES - ERROR EN CARGA DE INTERFAZ (CAS-15298)
       'Se modifica la consulta que verifica si existe superposici�n en el modelo 341
    
'Global Const Version = "5.38"
'Global Const FechaModificacion = "14/10/2014"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-16441 - H&A - PERU - PVS T-REGISTRO - Cambio Legal [Entrega 3]
       'Se creo el modelo 297 el cual importa estudios formales del empleado.
       'Se creo el modelo 296 el cual importa los datos parametricos del t-registro.

'Global Const Version = "5.37"
'Global Const FechaModificacion = "25/09/2014"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-16441 - H&A - PERU - PVS T-REGISTRO - Cambio Legal
       'Se creo el modelo 294 importacion de instituciones y 295 importacion de carreras



'Global Const Version = "5.36"
'Global Const FechaModificacion = "24/09/2014"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS-27132 - ACARA - bug en interfaz 230 pedido de vacaciones
       'Se redise�o la interfaz 230, se dejo la que estaba como actual como OLD...




'Global Const Version = "5.35"
'Global Const FechaModificacion = "19/09/2014"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS-27132 - ACARA - bug en interfaz 230 pedido de vacaciones
       'se volvio el cambio de la version 5.34 y se agrego mas info al log para poder identificar el problema.


'Global Const Version = "5.34"
'Global Const FechaModificacion = "16/09/2014"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS-27132 - ACARA - bug en interfaz 230 pedido de vacaciones
       'se armo una lista de periodos a excluir en el analisis, para no verlos mas de una vez.





'Global Const Version = "5.33"
'Global Const FechaModificacion = "03/09/2014"
'Global Const UltimaModificacion = "" 'Miriam Ruiz -CAS-27055 - NGA - Error en interfaz 246
'                   Se modific� modelo 246 se sac� la fecha de vigencia y se solucion� el problema con los decimales


'Global Const Version = "5.32"
'Global Const FechaModificacion = "29/08/2014"
'Global Const UltimaModificacion = "" 'Miriam Ruiz - CAS-17053 - Nac Brasil � Nacionalizaci�n Brasil - reporte Manad
'                   Se modific� modelo 668  y el 1006, se agreg� municipio para Brasil


'Global Const Version = "5.31"
'Global Const FechaModificacion = "29/08/2014"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-26942 - SYKES - ERROR EN CARGA DE INTERFAZ
'                   Se modifica modelo 341 por faltar enviar a ejecutar la consulta en la base, para el caso cuando se debe actualizar la novedad con vigencia del empleado


'Global Const Version = "5.30"
'Global Const FechaModificacion = "28/08/2014"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-25708 - VILLA MARIA - DOMICILIO PERU
'                   se agrega al modelo 1006 el campo auxchar para peru donde se grabara el ubigeo
'                   Se modifica modelo 668 importacion de domicilios MultiPais - Se agregan los campos necesarios para el domicilio de peru


'Global Const Version = "5.29"
'Global Const FechaModificacion = "13/08/2014"
'Global Const UltimaModificacion = "" 'MDZ - CAS-26582 - GE - GLPI
'                   Se cre� modelo 293 (copia del 354 pero el periodo  se indica con MMYYYY en lugar de pliqnro)
'                   Modelo 233: se agreg� mejora de validacion de observaciones y Fecha de certificado



'Global Const Version = "5.28"
'Global Const FechaModificacion = "11/08/2014"
'Global Const UltimaModificacion = "" ' LED - CAS-25180 - Uni�n Personal - Desarrollo de nueva interfaz para EyP
'                   Modelo 392: Nuevo interfaz de importacion de requerimiento de personal custom

'Global Const Version = "5.27"
'Global Const FechaModificacion = "08/08/2014"
'Global Const UltimaModificacion = "" ' Carmen Quintero -CAS-26657 - LA CAJA - Interfaz 640
'                   Modelo 640: Se agreg� validacion : si los datos a migrarse tienen el tilde de 'Fecha de Alta reconocida',
'                   y tambi�n esta tildada la 'Fecha de Alta Reconocida'  de la Fase Activa.
'                   Solo deber�a quedar tildada la Fecha de Alta Reconocida de la Fase Hist�rica.
                   
'Global Const Version = "5.26"
'Global Const FechaModificacion = "30/07/2014"   'FGZ
'Global Const UltimaModificacion = "" 'CAS-21778 - Sykes El Salvador - Interface modelo 233 y 335
'                   Modelo 335: nuevas validaciones para las novedades parciales fijas.
'                   Modelo 233: se agreg� mejora de validacion de periodo de vacaciones

'Global Const Version = "5.25"
'Global Const FechaModificacion = "30/07/2014"   'FGZ
'Global Const UltimaModificacion = "" 'CAS-21778 - Sykes El Salvador - Interface modelo 233 y 335
''                   Modelo 335: se corrigio el script de la insert. faltaba las comillas simples en las horas
''                   Modelo 233: se agreg� campo de fecha de certificado (Opcional)


'Global Const Version = "5.24"
'Global Const FechaModificacion = "17/07/2014"
'Global Const UltimaModificacion = "" ' CAS-22751 - VILLA MARIA - Adecuacion Empleos y Postulantes - Bug en la interfaz 921
'                                     ' Se modific� la consulta que valida si el nro de via existe en la tabla via.
'                                     ' Se modific� para que se guarde la descripcion de la empresa en el campo empresa de la tabla empant

'Global Const Version = "5.23"
'Global Const FechaModificacion = "16/07/2014"
'Global Const UltimaModificacion = "" ' CAS-26028 - H&A - Modificaciones R4 - Bug en interfaz 211 y 354
'                                     ' modelo 211: Se agreg� funci�n listaConceptosPermitidos()
'                                     ' modelo 354: Se agreg� funci�n listaConceptosPermitidos()

'Global Const Version = "5.22"
'Global Const FechaModificacion = "07/07/2014"
'Global Const UltimaModificacion = "" ' CAS-26326 - SYKES - ERROR EN INTERFAZ 341
                                     ' Se corrigi� el sql que verificaba la superposici�n de novedades


'Global Const Version = "5.21"
'Global Const FechaModificacion = "25/06/2014"
'Global Const UltimaModificacion = "" ' CGonzalez Nicol�s - CAS-25705 - Heidt - Nacionalizacion Mexico - Domicilio
                                     ' 1006 - Se corrigen logs
'Global Const Version = "5.20"
'Global Const FechaModificacion = "23/06/2014"
'Global Const UltimaModificacion = "" ' Carmen Quintero - CAS-22751 - VILLA MARIA - Adecuacion Empleos y Postulantes - Bug en la interfaz 921 [Entrega 2]
'                                     ' 921: Se reordenaron las columnas.


'Global Const Version = "5.19"
'Global Const FechaModificacion = "19/06/2014"
'Global Const UltimaModificacion = "" ' Gonzalez Nicol�s - CAS-23705 - SYKES - BUG EN CARGA DE INTERFAZ 341
'                                     ' 341: Se cierra vigencia anterior cuando se inserta un registro con la vigencia abierta.


'Global Const Version = "5.18"
'Global Const FechaModificacion = "19/06/2014"
'Global Const UltimaModificacion = "" ' Gonzalez Nicol�s - CAS-25705 - Heidt - Nacionalizacion Mexico - Domicilio
                                     ' 1006 - Interfaz Organizaci�n Territorial Multipa�s.
                                     ' 668 - Se crearon nuevas funciones para validaci�n de org. territorial.


'Global Const Version = "5.17"
'Global Const FechaModificacion = "12/06/2014"
'Global Const UltimaModificacion = "" ' Gonzalez Nicol�s - CAS-23705 - SYKES - BUG EN CARGA DE INTERFAZ 341
                                     ' 341: Se cierra la vigencia a un d�a anterior de la fecha desde.


'Global Const Version = "5.16"
'Global Const FechaModificacion = "11/06/2014"
'Global Const UltimaModificacion = "" ' Gonzalez Nicol�s - CAS-23705 - SYKES - BUG EN CARGA DE INTERFAZ 341
                                     ' Se agregaron control de vigencias


'Global Const Version = "5.15"
'Global Const FechaModificacion = "27/05/2014"
'Global Const UltimaModificacion = "" ' Gonzalez Nicol�s - CAS-25641 - VSO - PRAXAIR BOLIVIA - Interface Empleados
                                     ' Se agrego 2do nombre y 2do apellido en el update


'Global Const Version = "5.14"
'Global Const FechaModificacion = "26/05/2014"
'Global Const UltimaModificacion = "" ' Fernandez, Matias -CAS-25552 - MEDICUS - ERROR EN CARGA DE INTERFAZ DE VALES
                                     ' Interface 217, se detallo mas el log para poder identificar un problema de duplicacion identity



'Global Const Version = "5.13"
'Global Const FechaModificacion = "26/05/2014"
'Global Const UltimaModificacion = "" ' Gonzalez Nicol�s - CAS-25641 - VSO - PRAXAIR BOLIVIA - Interface Empleados
                                     ' Se cre� modelo 616 Importaci�n de Empleados Bolivia. (Copia 605)

'Global Const Version = "5.12"
'Global Const FechaModificacion = "13/05/2014"
'Global Const UltimaModificacion = "" 'FGZ - CAS-22751 - VILLA MARIA - Adecuacion Empleos y Postulantes - Bug en la interfaz 921
'                                    'Se modific� el modelo 668 (Importaci�n de Domicilios Multi-Pa�s) - Se corrigi� la insercion de Telefonos.


'Global Const Version = "5.11"
'Global Const FechaModificacion = "13/05/2014"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-22751 - VILLA MARIA - Adecuacion Empleos y Postulantes - Bug en la interfaz 921
''                                    'Se modific� el modelo 921 (Importacion de Postulantes de Per�)


'Global Const Version = "5.10"
'Global Const FechaModificacion = "29/04/2014"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-16441 - H&A - PERU - Cuentas de Acreditacion Bancaria [Entrega 3] (CAS-15298)
'                                    'Se mejora el modelo 271 de cuentas bancarias para peru

'Global Const Version = "5.09"
'Global Const FechaModificacion = "09/04/2014"
'Global Const UltimaModificacion = "" 'RD y FGZ - CAS-23307 - RAFFO - Escala de Comisiones.
'                                       Nuevo odelo 674 - Interface de Escalas de comisiones

'Global Const Version = "5.08"
'Global Const FechaModificacion = "07/04/2014"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-22132 - SGS - Interface Estudios Formales - Se creo Modelo 291 que inserta los estudios formales del empleado.

'Global Const Version = "5.07"
'Global Const FechaModificacion = "18/03/2014"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-24464 - VILLA MARIA - Importaci�n de Domicilios Multi-Pa�s - Se agregaron tipos de tel�fonos al modelo 668

'Global Const Version = "5.06"
'Global Const FechaModificacion = "17/03/2014"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS-21296 - Megatlon - Mejora Interfaz Familiares - se acomodaron indices mal calculados en campor previos
'                                      Gonzalez Nicol�s - CAS-24204 - NACIONALIZACION BOLIVIA - Modelo de domicilio Bolivia.
'                                      Modelo 668 - Se agreg� C�digo Postal Para Bolivia.

'Global Const Version = "5.05"
'Global Const FechaModificacion = "06/03/2014"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-24204 - NACIONALIZACION BOLIVIA - Modelo de domicilio Bolivia.
                                     ' Se agreg� modelo para Bolivia


'Global Const Version = "5.04"
'Global Const FechaModificacion = "21/02/2014"
'Global Const UltimaModificacion = "" '20/02/2014 - Fernandez, Matias - CAS-21778 - Sykes El Salvador- QA - Interface 230 - aborta al haber superposicion.


'Global Const Version = "5.03"
'Global Const FechaModificacion = "20/02/2014"
'Global Const UltimaModificacion = "" '20/02/2014 - Fernandez, Matias - CAS-21778 - Sykes El Salvador- QA - Interface 230 - Se chequea que no haya superposicion de dias

'Global Const Version = "5.02"
'Global Const FechaModificacion = "10/02/2014"
'Global Const UltimaModificacion = "" '10/02/2014 - Carmen Quintero - (CAS-23705 - SYKES - BUG EN CARGA DE INTERFAZ 341) Se modifico para que tenga cierre de novedades

'Global Const Version = "5.01"
'Global Const FechaModificacion = "06/02/2014"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS-21296 - Megatlon - Mejora Interfaz Familiares - comillas en campos varchar
                                     
'Global Const Version = "5.00"
'Global Const FechaModificacion = "05/02/2014"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-19564 - Raffo - Errores interfaz 275 - Se realizaron modificaciones varias al modelo 286
                                     ' para que quede como estandar para R3


'Global Const Version = "4.99"
'Global Const FechaModificacion = "29/01/2014"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS-23699 - Sykes El Salvador - Interface Modelo 229 - Se agregaron las columnas quincenal y quincena para prestamos en modelo 229


'Global Const Version = "4.98"
'Global Const FechaModificacion = "27/01/2014"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-17135 - HEIDT - Mejoras en origen de curso de GDD - Se modific� el modelo 290 por agregarse validaci�n en la �ltima
                                     'columna del archivo si se incluye el separador o no.
                                     
'Global Const Version = "4.97"
'Global Const FechaModificacion = "08/01/2014"
'Global Const UltimaModificacion = "" 'Dimatz Rafael - CAS 22751 - Se agrego al Domicilio el campo Distrito


'07/01/2014 - Gonzalez Nicol�s - Modelos 265 y 387  - Se corrige error al validar si ya existe un documento para otro empleado.


'Global Const Version = "4.96"
'Global Const FechaModificacion = "07/01/2014"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-17135 - HEIDT - Mejoras en origen de curso de GDD - Se creo Modelo 290 que inserta el diccionario de competencias basico
'                                      '02/01/2014 - Gonzalez Nicol�s - Modelo 265 - Se agrego variable insertaFecha para corregir error al insertar/update la fecha cuando venia con N/A


'Global Const Version = "4.95"
'Global Const FechaModificacion = "06/01/2014"
'Global Const UltimaModificacion = "" 'Dimatz Rafael - CAS 22751 - Se modificaron los indices de los arreglos para que coincidan con los datos de la columna del CSV
                                     

'Global Const Version = "4.94"
'Global Const FechaModificacion = "27/12/2013"
'Global Const UltimaModificacion = "" 'Dimatz Rafael - CAS 22751 - Se modificaron los indices de los arreglos para que coincidan con los datos de la columna del CSV

'Global Const Version = "4.93"
'Global Const FechaModificacion = "18/12/2013"
'Global Const UltimaModificacion = "" 'LED - CAS-22808 - SGS - Distribuci�n Contable
    'se agrego nuevo modelo 288, igual al 354 exportacion de novedades pero se le agrego distribucion contable
    'se agrego nuevo modelo 289, igual al 245 exportacion de novedades ajuste pero se le agrego distribucion contable

'Global Const Version = "4.92"
'Global Const FechaModificacion = "11/12/2013"
'Global Const UltimaModificacion = "" 'Fernandez, Matias - CAS- 22211 - flvl7s - Heidt - interfaces - se arreglo los calculos de los indices en el modelo 253


'Global Const Version = "4.91"
'Global Const FechaModificacion = "11/12/2013"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - Se creo modelo 287 que levanta tarjetas masivamente para Spec - CAS-22047 - VSO - AGCO - Interface SPEC x Web Service


'Global Const Version = "4.90"
'Global Const FechaModificacion = "10/12/2013"
'Global Const UltimaModificacion = "" 'Dimatz Rafael - Se creo modelo 921 Importacion de Postulantes Peru - CAS-22751 - VILLA MARIA - Adecuacion Empleos y Postulantes

'Global Const Version = "4.89" Esta Version nunca se Entrego
'Global Const FechaModificacion = "04/12/2013"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - Se agrego modelo 287 que levanta tarjetas masivamente para Spec - CAS-22047 - VSO - AGCO - Interface SPEC x Web Service


'Global Const Version = "4.88"
'Global Const FechaModificacion = "27/11/2013"
'Global Const UltimaModificacion = "" 'FGZ - Se cambiaron las validaciones de cantidad de caracteres parta los nros de documento en los
'                                       modelos 265, 387, 664 y 672
'                                     CAS-21288 - Sykes El Salvador - ADP  Nro Doc. Interface 265


'Global Const Version = "4.87"
'Global Const FechaModificacion = "18/11/2013"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - Se agrega el campo comentario al modelo 335
'                                     CAS-21493 - SYKES CR - Custom nueva columna para anormalidades


'Global Const Version = "4.86"
'Global Const FechaModificacion = "11/11/2013"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - Se valida que tenga saldo cuando es una licencia de vacaciones.
'                                     Modelo 344 - CAS-21381 - Sykes - Modificacion Interfaz 344



'Global Const Version = "4.85"
'Global Const FechaModificacion = "29/10/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-21260 - Sykes El Salvador - Modelo de Domicilio El Salvador.
'                                     Modelo 668 - Modificaci�n en domicilio SL

'Global Const Version = "4.84"
'Global Const FechaModificacion = "03/10/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-20391 - H&A - Multipais - Modelo 668 - Importaci�n Domicilios Multi-Pais.
'                                     Modelo 668 - Se agreg� domicilios de CR y SL


'Global Const Version = "4.83"
'Global Const FechaModificacion = "03/10/2013"
'Global Const UltimaModificacion = "" 'Fernandez Matias- CAS-21296 - Megatlon - Mejora Interfaz Familiares - se corrigio el update de la tabla familiar
 
 
 
'Global Const Version = "4.82"
'Global Const FechaModificacion = "02/10/2013"
'Global Const UltimaModificacion = "" 'Fernandez Matias- CAS-21296 - Megatlon - Mejora Interfaz Familiares - Se agrego: fecha documentacion,acta, tomo, folio
                                      ' tribunal, juzgado, secretaria,comuna
'Global Const Version = "4.81"
'Global Const FechaModificacion = "17/09/2013"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-20979 - VSO - POLLPAR - Importacion de Vales
''                                    modelo 217- se agrego como ultimo campo de la interfaz el campo descripcion, no es obligatorio


'Global Const Version = "4.80"
'Global Const FechaModificacion = "02/09/2013"
'Global Const UltimaModificacion = "" 'Mauricio Zwenger - CAS-20933 - SYKES CR - CUSTOM INTERFAZ 341 VIGENCIA
''                                      modelo 341 - se quito obligatoriedad de poner fecha hasta en las novedades que poseen vigencia.


'Global Const Version = "4.79"
'Global Const FechaModificacion = "28/08/2013"
'Global Const UltimaModificacion = "" 'Carmen Quintero -CAS-19580 - AGD - INTERFAZ DE ESCALAS (CAS-15298) [Entrega 1]
''                                      modelo 246 - Se modific� para que la Fecha Vigencia se actualice en la tabla cabgrilla

'Global Const Version = "4.78"
'Global Const FechaModificacion = "20/08/2013"
'Global Const UltimaModificacion = "" ' Gonzalez Nicol�s - CAS-19944 - SYKES CR - Error en licencias pagas
'                                 Modelos 334,336,337,338,339,341,358: Se modifico parametro que se la pasa a la fnc. getlastIdentity()

'Global Const Version = "4.77"
'Global Const FechaModificacion = "12/07/2013"
'Global Const UltimaModificacion = "" 'Carmen Quintero -CAS-19580 - AGD - INTERFAZ DE ESCALAS
''                                      modelo 246 - Se agreg� el campo de Fecha Vigencia

'Global Const Version = "4.76"
'Global Const FechaModificacion = "10/07/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-19126 -  RAPSODIA - Interfaz Domicilios multipais'
'                                      modelo 668 - Se comento EMAIL en modelo de Venezuela

'Global Const Version = "4.75"
'Global Const FechaModificacion = "05/07/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-19167 - RAPSODIA - Mejora interfaz familiares -
'                                     - modelo 602 - Se agregaron campos opcionales (TipoDoc2, documento2,guarderia, fecha guarderia)

'Global Const Version = "4.74"
'Global Const FechaModificacion = "24/06/2013"
'Global Const UltimaModificacion = "" '24/06/2013 - Mauricio Zwenger - CAS-15590 - Se modific� la secci�n de lectura de parametros
                                     'para que tenga en cuenta todos formatos posibles indicados para el modelo 211

'Global Const Version = "4.73"
'Global Const FechaModificacion = "18/06/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-19944 - SYKES CR - Error en licencias pagas
                                      'Nuevo Modelo 388 - Licencias Pagas - Sykes CR

'Global Const Version = "4.72"
'Global Const FechaModificacion = "13/06/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-17878 - SYKES CR - Reporte de pendiente de firmas.
                                      '- Se formatean a NULL hs inicio y hs hasta cuando es PV.

'Global Const Version = "4.71"
'Global Const FechaModificacion = "04/06/2013"
'Global Const UltimaModificacion = "" 'AA - CAS-13764 - Sykes - Bug en interfaz 603

'Global Const Version = "4.70"
'Global Const FechaModificacion = "30/05/2013"
'Global Const UltimaModificacion = "" 'FGZ - CAS-19167 -  RAPSODIA - Interfaz Documentos Multi-Pais
                                     'Modelo 387: Nuevo modelo de Documentos Multi-Pais.

'Global Const Version = "4.69"
'Global Const FechaModificacion = "27/05/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-19126 -  RAPSODIA - Interfaz Domicilios multipais
'                                     'Modelo 668: Se EscribeLogMiTooltIP() y campo partido para peru.

'Global Const Version = "4.68"
'Global Const FechaModificacion = "02/05/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-18330 - H&A - ECUADOR - Modelo de Domicilio
'                                     'Modelo 668: Se agreg� ECUADOR al modelo de domicilios.
                                     

'Global Const Version = "4.67"
'Global Const FechaModificacion = "25/04/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-19126 -  RAPSODIA - Interfaz Domicilios multipais
                                     'Modelo 668: Se agregaron Nuevos modelos de domicilios.
                                     ' Chile, Colombia, Brasil , Uruguay, Venezuela, M�jico y Paraguay

'Global Const Version = "4.66"
'Global Const FechaModificacion = "21/03/2013"
'Global Const UltimaModificacion = "" 'LED - CAS-15298 - Vision OutSourcers - Interfaz comisiones-(CAS-16012 )
                                     'Modelo 352-353: Correcion en la busqueda de tipo de tercero (1) por documento.


'Global Const Version = "4.65"
'Global Const FechaModificacion = "20/03/2013"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-16152 - TATA- E&P- interfaz de postulantes Uruguay
                                     'Nuevo Modelo 359: Interfaz de postulantes Uruguay
                                     
'Global Const Version = "4.64"
'Global Const FechaModificacion = "13/03/2013"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-18147 - Sykes - Replicar Interfaz 341
                                     'Nuevo Modelo 358: Interfaz de novedades Autorizadas por FF.


'Global Const Version = "4.63"
'Global Const FechaModificacion = "13/03/2013"
'Global Const UltimaModificacion = " version inexistente "


'Global Const Version = "4.62"
'Global Const FechaModificacion = "12/03/2013"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-16441 - H&A - PERU - Cuentas de Acreditacion Bancaria
                                    'Se Corregio el modelo 671, se agrego tipo de pago en el insert para que no supere el lim de porc


'Global Const Version = "4.61"
'Global Const FechaModificacion = "12/03/2013"
'Global Const UltimaModificacion = "" 'LED - CAS-17310 - CDA - Interfase Timesheet a Administrador de proyectos
                                    'Se Corregio consulta en el modelo 670, para ver si el proyecto tiene horas cargadas.

'Global Const Version = "4.60"
'Global Const FechaModificacion = "01/03/2013"
'Global Const UltimaModificacion = "" 'LED - CAS-13764 - H&A - Imputacion Contable
                                    'Se agrego la validacion si existe una hora para el mismo proyecto, d�a, tercero y fecha y acualiza o inserta.

'Global Const Version = "4.59"
'Global Const FechaModificacion = "27/02/2013"
'Global Const UltimaModificacion = "" 'EAM - CAS-17310 - CDA - Interfase Timesheet a Administrador de proyectos
                                    'Se agrego la validacion si existe una hora para el mismo proyecto, d�a, tercero y fecha y acualiza o inserta.

'Global Const Version = "4.58"
'Global Const FechaModificacion = "07/02/2013"
'Global Const UltimaModificacion = "" 'FGZ - CAS-17828 - CDA - Tipos de Documentos MultiPais
                                    'Nuevo Modelo 672: Interfaz de empleados reducida con soporte mutipais para documentos.
'                                   Es muy similar al modelo 664, pero permite levantar documentos multipasi muti bd
                                
                                     
'Global Const Version = "4.57"
'Global Const FechaModificacion = "17/01/2013"
'Global Const UltimaModificacion = "" 'Sebastian Stremel
'                                    'Nuevo Modelo 671 : CAS-16441 - H&A - PERU - Cuentas de Acreditacion Bancaria

'Global Const Version = "4.56"
'Global Const FechaModificacion = "10/01/2013"
'Global Const UltimaModificacion = "" 'Deluchi, Ezequiel
                                    'Nuevo Modelo 670 : Parte de Horas - Custom - CAS 17310
'Global Const Version = "4.55"
'Global Const FechaModificacion = "28/12/2012"
'Global Const UltimaModificacion = "" 'Margiotta, Emanuel
                                    'Nuevo Modelo 355 : Notificacion horarias
'Global Const Version = "4.54"
'Global Const FechaModificacion = "13/12/2012"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s
'                                    'Nuevo Modelo 354 : Copia 211 + Controla la vigencia de las novedades.
'Global Const Version = "4.53"
'Global Const FechaModificacion = "29/11/2012"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-17566 - PKF - Intefaz de Importacion de Novedades Micros Fidelio
                                     'Modelo 351: Se corrigio error en suma y update de novedades.

'Global Const Version = "4.52"
'Global Const FechaModificacion = "23/11/2012"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-17396 - H&A - Formato Migracion Colombia
                                     'Modelo 608: Se agrego terape2 y ternom2

'Global Const Version = "4.51"
'Global Const FechaModificacion = "22/11/2012"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-17566 - PKF - Intefaz de Importacion de Novedades Micros Fidelio
                                     'Modelo 351: Prmite actualizar / pisar / sumar novedades

'Global Const Version = "4.50"
'Global Const FechaModificacion = "21/11/2012"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-17567 - GRUPO BAPRO - ERROR EN INTERFACES DE NOVEDADES
                                     'Modelo 211: No se permite insertar una novedad con Vigencia cuando ya hay una existente con mismo CO y PAR.
'Global Const Version = "4.49"
'Global Const FechaModificacion = "16/11/2012"
'Global Const UltimaModificacion = "" 'Deluchi Ezequiel - CAS 16993 - DTT - Custom Interfaz estudios informales
                                    'Modelo 669: correccion en insert variable fecha hasta.

'Global Const Version = "4.48"
'Global Const FechaModificacion = "06/11/2012"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-16676 - Agregar campo en la interfaz - Market line-(CAS-15298)
'                                    'correccion errores modelo 233

'Global Const Version = "4.47"
'Global Const FechaModificacion = "02/11/2012"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-16676 - Agregar campo en la interfaz - Market line-(CAS-15298)
'                                    'correccion errores modelo 233

'Global Const Version = "4.46"
'Global Const FechaModificacion = "01/11/2012"
'Global Const UltimaModificacion = "" 'Deluchi Ezequiel - CAS 16993 - DTT - Custom Interfaz estudios informales
'                                    'Nuevo Modelo 669: importacion de estudios informales.

'Global Const Version = "4.45"
'Global Const FechaModificacion = "31/10/2012"
'Global Const UltimaModificacion = "" 'Deluchi Ezequiel - CAS 16012 - Vision Outsourcers - Chacomer Py - Interfaz Comisiones
'                                    'Modelo 352 y 353: cambio en la funcion cargar_confrep en el chequeo de que el parametro sea por novedad.

'Global Const Version = "4.44"
'Global Const FechaModificacion = "24/10/2012"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - CAS-17348 - NGA - AZ Colombia - Error Importacion de Prestamos
'                                    'Modelo 229: se cambio tipo de variable para el monto total

'Global Const Version = "4.43"
'Global Const FechaModificacion = "19/10/2012"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-16578 - Sykes - Error de Licencias CR
'                                    'Modelo 344: Se comento validaci�n para licencias con fecha posterior. (Sykes ya no lo utiliza)

'Global Const Version = "4.42"
'Global Const FechaModificacion = "09/10/2012"
'Global Const UltimaModificacion = "" 'Sebastian Stremel - se agrego campo observaciones al modelo 233
                                    


'Global Const Version = "4.41"
'Global Const FechaModificacion = "02/10/2012"
'Global Const UltimaModificacion = "" 'Manterola Maria Magdalena -CAS-16441 - H&A - PERU - Interfaz de Domicilios
                                    'Nuevo Modelo 668
                                     

'Global Const Version = "4.40"
'Global Const FechaModificacion = "21/09/2012"
'Global Const UltimaModificacion = "" 'Carmen Quintero -CAS-13764 � H&A � Error en Interfaz 640
'                                     Modelo 640 - Se verifica que el empleado no tenga una fase activa actual
'                                     a la fecha de la fase historica que se esta actualizando, al momento de desvincularlo.


'Global Const Version = "4.39"
'Global Const FechaModificacion = "18/09/2012"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-16578 - Sykes - Error de Licencias CR
'                                     Modelo 344 - Se calcula la cantidad de dias de la lic. 18 Seg�n l�mite por evento.


'Global Const Version = "4.38" '
'Global Const FechaModificacion = "17/09/2012"
'Global Const UltimaModificacion = "" 'Manterola Maria Magdalena - CAS-16441 - H&A - PERU - 605 multinacional reducida
                                     'Modelo 664 - Similar al Modelo 605 pero reducido. Se modific� para que en Nombre y Apellido se pueda ingresar sin @.

'Global Const Version = "4.37" '
'Global Const FechaModificacion = "10/09/2012"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-16833 - Vision Outsourcers - Chacomer Py - Interfaz Comisiones Produccion
                                     'Modelo 667 - Importaci�n de Novedades Globales con Vigencia Diaria. Estandar
'Global Const Version = "4.36" '
'Global Const FechaModificacion = "06/09/2012"
'Global Const UltimaModificacion = "" 'Manterola Maria Magdalena - CAS-16441 - H&A - PERU - 605 multinacional reducida
                                     'Modelo 664 - Similar al Modelo 605 pero reducido.

'Global Const Version = "4.35" '
'Global Const FechaModificacion = "24/08/2012"
'Global Const UltimaModificacion = "" 'Dimatz Rafael CAS-16341 - TATA
                                     'Modelo 275 - Se agrego en la tabla empant que guarde la empresa

'Global Const Version = "4.34" '
'Global Const FechaModificacion = "23/08/2012"
'Global Const UltimaModificacion = "" 'Brzozowski Juan Pablo CAS-16396 - NORTHGATE ARINSO
'                                     'Modelo 605 - Cuando se daba el caso en el que el doc estaba asociado a otro empleado,
'                                     '             no cortaba el proceso. Ahora corta con mensaje de error.

'Global Const Version = "4.33" '
'Global Const FechaModificacion = "17/08/2012"
'Global Const UltimaModificacion = "" 'Margiotta, Emanuel CAS-15533 - PKF - Interfaz de Importaci�n de Novedades Micros Fidelio
                                     'Modelo 351 - Se corrigio la funcion de resta. No estaba tomando las horas.
                                      
'Global Const Version = "4.32" '
'Global Const FechaModificacion = "06/08/2012"
'Global Const UltimaModificacion = "" 'CAS-15533 - PKF - Interfaz de Importaci�n de Novedades Micros Fidelio
'                                      'Modelo 351 - Se modifico para que sume y reste horas con el formato hhhhhmm


'Global Const Version = "4.31" '
'Global Const FechaModificacion = "31/07/2012"
'Global Const UltimaModificacion = "" 'Carmen Quintero - CAS-16536 - TATA - ADP - Interfaz de Familiares Nacionalizaci�n
'                                      'Modelo 909 - Se ajustaron las columnas apellido y nombre a este formato:
'                                      'apellido@apellido2;nombre@nombre2
'                                      'Se modific� la manera de obtener el valor del campo Fecha_IRPF.

'Global Const Version = "4.30" '
'Global Const FechaModificacion = "31/07/2012"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-16555 - Sykes - Error Interfaces -
''                                     Modelo 344 - Insertaba mal una licencia cuando era menor de 3 d�as.


'Global Const Version = "4.29" '
'Global Const FechaModificacion = "16/07/2012"
'Global Const UltimaModificacion = "" 'Brzozowski Juan Pablo: CAS-16396 - NORTHGATE ARINSO (OLX) - Error de version interfaces 4.27 4.28
                                     'Bug: Insertaba dos veces en ter_tip.  LineaModelo_605

'Global Const Version = "4.28" '
'Global Const FechaModificacion = "28/06/2012"
'Global Const UltimaModificacion = "" 'Gonzalez Nicol�s - CAS-16040 - Sykes - Interfaz de licencias
'          Modelo Custom 344: Si no se informa el usuario destino de firma, busca al Reporta A



'Global Const Version = "4.27" 'LED
'Global Const FechaModificacion = "26/06/2012"
'Global Const UltimaModificacion = "" 'Margiotta, Emanuel (CAS-15553)
'          Modelo Custom 351: Se corrgio para que en la columna de Hs Normales le reste las horas al 50% y las hs 100%
'


'Global Const Version = "4.26" 'LED
'Global Const FechaModificacion = "25/06/2012"
'Global Const UltimaModificacion = "" 'Nuevos Modelos (CAS-16012)
''                           modelo 352: Comisiones  - Monto
''                           modelo 353: Comisiones  - Cantidad


'Global Const Version = "4.25" ' Brzozowski Juan Pablo
'Global Const FechaModificacion = "18/06/2012"
'Global Const UltimaModificacion = "" 'Se realizo modificaciones en los modelos 603 y 605 en el proceso de la Interfaz General.
''                Modelo 603: Se modific� para que, si el documento no es un empleado y pertenece a un:
''                   Postulante: Modifico el estado del postulante a inactivo e inserto el tertip = 1 para el postulante.
''                   Familiar: Inserto un nuevo tertip = 1 para el familiar
''                Modelo Estandar 605: Si el documento pertenece a otro tercero que no sea un empleado, entonces informa en el log y no carga el empleado.



'Global Const Version = "4.24" ' Brzozowski Juan Pablo
'Global Const FechaModificacion = "14/06/2012"
'Global Const UltimaModificacion = "" 'Se controla que el Nro_Provincia y Nro_Pais no sea 0 en las funciones ValidarZona y ValidarProvincia


'Global Const Version = "4.23" ' Brzozowski Juan Pablo
'Global Const FechaModificacion = "06/06/2012" 'Se adapto para el estandar R3 el caso CAS-15661 (realizado en version 4.22).
'Global Const UltimaModificacion = "" 'Modelos:
'               Modelo 604: Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.
'               Modelo 605: Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.
'               Modelo 606: Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.
'               Modelo 607: Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.
'               Modelo 608: Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.
'               Modelo 609: Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.
'               Modelo 611: Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.
'               Modelo 612: Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.
'               Modelo 613: Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.


'Global Const Version = "4.22" ' Brzozowski Juan Pablo
'Global Const FechaModificacion = "04/06/2012" 'Modelo 603 - CAS-15661 - Sykes - Automatizacion de Ingresos-Egresos
'Global Const UltimaModificacion = "" 'Al incorporar el postulante como empleado, se habilitar� al mismo para el acceder al ESS para acceder con un perfil y password determinados por defecto  por la empresa.


'Global Const Version = "4.21" ' Gonzalez Nicol�s
'Global Const FechaModificacion = "24/05/2012" 'Modelo 344 - CAS-15972 - Sykes - Error Tope D�as de Licencias
'Global Const UltimaModificacion = "" 'Se toman en cuenta las licencias 46 y 48 para partici�n de lic. 8

'Global Const Version = "4.20" ' Zamarbide Juan - Modelo 605 - CAS-15767 - SOS - QA - BUG INTERFACE 605
'Global Const FechaModificacion = "15/05/2012" 'Modelo 605 - Se modific� la funci�n AsignarEstructura_SitRev2 por que no contemplaba el caso que est�n dando de alta un empleado
'Global Const UltimaModificacion = "" 'Adem�s se modific� la llamada por que s�lo realizaba los cambios a los empleados activos.

'Global Const Version = "4.19" ' Zamarbide Juan - Modelo 351 - CAS-15533 - PKF - Interfaz de Importaci�n de Novedades Micros Fidelio
'Global Const FechaModificacion = "04/05/2012" 'Modelo 351 - Se concaten� un 80 delante del N� de Legajo, para que coincida con el de RHPro en PKF
'Global Const UltimaModificacion = ""

'Global Const Version = "4.18" ' Sebastian Stremel modelo 341 CAS-14851 - Sykes - Modificaciones Varias
'Global Const FechaModificacion = "03/05/2012" 'Modelo 341
'Global Const UltimaModificacion = ""

'Global Const Version = "4.17" ' Dimatz Rafael - CAS-13764 - H&A - Modelo 606 - Migracion Empleados Uruguay
'Global Const FechaModificacion = "20/04/2012" 'Modelo 606 - Se puso como no obligatorio el CUIL
'Global Const UltimaModificacion = ""

'Global Const Version = "4.16" ' Zamarbide Juan - CAS-15298- Plan Obra Social Ley-Price-(CAS-13448) - Rechazo
'Global Const FechaModificacion = "20/04/2012" 'Modelo 605 - Se corrigieron los errores descritos en el formulario de rechazo
'Global Const UltimaModificacion = "" ' OS Ley y replica en OS Elegida (y viceversa) - Plan OS Ley y replica en Plan OS Elegida (y viceversa)

'Global Const Version = "4.15" ' Zamarbide Juan - CAS-15590 - NORTHGATEARINSO - Error en campos interface novedades
'Global Const FechaModificacion = "18/04/2012" 'Modelo 211 y 245 - Se Se cambi� el tipo de dato de la variable Monto de Single a Double
'Global Const UltimaModificacion = ""

'Global Const Version = "4.14" ' Zamarbide Juan - CAS-15298- Plan Obra Social Ley-Price-(CAS-13448)
'Global Const FechaModificacion = "16/04/2012" 'Modelo 605 - Se corrigieron los errores descritos en el formulario de rechazo
'Global Const UltimaModificacion = "" ' OS Ley y replica en OS Elegida (y viceversa) - Plan OS Ley y replica en Plan OS Elegida (y viceversa)

'Global Const Version = "4.13" ' Zamarbide Juan - CAS-15533 - PKF - Interfaz de Importaci�n de Novedades Micros Fidelio
'Global Const FechaModificacion = "13/04/2012" 'Modelo 351 - Se creo el modelo 351 como Custom para PKF, el mismo levanta desde un archivo de texto
'Global Const UltimaModificacion = "" ' las Novedades de Hs seg{un configuraci{on de las Columnas del Confrep N�368 - Al mismo se configuran 6 Columnas
                                     ' con nro de columna, la descripci�n, c�digo de tipo de par�metro y c�digo de concepto asociado a la novedad.
                                     ' Ejemplo: 1 Hs Normales  3  01270
'Global Const Version = "4.12"  'Sebastian Stremel
'Global Const FechaModificacion = "11/04/2012" 'Modelo 341: Se controla que si el empleado tiene una novedad cargada con vigencia,
'no se pueda cargar una sin vigencia, esto se controla porque en sykes CR hubo un problema con la carga de novedades mediante interfaz.
'CAS-14851 - Sykes - Modificaciones Varias
'Global Const UltimaModificacion = ""


'Global Const Version = "4.11"  'Gonzalez Nicol�s - CAS-15529 - Sykes - Error importaci�n novedades
'Global Const FechaModificacion = "09/04/2012" ' Se valida que usuario origen sea <> a destino (Firmas)
'Global Const UltimaModificacion = "" ' Modelo 341



'Global Const Version = "4.10"  'Zamarbide Juan Alberto - CAS-13764 - H&A - Modelo 606 - Migracion Empleados Uruguay - RECHAZO -
'Global Const FechaModificacion = "03/04/2012" ' Correcci�n de error - faltaban las comas de separaci�n de campos en el UPDATE de tercero
'Global Const UltimaModificacion = "" ' Modelo 606

'Global Const Version = "4.09"  'Zamarbide Juan Alberto - CAS-13764 - H&A - Modelo 606 - Migracion Empleados Uruguay
'Global Const FechaModificacion = "30/03/2012" 'Modelo 606: Se agrego la posibilidad de ingresar el segundo nombre y el segundo apellido mediante el separador @
'Global Const UltimaModificacion = ""   'Ejemplo:  Apellido RODRIGUEZ@LARRETA
                                       '          Nombre: JOSE@ALBERTO
                                       


'Global Const Version = "4.08"  'Dimatz Rafael
'Global Const FechaModificacion = "09/03/2012" 'Modelo 640: Se cambio en recorset StrSql cuando hay Situacion de Revista
'Global Const UltimaModificacion = ""

'Global Const Version = "4.07"  'Gonzalez Nicol�s
'Global Const FechaModificacion = "06/03/2012" 'Modelo 344: Se valida que no haya cargada una licencia de tipo 8,9 y 45 posterior a la fecha de la nueva licencia
''                                                        : Se agreg� validaci�n por confrep de las licencias permitidas para importaci�n
'Global Const UltimaModificacion = ""

'Global Const Version = "4.06"  'Gonzalez Nicol�s
'Global Const FechaModificacion = "16/02/2012" 'Se valida que este activo MI y se creo funcion EscribeLogMI() que traduce etiquetas del log.
'                                              Nuevos m�dulos: MdlMidioma y MdlInterfacesPortugal
'                                              Modelo 612: Nueva importaci�n de empleados - Portugal
'                                              Modelo 912: Nueva importaci�n de familiares - Portugal
'                                              Modelo 344: Se corrigi� error en UPDATE (Faltaba una coma)- Sykes
'                                              Modelo 336: Se agreg� BK de firmas y se valida el detalle del parte antes de insertar - Sykes
'                                              Se modific� el orden de las versiones (La �ltima esta primero)
'                                              Se modifico los valores segun base 0 en el modelo 231, de las formas de pago.
'                                              Modelo 341: Se modifico insercion del circuito de firmas - Sykes
'Global Const UltimaModificacion = "Se Incluye MI"

'Global Const Version = "4.05"  'Gonzalez Nicol�s
'Global Const FechaModificacion = "23/01/2012" 'Modelo 344: 'Se seteo elorden = 1 para las licencias de tipo 18 y 19 afecta a tablas emp_lic, y  gti_justificacion
'Global Const UltimaModificacion = " "

'Global Const Version = "4.04"  'FGZ
'Global Const FechaModificacion = "19/01/2012" 'Modelo 343: 'Se le agreg� depuracioon por vigencia
'Global Const UltimaModificacion = " "

'Global Const Version = "4.03"  'Manterola Maria Magdalena
'Global Const FechaModificacion = "12/01/2012" 'Modelo 603: 'Se le agreg� un cartel al log con el nombre y tipo de la estructura que se est� intentando crear.
'Global Const UltimaModificacion = " "

'Global Const Version = "4.02"  'FGZ
'Global Const FechaModificacion = "10/01/2012" 'Modelo 343: 'Se le agreg� depuracioon por vigencia
'Global Const UltimaModificacion = " "


'Global Const Version = "4.01"  'Manterola Maria Magdalena
'Global Const FechaModificacion = "22/12/2011" 'Modelo 349: Se elimin� la validaci�n fija por longitud. Ahora se compara con el maximo obtenido de la tabla gti_tiptar

'para el tipo de tarjeta importado.
'Global Const UltimaModificacion = " "

'Global Const Version = "4.00"  'Manterola Maria Magdalena
'Global Const FechaModificacion = "13/12/2011" 'Modelo 333: Se modific� la interfaz 333 para que actualice correctamente el progreso y que muestre un error en caso que

'el idtarjeta ingresado no sea numerico
'Global Const UltimaModificacion = " "

'Global Const Version = "3.99"  'Manterola Maria Magdalena
'Global Const FechaModificacion = "13/12/2011" 'Modelo 344: Se agreg� a la interfaz de licencias para Sykes un campo mas: dias habiles
'Global Const UltimaModificacion = " "

'Global Const Version = "3.98"  'Manterola Maria Magdalena
'Global Const FechaModificacion = "12/12/2011" 'Modelo 603: Se agreg� validaci�n por confrep (reporte 363) de habilitaci�n o no a modificar o crear estructuras o tipos

'de estructuras para Sykes.
'Global Const UltimaModificacion = " "

'Global Const Version = "3.97"  'Manterola Maria Magdalena
'Global Const FechaModificacion = "01/12/2011" 'Modelo 349: Se agreg� el nuevo modelo idem al 312 para Sykes pero sin la validaci�n de tipos de estructuras

'superpuestas, porque en Sykes un empleado puede tener varias tarjetas del mismo tipo al mismo tiempo.
                
                   
'Global Const UltimaModificacion = " "

'Global Const Version = "3.96"  'Zamarbide Juan - Cas-13448-Price-Plan de obra social por ley
'Global Const FechaModificacion = "23/11/2011" 'Modelo 605: Se agreg� la funci�n ValidaEstructura2 y adicion� una entrada m�s a la funci�n CreaComplemento por lo que

'se replica la obra social y el plan de os en la estructura correspondiente si no existe, tal cual lo hace el asp cuando se agrega una nueva estructura
                   
'Global Const UltimaModificacion = " "

'Global Const Version = "3.95"  'Sebastian Stremel
'Global Const FechaModificacion = "22/11/2011" 'Modelo 318: la validacion del tipo de estructura Plan contra la base era por desc de la tabla osocial ahora se cambio y

'se hace con replica_estr
                   
'Global Const UltimaModificacion = " "

'Global Const Version = "3.94"  'Gonzalez Nicol�s
'Global Const FechaModificacion = "16/11/2011" 'Modelo 312: Se comento validaci�n de legajo.
                                              'Modelo 341: se agreg� columna usuario y se modifico la forma en que valida el usuario FF.
'Global Const UltimaModificacion = " "

'Global Const Version = "3.93"  'Manterola Maria Magdalena
'Global Const FechaModificacion = "11/11/2011" 'Modelos 346 y 347: Se crearon dos nuevos modelos, los cuales son iguales al 211 y al 245, con la diferencia que son

'para Simulacion. (INTERFACE DE NOVEDADES)
'Global Const UltimaModificacion = " "

'Global Const Version = "3.92"  'Gonzalez Nicol�s
'Global Const FechaModificacion = "11/11/2011" 'Modelo 341: Si se modifica Novedad, elimina firma y crea nueva. - Sykes
                                              'Modelo 603: Se agreg� apellido2 y nombre2 - Sykes
                                              'Modelo 343: Interface de registraciones manuales. Sykes.
                                              'Modelos 211, 230, 233, 245: Se comento llamada a v_empleadosproc
'Global Const UltimaModificacion = " "


'Global Const Version = "3.91"  'Manterola Maria Magdalena
'Global Const FechaModificacion = "03/11/2011" 'Modelo 344: Se corrigi� error cuando se insertaba en gti_justificaci�on para la lic. tipo 8
'Global Const UltimaModificacion = " "

'Global Const Version = "3.90"  'Manterola Maria Magdalena
'Global Const FechaModificacion = "25/10/2011" 'Modelo 286: Se creo dicho modelo, el cual tiene el formato de la interfaz 275 est�ndar,
'                                              'mas 10 columnas para importar hasta 5 valores para tipos e items de informaci�n general.
'Global Const UltimaModificacion = " "

'Global Const Version = "3.89"  'Zamarbide Juan Alberto
'Global Const FechaModificacion = "24/10/2011" 'Modelo 613: Se creo dicho modelo para Migraci�n de Chile, versi�n custom INDAP, el cual es una copia del 607 a la fecha
'Global Const UltimaModificacion = " "         'Modelo 607: Se volvieron atr�s los cambios de los campos Bienios(custom INDAP), dado que no funcionaban para Deloite y

'dem�s empresas. Ahora la 607 con bienios es la 613

'Global Const Version = "3.88  'Gonzalez Nicol�s"
'Global Const FechaModificacion = "18/10/2011" 'Modelo 335: Se corrigieron inserts en las tablas gti_justificacion y gti_novedad - Sykes
'Global Const UltimaModificacion = " "

'Global Const Version = "3.87  'Gonzalez Nicol�s"
'Global Const FechaModificacion = "17/10/2011" 'Modelo 344: Se agreg� partici�n de licencia tipo 8 - Sykes
                                              'Modelo 335: Se corrigi� validaci�n de fechas y sql que validaba superposici�n de novedades
                                              'Modelo 603: Se ampli� el campo Calle a 250
'Global Const UltimaModificacion = " "

'Global Const Version = "3.87  'Deluchi Ezequiel"
'Global Const FechaModificacion = "04/10/2011" ' se agrego el modelo 345, Interface Historico de vacaciones Sykes
'Global Const UltimaModificacion = " "


'Global Const Version = "3.86  'Sebastian Stremel"
'Global Const FechaModificacion = "15/09/2011" ' se modifico modelo 300, 603,604,605,606,607,608 modelo de organizacion
'Global Const UltimaModificacion = " "


'Global Const Version = "3.85  'Sebastian Stremel"
'Global Const FechaModificacion = "12/09/2011" ' se corrigio modelo 280 bandas salariales
'Global Const UltimaModificacion = " "


'Global Const Version = "3.84"                 'FGZ
'Global Const FechaModificacion = "24/08/2011" ' modelo 233: Licencias
'Global Const UltimaModificacion = " "         ' modelo 344: Licencias Sykes
'                                                   a ambos modelos se le actualiz� el campo elmaxhoras para que se vean las cantidad masima de hs a justificar cuando

'la licencia es parcial variable

'Global Const Version = "3.83"                 'FGZ
'Global Const FechaModificacion = "18/08/2011" ' modelo 605: Se agreg� funci�n AsignarEstructura_SitRev2()
'Global Const UltimaModificacion = " "


'Global Const Version = "3.82"                 'Sebastian Stremel
'Global Const FechaModificacion = "17/08/2011" 'Se agreg� el tipo de telefono al modelo 602 y modelo 600 cuando se agrega un nuevo familiar
'Global Const UltimaModificacion = " "

'Global Const Version = "3.81"                 'Zamarbide Juan - CAS-12814 - Teletech - Interface 233 - Rechazo
'Global Const FechaModificacion = "11/08/2011" 'Se agreg� al comienzo de la validaci�n del per�odo de vacaciones, la verificaci�n del valor de la variable PeriodoVac
'Global Const UltimaModificacion = " "         'para que la misma permita o no realizar la misma. Ya que es una columna no obligatoria, en el caso de que exista se

'realizar�
                                              'la validaci�n correspondiente y la posterior inserci�n de datos.

'Global Const Version = "3.80"
'Global Const FechaModificacion = "20/07/2011"
'Global Const UltimaModificacion = " " ' FGZ
                                      ' nuevo modelo 911: MIgracion de Familiares de Chile.

'Global Const Version = "3.79"
'Global Const FechaModificacion = "19/07/2011"
'Global Const UltimaModificacion = " " ' FGZ
'                                      ' modelo 607: Interface de empleados de Chile.
'                                      ' Se le agregaron 2 campos opcionales (Fecha de Bienios y Cantidad de Bienios)

'Global Const Version = "3.78"
'Global Const FechaModificacion = "08/07/2011"
'Global Const UltimaModificacion = " " ' Sebastian Stremel
'                                      ' modelo 615: Se realizo cambio en cuit y razon social, si vienen datos
'                                      ' vacios, se inserta en blanco en la base de datos, antes se insertaba nulo y el
'                                      ' liquidador no lo tomaba.

'Global Const Version = "3.77"
'Global Const FechaModificacion = "04/07/2011"
'Global Const UltimaModificacion = " " ' Manterola Maria Magdalena
                                      ' modelo 630: Se Agreg� la verificaci�n para CrearComplemento (VerSiCrearComplemento)
                                      ' y CrearTercero(VerSiCrearTercero),dependiendo del tipo de estructura creado
                                      
                                      ' Modelo 611 (Empleados de Agencia): Estaba insertando mal el cadigo de tipo de tercero cuando es una agencia (es 7 y no 28)

'Global Const Version = "3.76"
'Global Const FechaModificacion = "29-06-2011"
'Global Const UltimaModificacion = " " ' Manterola Maria Magdalena
                                      ' modelo 605: Se Modific� la llamada a CrearComplemento con tipoEstr = 23 en vez que con 25


'Global Const Version = "3.75"
'Global Const FechaModificacion = "28-06-2011"
'Global Const UltimaModificacion = " " ' FGZ
                                      ' modelo 334: Se redefini� variable tienepermiso
                                      ' modelo 337: Se redefini� variable tienepermiso
                                      ' modelo 338: Se redefini� variable tienepermiso
                                      ' modelo 339: Se redefini� variable tienepermiso
                                      ' modelo 341: Se redefini� variable tienepermiso
                                      
                                      ' modelo 344: Se cambi� columna Per�odo de vacaciones x usuario firmante

'Global Const Version = "3.74"
'Global Const FechaModificacion = "22-06-2011"
'Global Const UltimaModificacion = " " ' FGZ
'                                      ' modelo 285: se le agreg� un parametro opcional para reflejar si esta saldado o no

'Global Const Version = "3.73"
'Global Const FechaModificacion = "17-06-2011"
'Global Const UltimaModificacion = " " ' FGZ
'                                      ' modelo 233: estaba insertando mal el codigo de tipo de justificacion
'                                      ' modelo 344: estaba insertando mal el codigo de tipo de justificacion

'Global Const Version = "3.72"
'Global Const FechaModificacion = "15-06-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
'                                      ' modelo 233: Se cambi� validaci�n del per�odo de vacaciones

'Global Const Version = "3.71"
'Global Const FechaModificacion = "14-06-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
                                      ' modelo 344: Se corrigi� error (Relizaba una consulta y validaba con otro nombre de objeto)

'Global Const Version = "3.70"
'Global Const FechaModificacion = "10-06-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
                                      ' modelo 334: Se agreg� funci�n listaConceptosPermitidos()
                                      ' modelo 337: Se agreg� funci�n listaConceptosPermitidos()
                                      ' modelo 338: Se agreg� funci�n listaConceptosPermitidos()
                                      ' modelo 339: Se agreg� funci�n listaConceptosPermitidos()
                                      ' modelo 341: Se agreg� funci�n listaConceptosPermitidos()
                                      ' modelo 344: Se cambi� columna Per�odo de vacaciones x usuario firmante

'Global Const Version = "3.69"
'Global Const FechaModificacion = "03-06-2011"
'Global Const UltimaModificacion = " " ' FGZ
                                      ' modelo 334: Importacion de cafeteria - Se cambi� la validacion del nro de legajo

'Global Const Version = "3.68"
'Global Const FechaModificacion = "02-06-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
'                                      ' modelo 344: Se agreg� nuevo modelo = 233 + CIRCUITO DE FIRMAS
'                                      ' Se creo nueva funci�n firmas_lista()
'                                      ' Se modifico funci�n firmas_Nov()

'Global Const Version = "3.67"
'Global Const FechaModificacion = "01-06-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
                                      ' modelo 605: Se agreg� funci�n AsignarEstructura_SitRev()

'Global Const Version = "3.66"
'Global Const FechaModificacion = "27-05-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
                                      ' modelo 343: Nuevo Modelo - Interface de registraciones manuales.

'Global Const Version = "3.65"
'Global Const FechaModificacion = "26-05-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
                                       ' modelo 334: Se agreg� variable firmas
                                       ' Modelo 337: Se agreg� variable firmas
                                       ' modelo 338: Se agreg� variable firmas
                                       ' modelo 339: Se agreg� variable firmas
                                       ' modelo 341: Se agreg� variable firmas
                                       ' modelo 233: Se agreg� columna Periodo de Vacaciones + inserta complemento

'Global Const Version = "3.64"
'Global Const FechaModificacion = "17-05-2011"
'Global Const UltimaModificacion = " " ' FGZ
                                      ' modelo 640: Se modific� el modelo para que inserte registros historicos aun cuando ya existan fases posteriores.

'Global Const Version = "3.63"
'Global Const FechaModificacion = "12-05-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
'                                      ' modelo 334: Se corrigi� error de formato en var firmas_ok
'                                      ' Modelo 337: Se corrigi� n� de modelo | Se corrigi� error de formato en var firmas_ok
'                                      ' modelo 338: Se corrigi� error de formato en var firmas_ok
'                                      ' modelo 339: Se corrigi� error de formato en var firmas_ok
'                                      ' modelo 341: Se corrigi� error de formato en var firmas_ok


'Global Const Version = "3.62"
'Global Const FechaModificacion = "04-05-2011"
'Global Const UltimaModificacion = " " ' FGZ  -
'                            Se modific� insert en la tabla telefonos (faltaba el tipotel y luego no los mostraba)
'                            Modelos afectados
                                      ' Modelo 226 (Postulantes DTT I)
                                      ' Modelo 239 (Postulantes DTT II)
                                      ' Modelo 241 (Postulantes Dabra)
                                      ' Modelo 263 (Postulantes Carsa)
                                      ' Modelo 275 (Postulantes Estandar)
                                      ' Modelo 300 (Empleados Teleperformance)
                                      ' Modelo 328 (Postulantes Medicus)
                                      ' Modelo 600 (Familiares)
                                      ' Modelo 602 (Familiares)
                                      ' Modelo 603 (Empleados CR)
                                      ' Modelo 604 (Empleados Wallmart)
                                      ' Modelo 605 (Empleados)
                                      ' Modelo 606 (Empleados Uruguay)
                                      ' Modelo 607 (Empleados Chile)
                                      ' Modelo 608 (Empleados Colombia)
                                      ' Modelo 611 (Empleados de Agencia)
                                      ' Modelo 663 (Prestadores Medicos)
                                      ' Modelo 909 (Familiares Uruguay)
                                      ' Modelo 910 (Familiares Colombia)

'Global Const Version = "3.61"
'Global Const FechaModificacion = "26-04-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
'                                      ' Modelo 28:Se modific� insert en la tabla emp_fr_comp

'Global Const Version = "3.60"
'Global Const FechaModificacion = "20-04-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
                                      ' modelo 339: Nuevo Modelo - Interface Banco Popular (Novedades de Liq) - Sykes CR
                                                   'Se cre� funcion Firmas_novliq() - Sykes - CR
                                      ' modelo 338: Se agreg� funcion Firmas_novliq()
                                      ' modelo 337: Se agreg� funcion Firmas_novliq()
                                      ' modelo 334: Se agreg� funcion Firmas_novliq()
                                      ' modelo 341: Nuevo Modelo - Replica Linea_modelo_211 Adecuado para Sykes CR
                                                    'Se cre� funci�n listaConceptosPermitidos() - Sykes - CR
                                      
                                      'Se elimin� c�digo en el MAIN asociado al modelo 333
                                        'EAM- Arma la Sql con la vista de Empleados para el usuario del proceso
                                        'StrSql = "SELECT object_definition (OBJECT_ID(N'dbo.v_empleado'))"
                                        'OpenRecordset StrSql, rs_batch_proceso
                                        'Sql_VistaEmpleado = Mid(rs_batch_proceso(0), InStr(1, rs_batch_proceso(0), "FROM"), Len(rs_batch_proceso(0)))
                                        'Sql_VistaEmpleado = Replace(Sql_VistaEmpleado, "SUSER_SNAME()", usuario, 1)
                                        'Sql_VistaEmpleado = "SELECT empleado.ternro " & Sql_VistaEmpleado

'Global Const Version = "3.59"
'Global Const FechaModificacion = "14-04-2011"
'Global Const UltimaModificacion = " " ' Leticia A.
                                      ' modelo 611: Nuevo Modelo - Interface Empleado de Agencia (en base al modelo 605)

'Global Const Version = "3.58"
'Global Const FechaModificacion = "13-04-2011"
'Global Const UltimaModificacion = " " ' FGZ
                                      ' modelo 333: ahora lee de una tabla wc_mov_horario

'Global Const Version = "3.57"
'Global Const FechaModificacion = "07-04-2011"
'Global Const UltimaModificacion = " " ' FGZ
'                                      ' - el modelo 333 ya no se usa, se coment� todo el procedimiento

'Global Const Version = "3.56"
'Global Const FechaModificacion = "06-04-2011"
'Global Const UltimaModificacion = " " ' Manterola Maria Magdalena
'                                      ' - Modificaci�n del n�mero de Modelo 329 por el n�mero 337
'                                      ' - Modificaci�n del n�mero de Modelo 331 por el n�mero 338

'Global Const Version = "3.55"
'Global Const FechaModificacion = "16-03-2011"
'Global Const UltimaModificacion = " " ' Leticia Amadio - Modelo 630 - se agrego conteo de errores.
                                      '                - Modelo 600- 265 - se completaron conteo de errores.
                                      '                - Modelo 640 - Si hay causa de baja chequear que ssi exista F Baja - agregar situacion de revista asociada a la

'causa de baja -  completar conteo de registros de errores

'Global Const Version = "3.54"
'Global Const FechaModificacion = "15-03-2011"
'Global Const UltimaModificacion = " " ' Mart�nez Nicol�s
                                      '- Migracion Inicial 603 Costa Rica - Se fijaron los tipos de telefonos.
                                      '- Telefono Personal  ->  tipo 1
                                      '- Telefono Laboral   ->  tipo 2
                                      '- Telefono Celular   ->  tipo 3

'Global Const Version = "3.53"
'Global Const FechaModificacion = "11-03-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
                                      '- Mod. Modelo 335 - Se agreg� nuevos N� de errores InsertaError()
                                      '- Mod. Modelo 336 - Se agreg� nuevos N� de errores InsertaError()

'Global Const Version = "3.52"
'Global Const FechaModificacion = "10-03-2011"
'Global Const UltimaModificacion = " " ' Mart�nez Nicol�s
                                      '- Migracion Inicial 603 Costa Rica - Se arreglaron problemas con telefonos: les faltaba el tipo de telefono.
                                      ' Segundo nombre y segundo apellido se insertan con string vacio en vez de en nullo.

'Global Const Version = "3.51"
'Global Const FechaModificacion = "09-03-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
                                      '- Mod. modelo 334 - Se modific� el n�mero del reporte por el n�mero 310
                                      '- Mod. Modelo 335 - Se agreg� validaci�n para Fin de firma
                                      '- Mod. Modelo 336 - Se agreg� validaci�n para Fin de firma

'Global Const Version = "3.50"
'Global Const FechaModificacion = "04-03-2011"
'Global Const UltimaModificacion = " " ' Manterola Mar�a Magdalena
                                      '- Mod. modelo 331 - Se modific� el n�mero del reporte por el n�mero 313

'Global Const Version = "3.49"
'Global Const FechaModificacion = "02-03-2011"
'Global Const UltimaModificacion = " " ' FGZ - modelo 605 - Se sac� la fecha de baja en el manejo de historico de estructuras.

'Global Const Version = "3.48"
'Global Const FechaModificacion = "23-02-2011"
'Global Const UltimaModificacion = " " ' Manterola Mar�a Magdalena
                                      ' - Mod. modelo 331 - Se utiliza distinto reporte que el modelo 329

'Global Const Version = "3.47"
'Global Const FechaModificacion = "16-02-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s
                                      ' - Nuevo modelo 335 - Interface de Importaci�n de Novedades Horarias - Sykes - Costa Rica.
                                      ' - Nuevo modelo 336 - Interface de Importaci�n de Partes Diarios - Sykes - Costa Rica.

'Global Const Version = "3.46"
'Global Const FechaModificacion = "15-02-2011"
'Global Const UltimaModificacion = " " ' Margiotta Emanuel - Nuevo modelo 333 - Interface de Importaci�n de Novedades - Sykes - Costa Rica.

'Global Const Version = "3.44"
'Global Const FechaModificacion = "11-02-2011"
'Global Const UltimaModificacion = " " ' Gonzalez Nicol�s - Nuevo modelo 334 - Interface de Importaci�n de Cafeter�a - Sykes Costa rica.

'Global Const Version = "3.43"
'Global Const FechaModificacion = "11-02-2011"
'Global Const UltimaModificacion = " " ' Nicolas Martinez - Nuevo modelo 603 -Migracion del empleado - Costa rica - AsoSykes.
                                      ' Se comento la llamada al modelo 333 en Custom porque esta incompleto.
                                      ' Lisandro Moro - Mod. Modelo 282 - se agrego la condicion al where

'Global Const Version = "3.42"
'Global Const FechaModificacion = "31-01-2011"
'Global Const UltimaModificacion = " " ' Manterola Maria Magdalena - Nuevo modelo 331 - Interface de Importacion Archivos de Afiliaci�n de AsoSykes

'Global Const Version = "3.41"
'Global Const FechaModificacion = "28-01-2011"
'Global Const UltimaModificacion = " " ' Manterola Maria Magdalena - Nuevo modelo 329 - Interface de Importacion Archivos de AsoSykes Deducciones


'Global Const Version = "3.40"
'Global Const FechaModificacion = "13-01-2011"
'Global Const UltimaModificacion = " " ' Stankunas Cesar - Se modifica el modelo 325
                                      '                   13/01/2011 - Cesar Stankunas - Se cambi� el chequeo del legajo de v_empleado a empleado

'Global Const Version = "3.39"
'Global Const FechaModificacion = "29-12-2010"
'Global Const UltimaModificacion = " " ' Stankunas Cesar - Se modifica el modelo 325
                                      '                   29/12/2010 - Cesar Stankunas - Se agreg� la columna descripcion

'Global Const Version = "3.38"
'Global Const FechaModificacion = "16-12-2010"
'Global Const UltimaModificacion = " " ' Leti A. - Modelo 607 - al dar de baja --> agregar causa de baja a la Fase e Inactivarla.
'                                      '            se agrega chequeo de situacion de revista (tiene causa baja desvinculac)
'                                      '         - Modelo 605-606-607-608-609 - se arreglo consulta en cuenta bancaria -

'Global Const Version = "3.37"
'Global Const FechaModificacion = "13-12-2010."
'Global Const UltimaModificacion = " " ' Leti A. - Modelo 602: se comento MsgBox("") en chequeo de Documento (colgaba el proceso)
                                      '         - Se agregaron comentarios al log.
                                      '         - Migraci�n Inicial: se inicializa la variable HuboError a False y nroclumna a 0 en Insertar_linea_...

'Global Const Version = "3.36"
'Global Const FechaModificacion = "24-11-2010."
'Global Const UltimaModificacion = " " ' Leti A. - se eliminan caracteres invalidos de campos de texto.
                                      '         - Modelo: 605-630-640-211-233-317-602
                                      '         - La funcion EliminarCHInvalidosII() se saco del modulo varios y se incluyo aca
                                      ' FGZ - Modelo 253 - Embargos. Se corrigi� la asignacion de cantidad de cuotas cuando es por porcentaje

'Global Const Version = "3.35"
'Global Const FechaModificacion = "11-11-2010_ "
'Global Const UltimaModificacion = " " ' Leti A. - Modelo 233 - Importacion de Licencia - si se genera error se escribe en el log de errores (Lineas_Errores-xxxx )
                                       '         - Modelo 630 - (si no pisa - opcion no reemp estr.) controlar que fecha alta sea menor fecha baja - Si se tiene un

'tipo de estructura con Fecha Desde igual a Fecha Alta - reemplazar la estructura, si la Fecha de Alta es menor - dar mensaje de error
                                       '         - Modelo 605 - Al dar de Baja, si la causa de baja tiene asociado una situacion de revista --> crea la estrcutura.

'Agrega causa de baja a la Fase.
                                       ' Winsy   - Modelo 605 - control - tipo y nro documento - causa de baja

'Global Const Version = "3.34"
'Global Const FechaModificacion = "10-11-2010"
'Global Const UltimaModificacion = " " ' FGZ - Modelo 262 Dias Correspondientes de Vacaciones
'                                      Los periodos de vacaciones ahora tienen alcance por estructura (se cambio tipo de alcance 19 por 21)

'Global Const Version = "3.33"
'Global Const FechaModificacion = "05-11-2010"
'Global Const UltimaModificacion = " " ' Leticia A. - Modelo 233 - cambiar la Situaci�n de Revista asociada a la licencia
''                                                  - Si hay una instancia previa del proceso corriendo se deja el proceso actual como 'Pendiente'

'Global Const Version = "3.32"
'Global Const FechaModificacion = "15/10/2010"
'Global Const UltimaModificacion = " " ' 06/10/2010 - Dimatz Rafael - Modelo 605 - Se hace comprobacion si DNI esta dentro del CUIL
'                                                                             si la persona es extranjera que no verifique si el
'                                                                             dni esta dentro del cuil, si no tiene dni no comprueba
'                                                                             si esta dentro del cuil


'Global Const Version = "3.31"
'Global Const FechaModificacion = "06/10/2010"
'Global Const UltimaModificacion = " 06/10/2010 - Dimatz Rafael - Modelo 605 - Se hace comprobacion si DNI esta dentro del CUIL"
'                                                                             si la persona es extranjera que no verifique si el
'                                                                             dni esta dentro del cuil, si no tiene dni no comprueba
'                                                                             si esta dentro del cuil

'Global Const Version = "3.30"
'Global Const FechaModificacion = "04/10/2010"
'Global Const UltimaModificacion = " 04/10/2010 - Lisandro Moro - Modificacion modelo 909 - Se agrego el campo fecha desde irpf"

'Global Const Version = "3.29"
'Global Const FechaModificacion = "28/09/2010" ' Agregado porque faltaba la 3.14 de Cesar
'Global Const UltimaModificacion = " 08/07/2010 - Cesar Stankunas - Creaci�n del Modelo 325 Interface de Anticipos"
''                                      Creaci�n del Modelo 325 Interface de Anticipos


'Global Const Version = "3.28"
'Global Const FechaModificacion = "15/09/2010"
'Global Const UltimaModificacion = " " 'Lisandro Moro
'                                      Se creo el modelo 328 - Interface Postuolantes - Medicus
'                                      Se corrigio el modelo 275 - Interface postulantes - estandar -
'                                           Error en los indices y se actualizaron los codigos de los pasos


'Global Const Version = "3.27"
'Global Const FechaModificacion = "09/09/2010"
'Global Const UltimaModificacion = " " 'FGZ
''                                      Modificacion de Modelo 640 - Fases
''                                      Duplica las fases de un empleado cuando se informan las mismas fechas y se cambian los tildes.
''                                           Se le agreg� cierta logica para que actualice cuando puede, inserte cuando puede e informe del error cuando no.


'Global Const Version = "3.26"
'Global Const FechaModificacion = "08/09/2010"
'Global Const UltimaModificacion = " " 'FGZ
''                                      sub LeeArchivo()
''                                      Se agreg� un control cuando el archivo a levantar est� vacion

'Global Const Version = "3.25"
'Global Const FechaModificacion = "06/09/2010"
'Global Const UltimaModificacion = " " 'Margiotta Emanuel
'''                                      Modelo 230 Pedido de Vacaciones
'''                                      Se cambio el n�mero de alcance por estructura de los periodos de vacaciones a 21



'Global Const Version = "3.24"
'Global Const FechaModificacion = "27/08/2010"
'Global Const UltimaModificacion = " " 'FGZ
''                                      Modificacion de sub ValidaEstructura()
''                                           Codigos externos de la s estructuras informadas
''                                       El cambio afecta a varios modelos
''                                            604 Importaci�n Empleados Walmart
''                                            605 Importaci�n de Empleados
''                                            606 Importaci�n de Empleados Uruguay
''                                            607 Importaci�n de Empleados Chile
''                                            609 Importaci�n de Empleados Radiotronica
''                                            630 Importaci�n Hist�rico de Estructuras
''                                            261 Importaci�n Empresas
''                                            317 Importaci�n de la Distribuci�n Contable
''                                            300 Importaci�n de Empleados Teleperformance

'Global Const Version = "3.23"
'Global Const FechaModificacion = "26/08/2010"
'Global Const UltimaModificacion = " " 'FGZ
''                                      Modificacion de Modelo 605
''                                           Validacion de CUIL, ahora cuando no valida deja mensaje de error y no inserta o modifica el cuil.

'Global Const Version = "3.22"
'Global Const FechaModificacion = "03/08/2010"
'Global Const UltimaModificacion = " " 'Dimatz Rafael
''                                      Modificaciones
''                                         Modificacion de Modelo 615 Interface de DDJJ
''                                               Se corrigio que cuando se encuentra una DDJJ que ya existe
''                                               no la agregue y ponga un cartel en el log

'Global Const Version = "3.21"
'Global Const FechaModificacion = "23/07/2010"
'Global Const UltimaModificacion = " " 'FGZ
'                                      Modificaciones
'                                         605 - fases. Si el empleado estaba inactivo y se actualiza a activo ==> crea una nueva fase.
'                                         630 - Historico de Estructuras. Cuando NO pisa estructuras ==> trata de actualizar vigencia
'                                               Solo actualiza cuando la estructura actual tiene hasta nulo.

'Global Const Version = "3.20"
'Global Const FechaModificacion = "22/07/2010"
'Global Const UltimaModificacion = " " 'Dimatz Rafael
''                                      Modificacion de Modelo 602
''                                           Se agrego que escriba el error en el log

'Global Const Version = "3.19"
'Global Const FechaModificacion = "16/07/2010"
'Global Const UltimaModificacion = " " 'Dimatz Rafael
'                                      Modificacion de Modelo 602
'                                           Valido que no exista el documento
'                                           que quiero insertar de un familiar
'                                           si existe muestra error

'Global Const Version = "3.18"
'Global Const FechaModificacion = "16/07/2010"
'Global Const UltimaModificacion = " " 'Dimatz Rafael
'                                      Modificacion de Modelo 605
'                                           Se agrego validacion de CUIL
'                                           para verificar si coincide con el nro de documento
'                                           y si esta bien formado

'Global Const Version = "3.17"
'Global Const FechaModificacion = "15/07/2010"
'Global Const UltimaModificacion = " " 'Dimatz Rafael
'                                      Modificacion de Modelo 233
'                                           Se corrigio para que busque el error correcto
'                                           cuando no encuentra el estado
'                                           Ademas se agregaron errores en la tabla inerror

'Global Const Version = "3.16"
'Global Const FechaModificacion = "15/07/2010"
'Global Const UltimaModificacion = " " 'Dimatz Rafael
'                                      Modificacion de Modelo 630
'                                           Se corrigio para que guarde la fecha hasta
'                                           cuando se da un alta de Importacion
'                                           Historico de Estructuras

'Global Const Version = "3.15"
'Global Const FechaModificacion = "14/07/2010"
'Global Const UltimaModificacion = " " 'Dimatz Rafael
'                                      Modificacion de Modelo 615 Interface de DDJJ
'                                           Se verifico comprobacion de fechas cuando
'                                           se intenta ingresas desde una interface
'                                           en definitiva corrobora que no exista para ingresarla

'Global Const Version = "3.14"
'Global Const FechaModificacion = "05/07/2010"
'Global Const UltimaModificacion = " " 'FGZ
'                                      Modificacion de Modelo 282 Interface de Complemento Remunerativo
'                                           Esta actualizando el complemento para todos los empleados, es decir,
'                                           cuando levantaba una linea se lo actualizaba a todos los empleados y
'                                           no solamnete al empleado correspondiente.

'Global Const Version = "3.13"
'Global Const FechaModificacion = "07/06/2010"
'Global Const UltimaModificacion = " " 'EGO - Liz Oviedo
''                                      Modificacion de Modelo 602 Interface de familiares
''                                      Se cambio el campo opcional de a�o de DDJJ por los campos desde y hasta fecha de DDJJ.

'Global Const Version = "3.12"
'Global Const FechaModificacion = "19/05/2010"
'Global Const UltimaModificacion = " " 'Cesar Stankunas - Se modific� el Modelo 245 - Se cambi� la l�gica para guardar las novedades de ajustes (se pisa, se deja como

'est� o se suma).

'Global Const Version = "3.11"
'Global Const FechaModificacion = "18/05/2010"
'Global Const UltimaModificacion = " " 'EGO - Liz Oviedo
''                                      Se agrego el modelo Modelo 322 Dependencia de Estructuras
''                                      Importacion de Dependencia de Estructuras.

'Global Const Version = "3.10"
'Global Const FechaModificacion = "16/04/2010"
'Global Const UltimaModificacion = " " 'Margiotta Emanuel
''                                      Modelo 262 Dias Correspondientes de Vacaciones
''                                      Los periodos de vacaciones ahora tienen alcance por estructura

'Global Const Version = "3.09"
'Global Const FechaModificacion = "12/04/2010"
'Global Const UltimaModificacion = " " 'FGZ
''                                      Modelo 605 Migracion de empleados
''                                      Antes si no tenia nro de cuenta no creaba la cuenta bancaria, ahora
''                                       crea la cuenta si tiene cta bancaria o CBU (ademas del banco y la forma de pago)

'Global Const Version = "3.08"
'Global Const FechaModificacion = "08/04/2010"
'Global Const UltimaModificacion = " " 'FGZ
''                                      Modelo 230 Pedido de Vacaciones
''                                      Los periodos de vacaciones ahora tienen alcance por estructura

'Global Const Version = "3.07"
'Global Const FechaModificacion = "19/03/2010"
'Global Const UltimaModificacion = " " 'Martin Ferraro - se truncan los campos zona y localidad en 60 en las interfaces
''                                      600, 602, 605, 606, 607, 608, 909, 663
''                                      Se genero nuevamente la interfaz 318
''                                      interfaz 233 se agrego al log el legajo
''                                      Se nivelo con una version anterior porque no existia la interfaz 282

'Global Const Version = "3.06"
'Global Const FechaModificacion = "22/02/2010"
'Global Const UltimaModificacion = " " 'Martin Ferraro - 317 - Cambios en la interfaz de distribucion contable

'Global Const Version = "3.05"
'Global Const FechaModificacion = "18/02/2010"
'Global Const UltimaModificacion = " " 'Martin Ferraro - 605 - Fecha de Nac Obligatoria. No permitir la carga de empleados menores a 14
'                                                       605 - No hacer nada con las fases cuando existe el empleado
'                                                       605 - Tope de cant de empleados a subir

'Global Const Version = "3.04"
'Global Const FechaModificacion = "16/02/2010" '08/02/2010
'Global Const UltimaModificacion = " " 'Lisandro Moro - Correccion modelo 658 -'Correccion en la funcion TraerCodCalendario
                                                      'Correccion en la funcion controlhora
                                                      'Se agregaron los logs de error al log general.
                                                      '657 -  calnroactualizar_calendario_participante : Se agrego la validacion si el ya existe la relacion Ternro, calnro
                                                      '655 - Se agregaron los tipo de modulo a la interfase, se creo la fn traercodtipomodulo

'Global Const Version = "3.03"
'Global Const FechaModificacion = "12/02/2010"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se modifico nuevamente la interfaz 211 para que si viene una novedad sin vig y en
'                                                       la base existe la misma pero con vig las levante y viceversa. Ademas permite multiples
'                                                       nov con vig con distintas fechas

'Global Const Version = "3.02"
'Global Const FechaModificacion = "10/02/2010"
'Global Const UltimaModificacion = " " 'Martin Ferraro - AsignarEstructura_NEW - No cerrar los his_estructura
'                                                       211 - Permitir novedades repetidas

'Global Const Version = "3.01"
'Global Const FechaModificacion = "13/01/2010"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se modifico la interfaz 317 - Distribucion Contable.
'                                      No permitia pisar distribuciones
'                                      Cuando controlaba total del % lo hacia mal
'                                      Permitia cargar estructuras para cualquier modelo

'Global Const Version = "3.00"
'Global Const FechaModificacion = "11/12/2009"
'Global Const UltimaModificacion = " " 'FGZ - Se modific� el modelo 217 - Interface de Vales
'                                       Se agreg� el periodo de descuento al modelo

'Global Const Version = "2.99"
'Global Const FechaModificacion = "11/12/2009"
'Global Const UltimaModificacion = " " 'FGZ - Se modific� el modelo 217 - Interface de Vales

'Global Const Version = "2.98"
'Global Const FechaModificacion = "11/12/2009"
'Global Const UltimaModificacion = " " 'FGZ - Se agreg� un nuevo Modelo 285 - Interface de Francos Compensatorios

'Global Const Version = "2.97"
'Global Const FechaModificacion = "11/12/2009"
'Global Const UltimaModificacion = " " 'FGZ - Se modific� el modelo 217 - Interface de Vales
'                                      '    - Se modific� el modelo 318 - Interface Valor Plan Obra Social


'Global Const Version = "2.96"
'Global Const FechaModificacion = "11/12/2009"
'Global Const UltimaModificacion = " " 'Stankunas Cesar - Se cre� el modelo 318 - Interface Valor Plan Obra Social - PRICE

'Global Const Version = "2.95"
'Global Const FechaModificacion = "07/12/2009"
'Global Const UltimaModificacion = " " 'Stankunas Cesar - Se cre� el modelo 317 - Interface Distribuci�n Contable - PRICE

'Global Const Version = "2.94"
'Global Const FechaModificacion = "04/12/2009"
'Global Const UltimaModificacion = " " 'Stankunas Cesar - Se modific� la Interface de Prestamos, modelo 229 - Se cambio el tipo de dato de "numero de comprobante" (Nro_Comprobante) a string para que acepte caracteres alfanumericos.

'Global Const Version = "2.93"
'Global Const FechaModificacion = "17/11/2009"
'Global Const UltimaModificacion = " " 'Manuel Lopez - Se corrigio error perf_conc en el modelo 211

'Global Const Version = "2.92"
'Global Const FechaModificacion = "13/11/2009"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se corrigio error de update sin where en el modelo 211

'Global Const Version = "2.91"
'Global Const FechaModificacion = "06/11/2009"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se modifico el modelo 211 - Validaciones de motivo y tipo

'Global Const Version = "2.90"
'Global Const FechaModificacion = "29/10/2009"
'Global Const UltimaModificacion = " " 'Cesar Stankunas - Se cre� el modelo 248 - Carga de Recibos - Custom Farmacity


'Global Const Version = "2.89"
'Global Const FechaModificacion = "08/10/2009"
'Global Const UltimaModificacion = " " 'FGZ  - Correccion en Modelo 635 Importaci�n de Titulos

'Global Const Version = "2.88"
'Global Const FechaModificacion = "09/09/2009"
'Global Const UltimaModificacion = " " 'FGZ  - Correccion en Modelo 659 Importaci�n de asistencias de Capacitaci�n

'Global Const Version = "2.87"
'Global Const FechaModificacion = "25/08/2009"
'Global Const UltimaModificacion = " " 'FGZ  - Unificacion de fuentes (se habian desarrollado cambios sobre fuentes viejos)

'Global Const Version = "2.86"
'Global Const FechaModificacion = "12/08/2009"
'Global Const UltimaModificacion = " " 'FGZ  - Se cre� el modelo 315 - Historico de Ranking - Custom IAMSA

'Global Const Version = "2.85"
'Global Const FechaModificacion = "12/08/2009"
'Global Const UltimaModificacion = " " 'FGZ  - Se modific� el modelo 245
''                                   Estaba mal implementada la seguridad por conceptos.

'Global Const Version = "2.84"
'Global Const FechaModificacion = "23/07/2009"
'Global Const UltimaModificacion = " " 'Martin Ferraro - LineaModelo_211 -
''                                                       Se cambio opcion de pisar novedad
''                                                       Interfaz 606 607 608 Se agrego variable pisa
''                                                       AsignarEstrcturaNew si empleado esta de baja no se cierran estruct


'Global Const Version = "2.83"
'Global Const FechaModificacion = "02/07/2009"
'Global Const UltimaModificacion = " " 'Martin Ferraro - LineaModelo_653 - Interfaz reporta A
'                                                       Utilizaba LineaError.writeline Mid(strReg, 1, Len(strReg)) que daba error
'                                                       Se reemplazo por Call Escribir_Log("floge", LineaCarga, "1", Texto, Tabs, strReg)

'Global Const Version = "2.82"
'Global Const FechaModificacion = "24/06/2009"
'Global Const UltimaModificacion = " " 'Diego Nu�ez - Se modific� el modelo 609 para la migraci�n de empleados de RADIOTRONICA. A los empleados ya
''                                                    existentes se le cierran las estructuras anteriores y se crean nuevas, salvo en el caso que se
''                                                    seleccione la opci�n de sobreescribir empleados, en cuyo caso se sobreescribe el hist�rico de
''                                                    estructuras.

'Global Const Version = "2.81"
'Global Const FechaModificacion = "22/05/2009"
'Global Const UltimaModificacion = " " 'Diego Nu�ez - Se cambiaron en todos los meodelos que utilizaban las variables localidad y zona, el tama�o
'                                                    de las mismas, extendiendolas a 60 caracteres.

'Global Const Version = "2.80"
'Global Const FechaModificacion = "14/04/2009"
'Global Const UltimaModificacion = " " 'Diego Nu�ez - Se agreg� un nuevo modelo (281) que realiza las altas masivas de remuneraciones seg�n
'                                                    las especificaciones recibidas de LA CAJA

'Global Const Version = "2.79"
'Global Const FechaModificacion = "26/03/2009"
'Global Const UltimaModificacion = " " 'Diego Nu�ez - Se modific� modelo 211 que realiza las migraciones de novedades, para que sea capaz
'                                                    de manejar tipos de motivos y motivos. Adem�s se se agregaron 5 nuevos formatos al archivo
'                                                    de entrada.

'Global Const Version = "2.78"
'Global Const FechaModificacion = "25/03/2009"
'Global Const UltimaModificacion = " " 'FGZ- Se modificaron los modelos 236 y 237
''                                             Modelo 236: IMPORTACION DE Totales de Cantidad de BULTOS  a  RH Pro
''                                             Modelo 237: IMPORTACION DE Detalle de Cantidad de BULTOS  a  RH Pro

'Global Const Version = "2.77"
'Global Const FechaModificacion = "12/03/2009"
'Global Const UltimaModificacion = " " 'Diego Nu�ez - Se agreg� un nuevo modelo (609) que realiza las migraciones de empleados seg�n
'                                                    las especificaciones recibidas de RADIOTRONICA

'Global Const Version = "2.76"
'Global Const FechaModificacion = "06/03/2009"
'Global Const UltimaModificacion = " " 'Diego Nu�ez - Se agreg� una funci�n para que al validar el cuil en una importaci�n
'                                                    de empleados, se verifique que el nro de documento est� incluido en
'                                                    el nro de cuil. Si esto no ocurre, se recalcula el numero de cuil basandose
'                                                    en el numero de documento.

'Global Const Version = "2.75"
'Global Const FechaModificacion = "03/03/2009"
'Global Const UltimaModificacion = " " 'Lisandro Moro - Se modific� el modelo 245
'                                   Se Modifico la seguridad por conceptos.
'                                   Si no hay ningun registro en concepto perfil asumo que NO hay permiso.

'Global Const Version = "2.74"
'Global Const FechaModificacion = "23/02/2009"
'Global Const UltimaModificacion = " " 'FGZ
''           Empleados 605 - Validacion del CUIL configurable


'Global Const Version = "2.73"
'Global Const FechaModificacion = "13/02/2009"
'Global Const UltimaModificacion = " " 'FGZ - Encriptacion de string de conexion
''      Nuevos Modelos para Colombia
''           Familiares 910
''           Empleados 608

'Global Const Version = "2.72"
'Global Const FechaModificacion = "15/12/2008"
'Global Const UltimaModificacion = " " 'Lisandro Moro - Se modific� el modelo 245
''                                   Se agrego la seguridad por conceptos.

'Global Const Version = "2.71"
'Global Const FechaModificacion = "27/10/2008"
'Global Const UltimaModificacion = " " 'FGZ - Se modific� el modelo 615 - LineaModelo_615 - migracion de DDJJ
''   No se estaba actualizando el campo prorratea
''    El Formato anterior era:
''    Legajo;A�o;Fecha Desde;Fecha Hasta;N�mero de Item;Monto;Cuit;Razon Social;
''
''    Ahora es:
''        Legajo;A�o;Fecha Desde;Fecha Hasta;N�mero de Item;Monto;Cuit;Razon Social;Prorratea

'Global Const Version = "2.70"
'Global Const FechaModificacion = "20/10/2008"
'Global Const UltimaModificacion = " " 'Diego Nu�ez - Se modific� el modelo 250 - Alta Pl�stica
''Se modific� el c�digo fuente de la interfase general del modelo 250 para que fuese capaz de manejar
''valores num�ricos con configuraci�n espa�ola, en un sistema con configuraci�n regional Inglesa.


'Global Const Version = "2.69"
'Global Const FechaModificacion = "09/09/2008"
'Global Const UltimaModificacion = " " 'Lisandro Moro - Se agrego el modelo 665 - Actionline

'Global Const Version = "2.68"
'Global Const FechaModificacion = "19/08/2008"
'Global Const UltimaModificacion = " " 'MB - modelo 657 y 659 saque el control de legajo <5 .
'                                       en funcion controlnum por esNum

'Global Const Version = "2.67"
'Global Const FechaModificacion = "04/07/2008"
'Global Const UltimaModificacion = " " 'FGZ - modelo 602(Interfaz de Familiares).
'                                       Se agreg� campo "fecha de inicio del Vinculo"

'Global Const Version = "2.66"
'Global Const FechaModificacion = "28/04/2008"
'Global Const UltimaModificacion = " " 'FGZ - Nuevo modelo 663(Interfaz de prestadores medicos) .

'Global Const Version = "2.65"
'Global Const FechaModificacion = "28/04/2008"
'Global Const UltimaModificacion = " " 'Lisandro Moro - Se corrigio la funcio ya que calculaba un dia de mas.

'Global Const Version = "2.64"
'Global Const FechaModificacion = "27/03/2008"
'Global Const UltimaModificacion = " " 'Gustavo Ring - Se modifico el m�delo 312.- Tarjetas. Ahora se inicializa "devuelta" en 0


'Global Const Version = "2.63"
'Global Const FechaModificacion = "28/02/2008"
'Global Const UltimaModificacion = " " 'Gustavo Ring - Se creo el m�delo 662 interfaz de ctas de mails de empleados


'Global Const Version = "2.62"
'Global Const FechaModificacion = "11/01/2008"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se modifico el modelo 243 (Importacion Ctas Bca). Cuando pisa la cuenta porque ya existe la misma no pisaba la forma de pago.

'Global Const Version = "2.61"
'Global Const FechaModificacion = "11/01/2008"
'Global Const UltimaModificacion = " " 'Gustavo Ring - Se cambio el modelo 662 al nro 602

'Global Const Version = "2.60"
'Global Const FechaModificacion = "07/01/2008"
'Global Const UltimaModificacion = " " 'Gustavo Ring - Se agrego el modelo 662 - Importaci�n de Familiares - Standard

'Global Const Version = "2.59"
'Global Const FechaModificacion = "28/12/2007"
'Global Const UltimaModificacion = " " 'Lisandro Moro - Se corriogio al momento de insertar los titulos en el modelo 635, por una modificaciones en el abm de titulos

'Global Const Version = "2.58"
'Global Const FechaModificacion = "20/12/2007"
'Global Const UltimaModificacion = " " 'FGZ - Modelo 253 (Interfase de Embargos).
''                                       Cuando chequeaba la maxima cantidad de embargos asociados al empleado no estaba armando bien el query
''                                       faltaba para que empleado por lo cual la interfaz nunca levantaba un ermbargo si algun empleado tenia un ebargo asociado.
''                                   Esto confirma que nadie lo estaba usando..
''                                   Ademas Le agregu� 2 campos mas al formato de Porcentaje (acumulador y porcentaje)
''                           Formato
''                               Porcentaje; ternro; embesext; tpenro; embest; embanioini; embmesini; embquinini; Monto; cuotas; embaniofin; embmesfin; embquinfin;

'monto; acumulador;porcentaje

'Global Const Version = "2.57"
'Global Const FechaModificacion = "10/09/2007"
'Global Const UltimaModificacion = " " 'FGZ - Modelo 211 (Interfase de Novedades). Se sac� del estandar lo que se agreg� en la version 2.49
''                                       Ese modelo 211 lo guard� como 211_Custom para recuperarlo y crearlo con un nuevo nro de modelo
''                                       Lo cre� como modelo CUSTOM 313

                                      
'Global Const Version = "2.56"
'Global Const FechaModificacion = "06/09/2007"
'Global Const UltimaModificacion = " " 'FGZ - Modelo 211 (Interfase de Novedades). Se sac� del estandar lo que se agreg� en la version 2.49
''                                       Ese modelo 211 lo guard� como 211_Custom para recuperarlo y crearlo con un nuevo nro de modelo

'Global Const Version = "2.55"
'Global Const FechaModificacion = "05/09/2007"
'Global Const UltimaModificacion = " " 'Gustavo Ring - Se creo el modelo 280: Interface de Bandas Salariales - Estandar

'Global Const Version = "2.54"
'Global Const FechaModificacion = "03/09/2007"
'Global Const UltimaModificacion = " " 'Gustavo Ring - Se creo el modelo 661: Importaci�n de Im�genes del Empleado - Estandar

'Global Const Version = "2.53"
'Global Const FechaModificacion = "17/08/2007"
'Global Const UltimaModificacion = " " 'FGZ - Se creo el modelo 660: Importaci�n de Cuestionarios de Capacitacion - Estandar

'Global Const Version = "2.52"
'Global Const FechaModificacion = "26/07/2007"
'Global Const UltimaModificacion = " " 'Gustavo Ring - Migraci�n Cap 657 - Se relaciono participantes con calendarios, se actualiza la cantidad de participantes y se agregaron Logs.
                         '             - Migraci�n Cap 655,656,657,658,659 se puso como opcional el �ltimo separador. Tambien se inicializa fecha de

'inicio y fin del evento

'Global Const Version = "2.51"
'Global Const FechaModificacion = "21/06/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se creo el modelo 909 para migracion de familiares de URUGUAY

'Global Const Version = "2.50"
'Global Const FechaModificacion = "19/06/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se agrego una columna mas al modelo 279. Acumulados Diarios para SMT

'Global Const Version = "2.49"
'Global Const FechaModificacion = "14/06/2007"
'Global Const UltimaModificacion = " " 'Gustavo Ring - Se agrego seguridad por concepto a los modelos 211 y 245 Novedades y Novedades Ajuste

'Global Const Version = "2.48"
'Global Const FechaModificacion = "29/05/2007"
'Global Const UltimaModificacion = " " 'Gustavo Ring - Migraci�n Cap modelos 655,656,657,658,659

'Global Const Version = "2.47"
'Global Const FechaModificacion = "30/04/2007"
'Global Const UltimaModificacion = " " 'Maximiliano Breglia - Se modific� el modelo 635 de Titulo que no andaba pq estaba mal el sql

'Global Const Version = "2.46"
'Global Const FechaModificacion = "26/04/2007"
'Global Const UltimaModificacion = " " 'Fernando Favre - Se agrego el modelo 279. Acumulados Diarios para SMT

'Global Const Version = "2.45"
'Global Const FechaModificacion = "01/03/2007"
'Global Const UltimaModificacion = " " 'Lisandro Moro - Se corrigio el modelo 217. Importacion Vales
                                      'Se agrego un campo mas al archivo de importacion de vales, el campo Moneda.
                              
'Global Const Version = "2.44"
'Global Const FechaModificacion = "26/02/2007"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se corrigio el modelo 607. Migracion Empleados Chile
                                       'No grababa en terape2 y el ternom2 en la tabla empleado. Tambien se cambio para que modifique el empleado cuando el rut ya existe en la BD

'Global Const Version = "2.43"
'Global Const FechaModificacion = "22/12/2006"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se Modifico el modelo 607. Migracion Empleados Chile
                                                 '   Tambien  se le agrego a Validalocalidad un parametro necesario para el modelo 607

'Global Const Version = "2.42"
'Global Const FechaModificacion = "22/11/2006"
'Global Const UltimaModificacion = " " 'Diego Rosso - Se corrigio el modelo 605 para que cuando se grabe un nuevo contrato en la tabla tipocont ponga la descripcion abreviada en el campo tcdesc
                                              ' y que grabe tambien el campo leynro a 1

'Global Const Version = "2.41"
'Global Const FechaModificacion = "26/10/2006"
'Global Const UltimaModificacion = " " 'Maxi - Se corrigio el modelo 615 que los ultimos 2 campos generaba error...en la validacion se puso <=

'Global Const Version = "2.40"
'Global Const FechaModificacion = "10/10/2006"
'Global Const UltimaModificacion = " " 'Se implemento el Modelo 312 - Importacion masiva de tarjetas.

'Global Const Version = "2.39"
'Global Const FechaModificacion = "07/09/2006"
'Global Const UltimaModificacion = " " 'Se creo el Modelo 604 - Empleados Walmart

'Global Const Version = "2.38"
'Global Const FechaModificacion = "26/09/2006"
'Global Const UltimaModificacion = " " 'Se modifico el Modelo 308 - Se modifico el Modelo 309 -

'Global Const Version = "2.37"
'Global Const FechaModificacion = "14/09/2006"
'Global Const UltimaModificacion = " " 'Se implemento el Modelo 308 - Importacion Acum. Diario. Legajo;Apellido y Nombre;Fecha;Desc. Tipo Hora;Tipo Hora;Cantidad.Tiene comillas " en los campos
                                      'Se implemento el Modelo 309 - Importacion Acum. Diario. Legajo;Apellido y Nombre;Fecha;Desc. Tipo Hora;Tipo Hora;Cantidad

'Global Const Version = "2.36"
'Global Const FechaModificacion = "14/09/2006"
'Global Const UltimaModificacion = " " 'Se implemento el Modelo 123 - Importacion Acum. Diario. Legajo;Fecha;Tipo Hora;Cantidad

'Global Const Version = "2.35"
'Global Const FechaModificacion = "10/08/2006"
'Global Const UltimaModificacion = " " 'Modelo 605 - Empleados. La estructura formaliq no creaba el complemento


' Ultima Mod.: 21/03/2006 - FGZ
' Descripcion: Fecha de Baja no es obligatoria
' Ultima Mod.: 24/05/2006 - LA
' Descripcion: Agregar la opcion de Pisar la Estructura - en mig his_estructura

'Global Const Version = "2.34"
'Global Const FechaModificacion = "04/05/2006"
'Global Const UltimaModificacion = " " 'Modelo 265 - Interface de Documentos. Fecha de Vencimiento no abligatoria

'Global Const Version = "2.33"
'Global Const FechaModificacion = "24/05/2006"
'Global Const UltimaModificacion = " " 'Modelo 630 - Interface Historico Estructura. Se agrego la opcion de que se reemplaze la estructura.

'Global Const Version = "2.32"
'Global Const FechaModificacion = "27/04/2006"
'Global Const UltimaModificacion = " " 'Modelo 303 - Interface Empleados de Agencia Halliburton. Se agrego la agencia en la tabla empelado. Se agregaron los guiones al cuil.


'Global Const Version = "2.31"
'Global Const FechaModificacion = "25/04/2006"
'Global Const UltimaModificacion = " " 'Modelo 303 - Interface Empleados de Agencia Halliburton

'Global Const Version = "2.30"
'Global Const FechaModificacion = "02/03/2006"
'Global Const UltimaModificacion = " " 'Modelo 256 - Engagement. Modificacion en lectura del campo final

'Global Const Version = "2.29"
'Global Const FechaModificacion = "01/03/2006"
'Global Const UltimaModificacion = " " 'Modelo 226. Paso de estandar a Custom de Deloitte
'                                      'Modelo 275. Nuevo modelo de Postulantes estandar.

'Global Const Version = "2.28"
'Global Const FechaModificacion = "02/02/2006"
'Global Const UltimaModificacion = " " 'Modelo 226. No cargaba la causa de despido en los trabajos anteriores.
'                                      'Errores en la SQL del UPDATE de idiomas
'                                      'Especializaciones. No se utiliza el arreglo.
'                                      'Se agregaron mas logs.

'Global Const Version = "2.27"
'Global Const FechaModificacion = "31/01/2006"
'Global Const UltimaModificacion = " " 'Modelo 211. Al ingresar novedades con vigencia (con la opcion reemplazar novedades activa en el modelo),
                                      'elimina todas las novedades con vigencia que coincidan en algun dia con la novedad que se desea ingresar.


'Global Const Version = "2.26"
'Global Const FechaModificacion = "16/01/2006"
'Global Const UltimaModificacion = " " 'Nuevo modelo 272. Interface novedades SPEC. Sin Separadores

'Global Const Version = "2.25"
'Global Const FechaModificacion = "29/12/2005"
'Global Const UltimaModificacion = " " 'Nuevo modelo 272. Interface novedades SPEC. CUSTOM

'Global Const Version = "2.24"
'Global Const FechaModificacion = "28/12/2005"
'Global Const UltimaModificacion = " " 'Nuevo modelo 272. Interface novedades SPEC. CUSTOM

'Global Const Version = "2.23"
'Global Const FechaModificacion = "26/12/2005"
'Global Const UltimaModificacion = " " 'Se agrego en el log la configuracion regional de la maquina


'Global Const Version = "2.22"
'Global Const FechaModificacion = "15/12/2005"
'Global Const UltimaModificacion = " " 'Modelo 211. InsertarError en el manejador de errores


'Global Const Version = "2.21"
'Global Const FechaModificacion = "02/11/2005"
'Global Const UltimaModificacion = " " 'Migracion de DDJJ Modelo 615. Conversion de fecha

'Global Const Version = "2.20"
'Global Const FechaModificacion = "02/11/2005"
'Global Const UltimaModificacion = " " 'Migracion de desgloce de ganancias modelo 620. Chequeo por nulo, 0 o n/a

'Global Const Version = "2.19"
'Global Const FechaModificacion = "01/11/2005"
'Global Const UltimaModificacion = " " 'Migracion de desgloce de ganancias modelo 620. Conversion de fecha


'Global Const Version = "2.18"
'Global Const FechaModificacion = "01/11/2005"
'Global Const UltimaModificacion = " " 'Migracion de Legajos modelo 605. Longitud de la sucursal del banco


'Global Const Version = "2.17"
'Global Const FechaModificacion = "28/10/2005"
'Global Const UltimaModificacion = " " 'Interface Licencias modelo 233
'                                      'Correccion Integer por Long proyecto entero
'                                      'Correccion cInt() por clng() proyecto entero

'Global Const Version = "2.16"
'Global Const FechaModificacion = "24/10/2005"
'Global Const UltimaModificacion = " " 'Interface Deloitte modelo 268

'Global Const Version = "2.15"
'Global Const FechaModificacion = "24/10/2005"
'Global Const UltimaModificacion = " " 'Correccion Integer por Long

'Global Const Version = "2.14"
'Global Const FechaModificacion = "20/10/2005"
'Global Const UltimaModificacion = " " 'Interfaces Wella modelos 266/7

'Global Const Version = "2.13"
'Global Const FechaModificacion = "13/10/2005"
'Global Const UltimaModificacion = " " 'modelo 620. Migracion de Desgloce de ganancias (desliq)

'Global Const Version = "2.12"
'Global Const FechaModificacion = "04/10/2005"
'Global Const UltimaModificacion = " " 'Nuevo modelo 265. Interface Docs. ESTANDAR

'Global Const Version = "2.11"
'Global Const FechaModificacion = "21/09/2005"
'Global Const UltimaModificacion = "Update" 'sexo en las interfaces de postulantes

'Global Const Version = "2.10"
'Global Const FechaModificacion = "25/07/2005"
'Global Const UltimaModificacion = "Inicial"


'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'FGZ - 19/06/2012 -------
Public Type Tconfrep
    columna  As Variant
    ConcCod  As Variant
    ConcNro  As Variant
    tpanro  As Variant
End Type
Global ArrConfrep() As Tconfrep
Global TDoc As Long
Global TERevendedor As Long

'FGZ - 19/06/2012 -------

Global crpNro As Long
Global RegLeidos As Long
Global RegError As Long
Global RegWarnings As Long
Global RegFecha As Date
Global NroProceso As Long

Global f
'Global HuboError As Boolean
Global Path
Global NArchivo
Global NroLinea As Long
Global LineaCarga As Long

Global separador As String
Global SeparadorDecimal As String
Global UsaEncabezado As Boolean

Global ProcPendiente As Boolean
Global ErroresNov As Boolean

Global ErrCarga
Global LineaError
Global LineaOK

Global PisaNovedad As Boolean
Global PisaPlan As Boolean
Global AccionNovedad As Integer
Global AccionNovedadAju As Integer
Global PisarDistCont As Integer
Global Vigencia As Boolean
Global Vigencia_Desde As String
Global Vigencia_Hasta As String
Global Pisa As Boolean
Global TikPedNro As Long
Global nombrearchivo As String
Global acuNro As Long 'se usa en el modelo 216 de Citrusvil y se carga por confrep
Global nro_ModOrg  As Long

Global NroModelo As Long
Global DescripcionModelo As String
Global Primera_Vez As Boolean
Global Banco As Long
'Global usuario As String
Global EncontroAlguno As Boolean

Global NroColumna As Long
Global Tabs As Long

Global Pliqnro As Long
Global Sql_VistaEmpleado
Global ModeloDom As Integer 'Se utiliza para los modelos de domicilios. Modelo 668
            
'FGZ - 10/01/2012 ------------------------------------------
'Global Lista_WC As New Collection
'Global Clave_Lista_WC As Long



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial de Interface.
' Autor      : FGZ
' Fecha      : 29/07/2004
' Ultima Mod.: 27/01/2012 - Gonzalez Nicol�s
' Descripcion: Habilita Traducciones si MI este activo. Se agreg� la funcion EscribeLogMI() en las etiquetas.
' ---------------------------------------------------------------------------------------------
Dim objconnMain As New ADODB.Connection
Dim strCmdLine
Dim Nombre_Arch As String
Dim Nombre_Arch_Errores As String
Dim Nombre_Arch_Correctos As String
Dim rs_batch_proceso As New ADODB.Recordset
Dim bprcparam As String
Dim PID As String
Dim ArrParametros
    
    strCmdLine = Command()
    ArrParametros = Split(strCmdLine, " ", -1)
    If UBound(ArrParametros) > 1 Then
        If IsNumeric(ArrParametros(0)) Then
            NroProcesoBatch = ArrParametros(0)
            Etiqueta = ArrParametros(1)
            EncriptStrconexion = CBool(ArrParametros(2))
            c_seed = ArrParametros(2)
        Else
            Exit Sub
        End If
    Else
        If UBound(ArrParametros) > 0 Then
            If IsNumeric(ArrParametros(0)) Then
                NroProcesoBatch = ArrParametros(0)
                Etiqueta = ArrParametros(1)
            Else
                Exit Sub
            End If
        Else
            If IsNumeric(strCmdLine) Then
                NroProcesoBatch = strCmdLine
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    'Obtiene los datos de como esta configurado el servidor actualmente
    Call ObtenerConfiguracionRegional
    
   
    
    Nombre_Arch = PathFLog & "Importacion_Exportacion" & "-" & NroProcesoBatch & ".log"
    Nombre_Arch_Errores = PathFLog & "Lineas_Errores" & "-" & NroProcesoBatch & ".log"
    Nombre_Arch_Correctos = PathFLog & "Lineas_Procesadas" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    Set FlogE = fs.CreateTextFile(Nombre_Arch_Errores, True)
    Set FlogP = fs.CreateTextFile(Nombre_Arch_Correctos, True)
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error Resume Next
    OpenConnection strconexion, objconnProgreso
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    On Error GoTo ME_Main
    
    '===========================
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    '===========================
    TiempoInicialProceso = GetTickCount
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprctiempo = 0,bprcprogreso = 0, bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 23 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    ProcPendiente = False
    ErroresNov = False
    Primera_Vez = False
    tplaorden = 0
    
    If Not rs_batch_proceso.EOF Then
        bprcparam = rs_batch_proceso!bprcparam
        'Flog.writeline Espacios(Tabulador * 0) & "Parametros del proceso = " & bprcparam
        usuario = rs_batch_proceso!iduser
        rs_batch_proceso.Close
        
        '____________________________________________________
        'NG - VALIDA QUE ESTE ACTIVO LA TRADUCCION A MULTI IDIOMA
        Call Valida_MultiIdiomaActivo(usuario)
        
        '-----------------------------------------------------------------
        'ESCRIBO ENCABEZADO CON MI
        '-----------------------------------------------------------------
        Flog.writeline "-----------------------------------------------------------------"
        Flog.writeline EscribeLogMI("Version") & " = " & Version
        Flog.writeline EscribeLogMI("Modificaci�n") & " = " & UltimaModificacion
        Flog.writeline EscribeLogMI("Fecha") & " = " & FechaModificacion
        Flog.writeline "-----------------------------------------------------------------"
        Flog.writeline EscribeLogMI("Numero") & ", " & EscribeLogMI("separador decimal") & "    : " & NumeroSeparadorDecimal
        Flog.writeline EscribeLogMI("Numero") & ", " & EscribeLogMI("separador de miles") & "   : " & NumeroSeparadorMiles
        Flog.writeline EscribeLogMI("Moneda") & ", " & EscribeLogMI("separador decimal") & "    : " & MonedaSeparadorDecimal
        Flog.writeline EscribeLogMI("Moneda") & ", " & EscribeLogMI("separador de miles") & "   : " & MonedaSeparadorMiles
        Flog.writeline EscribeLogMI("Formato de Fecha del Servidor") & ": " & FormatoDeFechaCorto
        Flog.writeline "-----------------------------------------------------------------"
        Flog.writeline
        Flog.writeline "PID = " & PID
        '-----------------------------------------------------------------
        '-----------------------------------------------------------------
        Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Parametros del proceso") & " = " & bprcparam
        
        'EAM- Crea la vista del empelado para que se usan en algunos modelos. (nombre vista:v_empleadoproc)
        'Call CreaVistaEmpleadoProceso("V_EMPLEADO", usuario)
        
        
        Set rs_batch_proceso = Nothing
        'Flog.Writeline Espacios(Tabulador * 1) & "Levanta parametros"
        Call LevantarParamteros(bprcparam)
        'Flog.Writeline Espacios(Tabulador * 1) & "fin levanta parametros"
        LineaCarga = 0
        
        Call ComenzarTransferencia
    Else
        '-----------------------------------------------------------------
        'ESCRIBO ENCABEZADO SIN MI - CUANDO NO ENCUENTRA PROCESO
        '-----------------------------------------------------------------
        Flog.writeline "-----------------------------------------------------------------"
        Flog.writeline "Version = " & Version
        Flog.writeline "Modificaci�n = " & UltimaModificacion
        Flog.writeline "Fecha = " & FechaModificacion
        Flog.writeline "-----------------------------------------------------------------"
        Flog.writeline "Numero, separador decimal    : " & NumeroSeparadorDecimal
        Flog.writeline "Numero, separador de miles   : " & NumeroSeparadorMiles
        Flog.writeline "Moneda, separador decimal    : " & MonedaSeparadorDecimal
        Flog.writeline "Moneda, separador de miles   : " & MonedaSeparadorMiles
        Flog.writeline "Formato de Fecha del Servidor: " & FormatoDeFechaCorto
        Flog.writeline "-----------------------------------------------------------------"
        Flog.writeline
        Flog.writeline "PID = " & PID
        '-----------------------------------------------------------------
        '-----------------------------------------------------------------
    End If
    
    
Final:
    TiempoAcumulado = GetTickCount
    If ProcPendiente Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprctiempo ='0', bprcprogreso = 0, bprcestado = 'Pendiente' WHERE bpronro = " & NroProcesoBatch
    Else
        If Not HuboError Then
            If ErroresNov Then
                StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcprogreso = 100, bprcestado = 'Incompleto' WHERE bpronro = " & NroProcesoBatch
            Else
                StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcprogreso = 100, bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
            End If
        Else
            StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcprogreso = 100, bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
        End If
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Resumen
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "===================================================================="
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Lineas Leidas") & "    : " & RegLeidos
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Lineas Erroneas") & "  : " & RegError
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Warnings") & "         : " & RegWarnings
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Lineas Procesadas") & ": " & RegLeidos - RegError
    Flog.writeline Espacios(Tabulador * 0) & "===================================================================="
    objConn.Close
    objconnProgreso.Close
    Flog.Close
    FlogE.Close
    FlogP.Close
    End
    
ME_Main:
    HuboError = True
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & " " & EscribeLogMI("Error General") & " " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("SQL Ejecutado") & ": " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Final
End Sub





Private Sub LeeArchivo(ByVal nombrearchivo As String)
' Descripcion:
' Autor      : ?
' Fecha      : ?
' Modificado : 13/02/2012 - Gonzalez Nicol�s - Se agreg� Multilenguaje
' Modificado : 17/06/2014 - Gonzalez Nicol�s - Se Formatea incporc a 14 decimales.

Const ForReading = 1
Const TristateFalse = 0
Dim strLinea As String
Dim Archivo_Aux As String
Dim rs_Lineas As New ADODB.Recordset
Dim rs_modelo As New ADODB.Recordset
Dim Ciclos As Long

    If App.PrevInstance Then
        'Flog.writeline Espacios(Tabulador * 0) & "Hay una instancia previa del proceso corriendo - El proceso actual queda en estado Pendiente. "
        Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Hay una instancia previa del proceso corriendo") & " - " & EscribeLogMI("El proceso actual queda en estado Pendiente.")
        ProcPendiente = True ' para dejar el proceso pendiente
        Exit Sub
    End If
    
    'Espero hasta que se crea el archivo
    
    On Error Resume Next
    Err.Number = 1
    Ciclos = 0
    Do Until Err.Number = 0
        Err.Number = 0
        Set f = fs.GetFile(nombrearchivo)
        If f.Size = 0 Then
            If Ciclos > 100 Then
                Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("No anda el getfile")
            Else
                Err.Number = 1
                Ciclos = Ciclos + 1
            End If
        End If
    Loop
    On Error GoTo 0
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Archivo creado") & ": " & nombrearchivo
   
   'Abro el archivo
    On Error GoTo CE
    Set f = fs.OpenTextFile(nombrearchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO inter_pin(bpronro,modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      NroProcesoBatch & "," & NroModelo & ",'" & Left(nombrearchivo, 60) & "',0,0," & ConvFecha(Date) & ",'" & Left(DescripcionModelo, 18) & ": " & Date & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "inter_pin")
        Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Ultimo inter_pin") & ": " & crpNro
    Else
        Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("No se pudo abrir el archivo") & ": " & nombrearchivo
    End If
                
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, rs_modelo
    If rs_modelo.EOF Then
        Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("No se encontr� el modelo") & ": " & NroModelo
        Exit Sub
    End If
                    
    StrSql = "SELECT * FROM modelo_filas WHERE bpronro =" & NroProcesoBatch
    StrSql = StrSql & " ORDER BY fila "
    OpenRecordset StrSql, rs_Lineas
    If Not rs_Lineas.EOF Then
        rs_Lineas.MoveFirst
    Else
        Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("No hay filas seleccionadas")
    End If
    
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = rs_Lineas.RecordCount
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = FormatNumber((99 / CEmpleadosAProc), 14)
    
    Do While Not f.AtEndOfStream And Not rs_Lineas.EOF
        strLinea = f.ReadLine
        NroLinea = NroLinea + 1
        If NroLinea = 1 And UsaEncabezado Then
            strLinea = f.ReadLine
            'NroLinea = NroLinea + 1
            'rs_Lineas.MoveNext
        End If
        If Trim(strLinea) <> "" And NroLinea = rs_Lineas!fila Then
            
            'Flog.Writeline Espacios(Tabulador * 0) & "Linea " & NroLinea
            Select Case rs_modelo!modinterface
                Case 1:
                    Call Insertar_Linea_Segun_Modelo_Estandar(strLinea)
                    RegLeidos = RegLeidos + 1
                Case 2:
                    Call Insertar_Linea_Segun_Modelo_Custom(strLinea)
                    RegLeidos = RegLeidos + 1
                Case 3:
                    Call Insertar_Linea_Segun_Modelo_MigraInicial(strLinea)
                    RegLeidos = RegLeidos + 1
                Case Else
                    'Flog.writeline Espacios(Tabulador * 0) & "El Modelo " & NroModelo & " no tiene configurado el campo modinterface"
                    Flog.writeline Replace(EscribeLogMI("El Modelo @@NUM@@ no tiene configurado el campo modinterface"), "@@NUM@@", NroModelo)
            End Select
            
            rs_Lineas.MoveNext
                     

            
            'Como actualizo el progreso aca si no se cuantas lineas tiene el archivo
            'Incremento el progreso para que el servidor de aplicaciones no vea a este proceso
            'como colgado
            TiempoAcumulado = GetTickCount
            Progreso = Progreso + IncPorc
            Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Progreso") & " = " & CLng(Progreso) & " (" & EscribeLogMI("Incremento") & " = " & IncPorc & ")"
            StrSql = "UPDATE batch_proceso SET bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & "', bprcprogreso = " & CLng(Progreso) & " WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Progreso actualizado")
        End If
        
    Loop
    
    If NroModelo = "287" Then
        'spec carga masiva
            'inserto en batch proceso para que dispare el proceso de lectura de registraciones
            StrSql = " INSERT INTO batch_proceso "
            StrSql = StrSql & " (btprcnro, bprcfecha, bprchora, iduser, bprcfecdesde, bprcfechasta, bprcparam,"
            StrSql = StrSql & " bprcestado, bprcprogreso, bprcfecfin, bprchorafin, bprctiempo, empnro, "
            StrSql = StrSql & " bprcempleados,bprcurgente,bprcTipoModelo) "
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & " (366," & ConvFecha(Date) & ",'" & Format(Now, "hh:mm:ss ") & "','" & usuario & "', " & ConvFecha(Date) & "," & ConvFecha(Date) & " ," & NroProcesoBatch & ", 'Pendiente', null , null, null, null, 0, null,0,null)"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "Inserto en batch_proceso para que dispare el proceso de lectura de registraciones SQL:" & StrSql
            'hasta aca
    End If
    
    StrSql = "UPDATE inter_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    f.Close
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Archivo procesado") & ": " & nombrearchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'Borrar el archivo
    fs.DeleteFile nombrearchivo, True
    
Fin:
    If rs_Lineas.State = adStateOpen Then rs_Lineas.Close
    Set rs_Lineas = Nothing
    Exit Sub
    
CE:
    HuboError = True
    
    MyRollbackTrans
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Error") & ". " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Error") & ": " & Err.Number
    Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("Descripcion") & ": " & Err.Description
    Flog.writeline
    'Flog.writeline Espacios(Tabulador * 0) & "Linea " & RegLeidos & " del archivo procesado"
    Flog.writeline Espacios(Tabulador * 0) & Replace(EscribeLogMI("Linea @@NUM@@ del archivo procesado"), "@@NUM@@", RegLeidos)
    If InStr(1, Err.Description, "ODBC") > 0 Then
        'Fue error de Consulta de SQL
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 0) & EscribeLogMI("SQL Ejecutado") & ": " & StrSql
        Flog.writeline
    End If
    Flog.writeline Espacios(Tabulador * 0) & "**********************************************************"
    Flog.writeline
    GoTo Fin
End Sub


' _____________________________________________________________________________________
' Descripcion: elimina caracteres invalidos de una cadena en base a un conjunto de carateres validos.
'              el parametro ConjCH indica que conjunto de carateres validos se usar�
'              - conjCH = -1 - el conjubnto de caracteres que se entra x parametro (conjCHs)
'              - conjCH = 0 - caracteres valido - nombre
'              - conjCH = 1 - caracteres valido - string
'              - conjCH = 2 - caracteres valido - fechas
'              - conjCH = 3 - caracteres valido - telefono
'              - conjCH = 4 - caracteres valido - mail
' Autor      : Leticia A.
' Fecha      : 24-11-2010
' Modificado : 20/01/2012 - Gonzalez Nicol�s - Se agregaron car�cteres v�lidos faltantes en cadalfabeto2.
'            : 13/02/2012 - Gonzalez Nicol�s - Se agreg� Multilenguaje
'            : 16/10/2012 - FGZ - agregu�  � y � a los caracteres validos
'            : 20/10/2014 - Sebastian Stremel - Se agrega el caracter "&" a la lista de caracteres validos.
'            : 27/10/2014 - Carmen Quintero - Se agrego el caracter "�" a la lista de caracteres validos.
' _____________________________________________________________________________________
Public Function EliminarCHInvalidosII(ByVal cadena As String, ByVal conjCH As Integer, conjCHs As String) As String

Dim conjCH0, conjCH1, conjCH2, conjCH3, conjCH4 As String
Dim conjCaractValidos As String
Dim cadenaAux As String
Dim ch As String
Dim I As Long

Dim cadnumero, cadalfabeto, cadalfabeto2 As String
Dim chInv As String


' _______________________________________________________________________________
' conjunto de carateres validos _________________________________________________

    cadnumero = "0123456789"
    'cadalfabeto = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    'FGZ - 16/10/2012 - agregu�  � y � a los caracteres validos
    cadalfabeto = "abcdefghijklmn�opqrstuvwxyzABCDEFGHIJKLMN�OPQRSTUVWXYZ"
    'cadalfabeto2 = "����������������������"
    cadalfabeto2 = "������������������������������������"
    
    ' caracteres valido - nombre
    conjCH0 = cadalfabeto + cadalfabeto2 + " ()*�`."
    ' caracteres valido - string
    conjCH1 = cadalfabeto + cadalfabeto2 + cadnumero + " !$;,.-/_()&*�`\\<> \n\r"
    ' caracteres valido - Fechas
    conjCH2 = cadnumero + "/-"
    ' caracteres valido - Telefono
    conjCH3 = cadnumero + " ()-*#"
    'caracteres valido - mail
    conjCH4 = cadalfabeto + cadnumero + ".-@_"
    
   

' conjunto de carateres validos _________________________________________________
' _______________________________________________________________________________

    
    
    cadenaAux = cadena
    
    Select Case conjCH
    Case -1:
        conjCaractValidos = conjCHs
    Case 0:
        conjCaractValidos = conjCH0
    Case 1:
        conjCaractValidos = conjCH1
    Case 2:
        conjCaractValidos = conjCH2
    Case 3:
        conjCaractValidos = conjCH3
    Case 4:
        conjCaractValidos = conjCH4
    Case Else:
        conjCaractValidos = conjCH0
    End Select



chInv = ""

For I = 1 To Len(cadena)
   
    ch = Mid(cadena, I, 1)
    
    ' si encuentra ch en conjunto de caracteres validos esta ok, sino eliminar el caracter
    If InStr(conjCaractValidos, ch) = 0 Or InStr(conjCaractValidos, ch) = -1 Then
        If Not (Asc(ch) = 186) And Not (conjCH = 1) Then ' 27/10/2014 validacion para el caracter �
            cadenaAux = Replace(cadenaAux, ch, "")
            chInv = chInv & " " & ch
        End If
    End If
 
Next

        
If chInv <> "" Then
    'Texto = "En la cadena: " & cadena & ", se eliminaron los siguientes caracteres: " & chInv
    Texto = Replace(EscribeLogMI("En @@TXT@@, se eliminaron los siguientes caracteres:@@TXT1@@"), "@@TXT@@", cadena)
    Texto = Replace(Texto, "@@TXT1@@", chInv)
    Call Escribir_Log("flogp", LineaCarga, NroColumna, Texto, Tabs, "")
End If

EliminarCHInvalidosII = cadenaAux



End Function


Public Sub LevantarParamteros(ByVal parametros As String)
Dim pos1 As Long
Dim pos2 As Long

Dim NombreArchivo1 As String
Dim NombreArchivo2 As String
Dim NombreArchivo3 As String


separador = "@"
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then

        'Nro de Modelo
        pos1 = 1
        pos2 = InStr(pos1, parametros, separador) - 1
        NroModelo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        'Nombre del archivo a levantar
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, separador) - 1
        If pos2 > 0 Then
            nombrearchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        Else
            pos2 = Len(parametros)
            nombrearchivo = Mid(parametros, pos1, pos2 - pos1 + 1)
        End If
        
        'Dependiendo del modelo puede que vengan mas parametros
        'Flog.Writeline Espacios(Tabulador * 2) & "Modelo nro " & NroModelo
        Select Case NroModelo
        Case 211: 'Interface de Novedades
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 212: 'GTI - Mega Alarmas
        Case 213: 'GTI - Acumulado Diario
        Case 214: 'Tickets
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            TikPedNro = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 215: 'Interface de Acumuladores de Agencia
        Case 216: 'Interface de Acumuladores de Agencia para Citrusvil
        Case 217: 'Interface de Vales
        Case 218: 'Libre
        Case 219: 'Libre
        Case 220: 'Libre
        Case 221: 'Libre
        Case 222: 'Libre
        Case 223: 'Libre
        Case 224: 'Libre
        Case 225: 'Libre
        Case 226: 'Interface de Postulantes Deloitte
            Call Cargar_datos_Estandar
        Case 227: 'Libre
        Case 228: 'Declaracion Jurada (LA ESTRELLA)
        Case 229: 'Interface de Prestamos
        Case 230: 'Interface de Pedidos de Vacaciones
        Case 231: 'Exportacion / Interface Banco Nacion
        Case 232: 'Interface Bumerang
        Case 233: 'Interface de Licencias
        Case 234: 'Exportacin JDE
        Case 235: 'Interface de Estadisticas de Accidentes
        Case 236: 'Interface de Bultos
        Case 239: 'Interfase Deloitte
            Call Cargar_datos_deloitte
        Case 241: 'Interface Dabra
        Case 242: 'Interface SAP
        Case 243: 'Interface Cuentas Bancarias
        
        Case 244: '
        Case 245: 'Interface de Nov Ajuste
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)

        Case 246: '
        Case 247: 'Interface de Acumulado de Horas TELEPERFORMANCE
        Case 248: 'INFOTIPOS DELOITTE
        Case 249: 'INFOTIPOS MAPEOS
        Case 250: 'Importacion de Acum Men
        Case 251: 'Exportaci�n Roche
        Case 252: 'Exportaci�n Indura
        Case 253: 'Importaci�n de Embargos
        Case 254: 'Exp. Excel Comp. Ac.
        Case 255: 'Exportaci�n BPS
        Case 256: 'Importacion de Engagements
        Case 257: 'Interface TTI (reservado)
        Case 258: 'Exportaci�n de SIJP (reservado)
        Case 259: 'Exportaci�n de SICORE (reservado)
        Case 260: 'Exportaci�n de Libro Ley (reservado)
        Case 261: 'Importacion de Empresas
        Case 262: 'Importaci�n Dias Correspondientes
        Case 263: 'Interfase de Postulantes CARSA
        Case 264: 'Reservado en otro proceso
        Case 265: 'Importacion de Documentos
        Case 266: 'Interface Wella Zonas de Ventas
        Case 267: 'Interface Wella Cab de Facturacion
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            'Pisa = Mid(parametros, pos1, pos2)
            
        Case 268: 'Migracion de Empleados
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 269:   'Exportacion Conceptos Liquidados
                    'reservado en otro proceso
        Case 270:   'Exportacion Direction
                    'reservado en otro proceso
        Case 271:   'Exportacion Word Meeting
                    'reservado en otro proceso
        Case 272:   'Interface Novedades SPEC
            'Pisa Novedades
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, separador) - 1
            PisaNovedad = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
                    
            'Tiene vigencia
            pos1 = pos2 + 2
            pos2 = InStr(pos1, parametros, separador) - 1
            If pos2 > 0 Then
                Vigencia = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
                
                If Vigencia Then
                    'Fecha desde
                    pos1 = pos2 + 2
                    pos2 = InStr(pos1, parametros, separador) - 1
                    If pos2 > 0 Then
                        Vigencia_Desde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
                
                        'Fecha Hasta
                        pos1 = pos2 + 2
                        pos2 = Len(parametros)
                        If pos2 > pos1 Then
                            Vigencia_Hasta = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
                        End If
                    Else
                        pos2 = Len(parametros)
                        Vigencia_Desde = CDate(Mid(parametros, pos1, pos2 - pos1 + 1))
                        
                        Vigencia_Hasta = ""
                    End If
                End If
            Else
                Vigencia = False
            End If
        Case 273: 'Reservado
        Case 274: 'reservado
        Case 275: 'Interface de Postulantes Estandar
            Call Cargar_datos_Estandar
        ' .....
        Case 279: 'Interface Acum. Diarios SMT
        Case 280: 'Interface bandas salariales
        Case 281: 'Alta masiva de Remuneraciones (LA CAJA)
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 282: ' Interface complemento de remuneraciones
        Case 286: 'Interface de Postulantes Estandar R3
            'Call Cargar_datos_Estandar
        Case 288:
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)
            
        Case 289: 'Interface de Nov Ajuste
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedadAju = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        Case 293: 'Interfaz de novedades
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)
            
        Case 294: 'Importacion Instituciones
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 295: 'Importacion carreras
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 300: 'Migracion de Empleados para TELEPERFORMANCE ( + 3 columnas)
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        '.....
        Case 303: 'Migracion de Empleados para Halliburton
        Case 304: 'Interfase de motivos Wella
        Case 305: 'Interfase de empleados de zona Wella
        Case 306: 'Interfase de empleados de zona motivo Wella
        Case 307: 'Interfase de importacion de resultados de eventos de evaluacion
        Case 311: 'Interfase de ??
        Case 312: 'Interfase de ??
        Case 313: 'Interfase de Novedades con seguridad
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            PisaNovedad = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 317: 'Interface Distribuci�n Contable - PRICE
            'pos1 = pos2 + 2
            'pos2 = Len(parametros)
            'PisarDistCont = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 318: 'Interface Valor Plan Obra Social - PRICE
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            PisaPlan = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 322: ' Interface dependencia de estructuras
        Case 337: 'Interface de Importacion Archivos de AsoSykes
            'PisaPlan = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 341: 'Interface de Novedades Con Firmas SYKES
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 346: 'Interface de Novedades Simulacion
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 347: 'Interface de Nov Ajuste Simulaci�n
        Case 351: 'Interfase de Novedades - Custom para PKF
                Dim fila As Integer
                For fila = 1 To 13
                    StrSql = "DELETE modelo_filas WHERE bpronro = " & NroProcesoBatch & "AND fila = " & fila
                    objConn.Execute StrSql, , adExecuteNoRecords
                Next fila
                'v 4.51 - 21/11/2012
                pos1 = pos2 + 2
                pos2 = Len(parametros)
                AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 352: 'Interfase de Comisiones Monto - CAS 16012
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pliqnro = Mid(parametros, pos1, pos2 - pos1 + 1)

        Case 353: 'Interfase de Comisiones Cantidad - CAS 16012
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pliqnro = Mid(parametros, pos1, pos2 - pos1 + 1)
            
        Case 354: 'Interface de Novedades - Controla Vigencia al remplazar.
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 394: 'Importacion de novedades horarias
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 396:
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            AccionNovedad = Mid(parametros, pos1, pos2 - pos1 + 1)
        Case 600: 'Migracion de Familiares
        Case 601: 'Familiares - Goyaike
        Case 605, 603, 611, 664: 'Migracion de Empleados - 611 migrac Emp de Agencia
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 606: 'Migracion de Empleados URU
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 607: 'Migracion de Empleados CHILE
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 608: 'Migracion de Empleados COLOMBIA
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 609: 'Migracion de Empleados 'Radiotronica
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 610: 'DesmenFamiliar
        Case 613: 'Migracion de Empleados CHILE - Custom INDAP
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 615: 'DDJJ
        Case 620: 'Desglose de Ganancias - Se hizo para Accor
        Case 625: 'Liquidaciones - Se hizo para Accor
        Case 630: 'Migracion de Historico de estructuras
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 635: 'Titulos
        Case 640: 'Fases
        Case 645  'Migracion de acumuladores mensuales
            'este modelo fue borrado. Se puede reutilizar
        Case 653  'Migracion de Reporta a (Estandar)
            '
        Case 661: ' Migraci�n de Im�genes de Empleados
            pos1 = pos2 + 2
            pos2 = Len(parametros)
            Pisa = CBool(Mid(parametros, pos1, pos2 - pos1 + 1))
        Case 668: ' Migraci�n de Domicilios de Per�, Chile, Colombia, Brasil , Uruguay, Venezuela, Ecuador, M�jico y Paraguay
            'Guardo el modelo de domicilio. modnro
            ModeloDom = CInt(Mid(parametros, pos2 + 2, Len(pos2)))
        Case 922: ' Migraci�n de Domicilios Multipais (Tipo y N� de DOC)
            'Guardo el modelo de domicilio. modnro
            ModeloDom = CInt(Mid(parametros, pos2 + 2, Len(pos2)))
        Case 1006: 'Organizaci�n Territorial Multipa�s
            'Guardo el modelo de domicilio. modnro
            ModeloDom = CInt(Mid(parametros, pos2 + 2, Len(pos2)))

        End Select
    End If
End If

End Sub



Public Sub ComenzarTransferencia()
Dim directorio As String
Dim CArchivos
Dim archivo
Dim Folder

    StrSql = "SELECT sis_direntradas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = Trim(objRs!sis_direntradas)
    Else
        'Flog.writeline Espacios(Tabulador * 1) & "No se encontr� el registro de la tabla sistema nro 1"
        Flog.writeline Espacios(Tabulador * 1) & EscribeLogMI("Debe ingresar un directorio.")
        
        Exit Sub
    End If
    
    
    StrSql = "SELECT * FROM modelo WHERE modnro = " & NroModelo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        directorio = directorio & Trim(objRs!modarchdefault)
        'Directorio = "\\rhdesa\Fuentes\4000_rhprox2_r4\In-Out"
        separador = IIf(Not IsNull(objRs!modseparador), objRs!modseparador, ",")
        SeparadorDecimal = IIf(Not IsNull(objRs!modsepdec), objRs!modsepdec, ".")
        UsaEncabezado = IIf(Not IsNull(objRs!modencab), CBool(objRs!modencab), False)
        DescripcionModelo = objRs!moddesc
        
        Flog.writeline Espacios(Tabulador * 1) & EscribeLogMI("Modelo") & " " & NroModelo & " " & objRs!moddesc
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & EscribeLogMI("Directorio de importaci�n") & " :  " & directorio
     Else
        Flog.writeline Espacios(Tabulador * 1) & EscribeLogMI("No se encontr� el modelo") & " " & NroModelo
        Exit Sub
    End If
    
    'Algunos modelos no se comportan de la misma manera ==>
    
    'FGZ - 19/01/2012 ------------------------------------
    'Clave_Lista_WC = 0
    
    'cargo nombre de wf
    Call CargarNombresTablasTemporales
    Call CreateTempTable(TTempWF_MOV_HORARIOS)
    'FGZ - 19/01/2012 ------------------------------------
    
    'FGZ - 19/06/2012 ------------------------------------
    Select Case NroModelo
    Case 352:   'interface comisiones - Monto
        Call Cargar_Confrep(372)
    Case 353:   'interface comisiones - Cantidad
        Call Cargar_Confrep(373)
    End Select
    'FGZ - 19/06/2012 ------------------------------------
    If HuboError Then
        Exit Sub
    End If
    
    Select Case NroModelo
'    Case 222:
'        Call LineaModelo_222
    Case 333:
        Call LineaModelo_333
        
    Case Else
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        Path = directorio
        
        Dim fc, F1, s2
        Set Folder = fs.GetFolder(directorio)
        Set CArchivos = Folder.Files
        
        HuboError = False
        EncontroAlguno = False
        For Each archivo In CArchivos
            EncontroAlguno = True
            If UCase(archivo.Name) = UCase(nombrearchivo) Then
                NArchivo = archivo.Name
                Flog.writeline Espacios(Tabulador * 1) & EscribeLogMI("Procesando archivo") & " " & archivo.Name
                Call LeeArchivo(directorio & "\" & archivo.Name)
            End If
        Next
        If Not EncontroAlguno Then
            Flog.writeline Espacios(Tabulador * 1) & EscribeLogMI("No se encontr� el archivo") & " " & nombrearchivo
        End If
    End Select
    'FGZ - 19/01/2012 ------------------------------------
    Call BorrarTempTable(TTempWF_MOV_HORARIOS)
    'FGZ - 19/01/2012 ------------------------------------
End Sub

Public Sub InsertaError(NroCampo As Byte, nroError As Long)
    StrSql = "INSERT INTO inter_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
             crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    Flog.writeline "error:" & StrSql
    objConn.Execute StrSql, , adExecuteNoRecords
    
    RegError = RegError + 1
    
    ErroresNov = True
End Sub



Public Sub Escribir_Log(ByVal TipoLog As String, ByVal Lin As Long, ByVal col As Long, ByVal msg As String, ByVal CantTab As Long, ByVal strLinea As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Escribe un mensage determinado en uno de 3 archivos de log
' Autor      : FGZ
' Fecha      : 18/04/2005
' Ultima Mod.:
' Descripcion: 13/02/2012 - Gonzalez Nicol�s - Se agreg� Multilenguaje
' ---------------------------------------------------------------------------------------------
Dim Texto
Texto = EscribeLogMI("Linea") & " " & Lin
Texto = Texto & " " & EscribeLogMI("Columna") & " " & col
Select Case UCase(TipoLog)

    Case "FLOG" 'Archivo de Informacion de resumen
            Flog.writeline Espacios(Tabulador * CantTab) & msg
    Case "FLOGE" 'Archivo de Errores
            'FlogE.writeline Espacios(Tabulador * CantTab) & "Linea " & Lin & " Columna " & col & ": " & msg
            FlogE.writeline Espacios(Tabulador * CantTab) & Texto & ": " & msg
            FlogE.writeline Espacios(Tabulador * CantTab) & strLinea
    Case "FLOGP" 'Archivo de lineas procesadas
            'FlogP.writeline Espacios(Tabulador * CantTab) & "Linea " & Lin & " Columna " & col & ": " & msg
            FlogP.writeline Espacios(Tabulador * CantTab) & Texto & ": " & msg
    Case Else
        'Flog.writeline Espacios(Tabulador * CantTab) & "Nombre de archivo de log incorrecto " & TipoLog
        Flog.writeline Espacios(Tabulador * CantTab) & Replace(EscribeLogMI("Nombre de archivo de log incorrecto"), "@@TXT@@", TipoLog)
End Select

End Sub

Public Function Firmas_novliq(ByVal iduser As String, ByVal cystipnro As Integer, ByVal ConcCod As String, ByVal lista_orden As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: 'Valida que el usuario sea fin de firma / tenga permisos delegados /
'               Tenga un complemento asociado al circuito de firmas
'               Devuelve cysfirusuario,cysfirautoriza,cysfirdestino,cysfirfin,cysfiryaaut,cysfirrecha - Para insertar en cysfirmas
'               Aplica solo para las novedades de liquidaci�n
' Autor      : Gonzalez Nicol�s
' Fecha      : 19/04/2011
' Ultima Mod.: 12/05/2011 - Gonzalez Nicol�s: Se setearon las variables en 0, para cuando hay error
' Ultima Mod.: 02/06/2011 - Gonzalez Nicol�s: Si no tiene complemento o no tiene el concepto en la lista elige un fin de firma
'
' ---------------------------------------------------------------------------------------------


Dim rs_firmas As New ADODB.Recordset
Dim rs_firmas2 As New ADODB.Recordset
Dim rs_firmas3 As New ADODB.Recordset

Dim Esfin
Dim cysfirusuario
Dim cysfirautoriza
Dim cysfirdestino
Dim cysfirfin
Dim cysfiryaaut
Dim cysfirrecha
Dim cyslfirmnro
Dim Tenro
Dim l_listperfnro
Dim listestrnro
Dim cyslfirmdetnro
Dim tipoorigen
Dim strLinea
Dim l_listperfnro_aux
Dim a
Dim StrSql_aux


'Seteo todo en 0
cysfirusuario = ""
cysfirautoriza = ""
cysfirdestino = ""
cysfirfin = 0
cysfiryaaut = 0
cysfirrecha = 0
l_listperfnro = ""
tipoorigen = ""
'=====================================
'FIN DE FIRMA
'=====================================
StrSql = "SELECT * FROM cysfincirc "
StrSql = StrSql & " WHERE userid = '" & iduser & "' and cystipnro = " & cystipnro
OpenRecordset StrSql, rs_firmas
If Not rs_firmas.EOF Then
    Esfin = True
    cysfirusuario = iduser
    cysfirautoriza = iduser
    cysfirdestino = ""
    cysfirfin = -1
    cysfiryaaut = -1
    cysfirrecha = 0
Else
    Esfin = False
End If
rs_firmas.Close

If Esfin = False Then
    '=====================================
    'QUE TENGA DELEGADO UN PERMISO
    '=====================================
    StrSql = "SELECT bk_cab.iduser, bkcystipnro "
    StrSql = StrSql & " From bk_cab "
    StrSql = StrSql & " INNER JOIN bk_firmas on bk_firmas.bkcabnro = bk_cab.bkcabnro "
    StrSql = StrSql & " Where fdesde <= " & ConvFecha(Date)
    StrSql = StrSql & " AND (fhasta >= " & ConvFecha(Date) & " OR fhasta IS NULL)"
    StrSql = StrSql & " AND bk_firmas.iduser = '" & iduser & "'"
    StrSql = StrSql & " AND bkcystipnro = " & cystipnro
    StrSql = StrSql & " AND bk_cab.iduser <> '" & iduser & "'"
    OpenRecordset StrSql, rs_firmas
    
    If Not rs_firmas.EOF Then
        Esfin = True
        cysfirusuario = rs_firmas!iduser
        cysfirautoriza = iduser
        cysfirdestino = ""
        
        cysfirfin = -1
        cysfiryaaut = -1
        cysfirrecha = 0
    Else
        Esfin = False
    End If
    rs_firmas.Close
End If
'-----

If Esfin = False Then
'Si no es fin de firma valida que tenga asociado un complemento a la lista
'y determina el primer usuario de la lista como siguiente en el circuito
    StrSql = "SELECT cyscompdet.cyslfirmnro,cyscomp.cyscomtipnro,cyslfirmantes_det.tipoorigen,cyslfirmantes_det.cyslfirmdetnro,orden,cyscomdetdesc "
    StrSql = StrSql & " FROM cyscomp "
    StrSql = StrSql & " INNER JOIN cyscompdet ON cyscompdet.cyscomnro = cyscomp.cyscomnro "
    StrSql = StrSql & " INNER JOIN cyslfirmantes_det ON cyslfirmantes_det.cyslfirmnro = cyscompdet.cyslfirmnro "
    StrSql = StrSql & " Where cystipnro = " & cystipnro
    StrSql = StrSql & " AND cyscomdetdesc = " & ConcCod
    StrSql = StrSql & " AND orden = " & lista_orden
    OpenRecordset StrSql, rs_firmas

    If Not rs_firmas.EOF Then
        tipoorigen = rs_firmas!tipoorigen
        cyslfirmnro = rs_firmas!cyslfirmnro
        cyslfirmdetnro = rs_firmas!cyslfirmdetnro
    Else
        StrSql = "SELECT userid FROM cysfincirc"
        StrSql = StrSql & " WHERE userid <> '" & iduser & "' and cystipnro =" & cystipnro
        OpenRecordset StrSql, rs_firmas2
        If Not rs_firmas2.EOF Then
            Esfin = True
            cysfirusuario = rs_firmas2!userid
            cysfirautoriza = iduser
            cysfirdestino = ""
        
            cysfirfin = -1
            cysfiryaaut = -1
            cysfirrecha = 0
      
        Else
            Firmas_novliq = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
            Texto = ": " & "No existen usuarios fin de firma"
            NroColumna = 3
            Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
            Call InsertaError(3, 119)
            HuboError = True
            Exit Function
        End If
        rs_firmas2.Close
    End If
    rs_firmas.Close
    
    If tipoorigen <> "" Then
    'Selecciona el primer usuario dependiendo del tipo de lista
    Select Case tipoorigen
                Case 1: 'Perfiles----------------------------------------
                    StrSql = "SELECT detorigen FROM cyslfirmantes_det"
                    StrSql = StrSql & " WHERE cyslfirmnro = " & cyslfirmnro
                    OpenRecordset StrSql, rs_firmas2
                    If Not rs_firmas2.EOF And rs_firmas2!detorigen = 0 Then
                        'SI SON TODOS LOS PERFILES
                        StrSql = "SELECT perfnro, perfnom  FROM perf_usr"
                        OpenRecordset StrSql, rs_firmas3
                                              
                        If Not rs_firmas3.EOF Then
                            Do While Not rs_firmas3.EOF
                              l_listperfnro_aux = l_listperfnro_aux & "," & rs_firmas3!perfnro
                              rs_firmas3.MoveNext
                            Loop
                        Else
                            Firmas_novliq = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                            Texto = ": " & "No hay existe ning�n tipo de perfil"
                            NroColumna = 3
                            Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                            Call InsertaError(3, 120)
                            HuboError = True
                            Exit Function
                        End If
                    End If
                    '------SI NO SON TODOS
                    If l_listperfnro_aux = "" Then
                        StrSql = "SELECT cyslfirmantes_det_perf.listperfnro "
                        StrSql = StrSql & " FROM cyslfirmantes_det "
                        StrSql = StrSql & " INNER JOIN cyslfirmantes_det_perf ON cyslfirmantes_det_perf.cyslfirmdetperfnro = cyslfirmantes_det.detorigen "
                        StrSql = StrSql & " WHERE cyslfirmantes_det.cyslfirmnro = " & cyslfirmnro
                        StrSql = StrSql & " AND cyslfirmantes_det.orden = " & lista_orden
                        OpenRecordset StrSql, rs_firmas
                        If Not rs_firmas.EOF Then
                            l_listperfnro_aux = rs_firmas!listperfnro
                        Else
                            l_listperfnro_aux = ""
                        End If
                        rs_firmas.Close
                    End If
                    
                    If l_listperfnro_aux <> "" Then
                        l_listperfnro = Split(l_listperfnro_aux, ",")
                        '----Crea la lista de perfiles
                        StrSql = "SELECT iduser,listperfnro "
                        StrSql = StrSql & " FROM  user_perfil "
                        StrSql = StrSql & " WHERE (',' + listperfnro + ',' like '%," & l_listperfnro(0) & ",%'"
                        If UBound(l_listperfnro) > 0 Then
                            For a = 1 To UBound(l_listperfnro)
                                StrSql_aux = StrSql_aux & " OR ',' + listperfnro + ',' like '%," & l_listperfnro(a) & ",%' "
                            Next
                        End If
                        StrSql = StrSql & StrSql_aux
                        StrSql = StrSql & ") AND iduser <> '" & iduser & "'"
                        OpenRecordset StrSql, rs_firmas
            
                        If Not rs_firmas.EOF Then
                        'Guarda el primer usuario de la lista
                            cysfirusuario = iduser
                            cysfirautoriza = iduser
                            cysfirdestino = rs_firmas!iduser
                            
                            cysfirfin = 0
                            cysfiryaaut = 0
                            cysfirrecha = 0
                        Else
                            Firmas_novliq = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                            Texto = ": " & "No se encuentran usuarios con perfil "
                            NroColumna = 3
                            Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                            Call InsertaError(3, 120)
                            HuboError = True
                            Exit Function
                        
                        End If
                        rs_firmas.Close
                    End If
                 Case 2: 'Usuarios ---------------------------------------------
                    StrSql = "SELECT detorigen FROM cyslfirmantes_det"
                    StrSql = StrSql & " WHERE cyslfirmnro = " & cyslfirmnro
                    OpenRecordset StrSql, rs_firmas2
                    'Si son todos los usuarios
                    If Not rs_firmas2.EOF And rs_firmas2!detorigen = 0 Then
                        StrSql = "SELECT iduser FROM user_ter"
                        StrSql = StrSql & " WHERE iduser <> '" & iduser & "'"
                        OpenRecordset StrSql, rs_firmas3
                        If Not rs_firmas3.EOF Then
                              cysfirusuario = iduser
                              cysfirautoriza = iduser
                              cysfirdestino = rs_firmas3!iduser
                              cysfirfin = 0
                              cysfiryaaut = 0
                              cysfirrecha = 0
                        Else
                            Firmas_novliq = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                            Texto = ": " & "No hay usuarios de sistema existentes"
                            NroColumna = 3
                            Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                            Call InsertaError(3, 120)
                            HuboError = True
                            Exit Function
                        
                        End If
                        rs_firmas3.Close
                    Else
                        'Si NO son todos los usuarios
                         StrSql = "SELECT iduser FROM cyslfirmantes "
                         StrSql = StrSql & " INNER JOIN cyslfirmantes_det ON cyslfirmantes_det.cyslfirmnro = cyslfirmantes.cyslfirmnro "
                         StrSql = StrSql & " INNER JOIN cyslfirmantes_det_usr ON cyslfirmantes_det_usr.cyslfirmdetnro = cyslfirmantes_det.cyslfirmdetnro "
                         StrSql = StrSql & " WHERE cyslfirmantes.cyslfirmnro = " & cyslfirmnro
                         StrSql = StrSql & " AND orden = " & lista_orden
                         StrSql = StrSql & " AND iduser <> '" & iduser & "'"
                         OpenRecordset StrSql, rs_firmas
                         If Not rs_firmas.EOF Then
                             'Guarda el primer usuario de la lista
                              cysfirusuario = iduser
                              cysfirautoriza = iduser
                              cysfirdestino = rs_firmas!iduser
                                 
                              cysfirfin = 0
                              cysfiryaaut = 0
                              cysfirrecha = 0
            
                         Else
                             Firmas_novliq = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                             Texto = ": " & "No se encuentran usuarios en la lista "
                             NroColumna = 3
                             Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                             Call InsertaError(3, 121)
                             HuboError = True
                             Exit Function
                         End If
                         rs_firmas.Close
   
                    End If
                    rs_firmas2.Close
                Case 3: 'Reporta A: -------------------------------------------------------------
                    StrSql = "SELECT urepo.iduser "
                    StrSql = StrSql & " FROM user_ter "
                    StrSql = StrSql & " INNER JOIN v_empleado ON v_empleado.ternro = user_ter.ternro "
                    StrSql = StrSql & " AND user_ter.iduser = '" & iduser & "'"
                    StrSql = StrSql & " AND v_empleado.empest= -1 "
                    StrSql = StrSql & " INNER JOIN v_empleado repo ON repo.ternro = v_empleado.empreporta "
                    StrSql = StrSql & " AND repo.empest = -1"
                    StrSql = StrSql & " INNER JOIN user_ter urepo ON urepo.ternro = repo.ternro "
                    StrSql = StrSql & " AND urepo.iduser <> '" & iduser & "'"
                    OpenRecordset StrSql, rs_firmas
                    
                    If Not rs_firmas.EOF Then
                        'Guarda el primer usuario de la lista
                         cysfirusuario = iduser
                         cysfirautoriza = iduser
                         cysfirdestino = rs_firmas!iduser
                            
                         cysfirfin = 0
                         cysfiryaaut = 0
                         cysfirrecha = 0
       
                    Else
                        Firmas_novliq = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                        Texto = ": " & "No se encuentran usuarios Reporta a: en la lista "
                        NroColumna = 3
                        Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                        Call InsertaError(3, 122)
                        HuboError = True
                        Exit Function
                    End If
                    rs_firmas.Close
                Case 4: 'Busco lista de estructuras --------------------------------
                    
                    StrSql = "SELECT listestrnro,tenro FROM cyslfirmantes_det_estr "
                    StrSql = StrSql & " WHERE cyslfirmdetnro = " & cyslfirmdetnro
                    OpenRecordset StrSql, rs_firmas
                    If Not rs_firmas.EOF Then
                        listestrnro = rs_firmas!listestrnro
                        Tenro = rs_firmas!Tenro
                    End If
                    rs_firmas.Close
                    
                    'Busco lista de usuarios
                    StrSql = "SELECT user_ter.iduser FROM v_empleado "
                    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = v_empleado.ternro "
                    StrSql = StrSql & " AND his_estructura.tenro = " & Tenro
                    StrSql = StrSql & " AND (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(Date) & ") "
                    StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Date) & ""
                    StrSql = StrSql & " AND his_estructura.estrnro IN (" & listestrnro & ") AND v_empleado.empest= -1"
                    StrSql = StrSql & " INNER JOIN user_ter ON user_ter.ternro = v_empleado.ternro"
                    StrSql = StrSql & " AND user_ter.iduser <> '" & iduser & "'"
                    OpenRecordset StrSql, rs_firmas
                    If Not rs_firmas.EOF Then
                        'Guarda el primer usuario de la lista
                         cysfirusuario = iduser
                         cysfirautoriza = iduser
                         cysfirdestino = rs_firmas!iduser
                            
                         cysfirfin = 0
                         cysfiryaaut = 0
                         cysfirrecha = 0
                     Else
                        Firmas_novliq = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                        Texto = ": " & "No se encuentran usuarios por Estructura en la lista "
                        NroColumna = 3
                        Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                        Call InsertaError(3, 123)
                        HuboError = True
                        Exit Function
                    End If
                    rs_firmas.Close
        
            End Select
    
    End If 'Cierra tipoorigen <> ""
    

End If
Firmas_novliq = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha

If rs_firmas.State = adStateOpen Then rs_firmas.Close
Set rs_firmas = Nothing
'rs_firmas.Close


End Function
Public Function Firmas_lista(ByVal iduser As String, ByVal cystipnro As Integer, ByVal lista_orden As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: 'Valida que el usuario sea fin de firma / tenga permisos delegados /
'               Tenga un complemento asociado al circuito de firmas
'               Devuelve cysfirusuario,cysfirautoriza,cysfirdestino,cysfirfin,cysfiryaaut,cysfirrecha - Para insertar en cysfirmas
'               Aplica solo para las novedades de liquidaci�n
' Autor      : Gonzalez Nicol�s
' Fecha      : 01/06/2011
' Ultima Mod.:
'
' ---------------------------------------------------------------------------------------------


Dim rs_firmas As New ADODB.Recordset
Dim rs_firmas2 As New ADODB.Recordset
Dim rs_firmas3 As New ADODB.Recordset

Dim Esfin
Dim cysfirusuario
Dim cysfirautoriza
Dim cysfirdestino
Dim cysfirfin
Dim cysfiryaaut
Dim cysfirrecha
Dim cyslfirmnro
Dim Tenro
Dim l_listperfnro
Dim listestrnro
Dim cyslfirmdetnro
Dim tipoorigen
Dim strLinea
Dim l_listperfnro_aux
Dim a
Dim StrSql_aux


'Seteo todo en 0
cysfirusuario = ""
cysfirautoriza = ""
cysfirdestino = ""
cysfirfin = 0
cysfiryaaut = 0
cysfirrecha = 0
l_listperfnro_aux = ""

'=====================================
'FIN DE FIRMA
'=====================================
StrSql = "SELECT * FROM cysfincirc "
StrSql = StrSql & " WHERE userid = '" & iduser & "' and cystipnro = " & cystipnro
OpenRecordset StrSql, rs_firmas
If Not rs_firmas.EOF Then
    Esfin = True
    cysfirusuario = iduser
    cysfirautoriza = iduser
    cysfirdestino = ""
    cysfirfin = -1
    cysfiryaaut = -1
    cysfirrecha = 0
Else
    Esfin = False
End If
rs_firmas.Close

If Esfin = False Then
    '=====================================
    'QUE TENGA DELEGADO UN PERMISO
    '=====================================
    StrSql = "SELECT bk_cab.iduser, bkcystipnro "
    StrSql = StrSql & " From bk_cab "
    StrSql = StrSql & " INNER JOIN bk_firmas on bk_firmas.bkcabnro = bk_cab.bkcabnro "
    StrSql = StrSql & " Where fdesde <= " & ConvFecha(Date)
    StrSql = StrSql & " AND (fhasta >= " & ConvFecha(Date) & " OR fhasta IS NULL)"
    StrSql = StrSql & " AND bk_firmas.iduser = '" & iduser & "'"
    StrSql = StrSql & " AND bkcystipnro = " & cystipnro
    StrSql = StrSql & " AND bk_cab.iduser <> '" & iduser & "'"
    OpenRecordset StrSql, rs_firmas
    
    If Not rs_firmas.EOF Then
        Esfin = True
        cysfirusuario = rs_firmas!iduser
        cysfirautoriza = iduser
        cysfirdestino = ""
        
        cysfirfin = -1
        cysfiryaaut = -1
        cysfirrecha = 0
    Else
        Esfin = False
    End If
    rs_firmas.Close
End If
'-----

If Esfin = False Then
'Si no es fin de firma valida que tenga asociado un complemento a la lista
'y determina el primer usuario de la lista como siguiente en el circuito
    StrSql = "SELECT cyscompdet.cyslfirmnro,cyscomp.cyscomtipnro,cyslfirmantes_det.tipoorigen,cyslfirmantes_det.cyslfirmdetnro,orden,cyscomdetdesc "
    'StrSql = "SELECT cyscomp.cyscomtipnro,cyslfirmantes_det.tipoorigen,cyslfirmantes_det.cyslfirmdetnro,orden "
    StrSql = StrSql & " FROM cyscomp "
    StrSql = StrSql & " INNER JOIN cyscompdet ON cyscompdet.cyscomnro = cyscomp.cyscomnro "
    StrSql = StrSql & " INNER JOIN cyslfirmantes_det ON cyslfirmantes_det.cyslfirmnro = cyscompdet.cyslfirmnro "
    StrSql = StrSql & " Where cystipnro = " & cystipnro
    StrSql = StrSql & " AND orden = " & lista_orden
    OpenRecordset StrSql, rs_firmas

    If Not rs_firmas.EOF Then
        tipoorigen = rs_firmas!tipoorigen
        cyslfirmnro = rs_firmas!cyslfirmnro
        cyslfirmdetnro = rs_firmas!cyslfirmdetnro
    Else
    'SI NO TIENE COMPLEMENTO ASOCIADO BUSCO LISTA DEL CIRCUITO
        StrSql = "SELECT cy_det.tipoorigen,cy_det.cyslfirmdetnro,cy_det.cyslfirmnro FROM cyslfirmantes_det cy_det"
        StrSql = StrSql & " INNER join cystipo on cystipo.cyslfirmnro = cy_det.cyslfirmnro"
        StrSql = StrSql & " WHERE cystipo.cystipnro = " & cystipnro
        OpenRecordset StrSql, rs_firmas2
        If Not rs_firmas2.EOF Then
            tipoorigen = rs_firmas2!tipoorigen
            cyslfirmnro = rs_firmas2!cyslfirmnro
            cyslfirmdetnro = rs_firmas2!cyslfirmdetnro
        Else
            Firmas_lista = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
            Texto = ": " & "No Existe ninguna lista asociada al c�digo del circtuito " & cystipnro
            NroColumna = 3
            Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
            Call InsertaError(3, 119)
            HuboError = True
            Exit Function
        End If
    End If
    rs_firmas.Close
    rs_firmas2.Close
    
    If tipoorigen <> "" Then
    'Selecciona el primer usuario dependiendo del tipo de lista
    Select Case tipoorigen
                Case 1: 'Perfiles----------------------------------------
                    StrSql = "SELECT detorigen FROM cyslfirmantes_det"
                    StrSql = StrSql & " WHERE cyslfirmnro = " & cyslfirmnro
                    OpenRecordset StrSql, rs_firmas2
                    If Not rs_firmas2.EOF And rs_firmas2!detorigen = 0 Then
                        'SI SON TODOS LOS PERFILES
                        StrSql = "SELECT perfnro, perfnom  FROM perf_usr"
                        OpenRecordset StrSql, rs_firmas3
                                              
                        If Not rs_firmas3.EOF Then
                            Do While Not rs_firmas3.EOF
                              l_listperfnro_aux = l_listperfnro_aux & "," & rs_firmas3!perfnro
                              rs_firmas3.MoveNext
                            Loop
                        Else
                            Firmas_lista = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                            Texto = ": " & "No hay existe ning�n tipo de perfil"
                            NroColumna = 3
                            Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                            Call InsertaError(3, 120)
                            HuboError = True
                            Exit Function
                        End If
                    End If
                    '------SI NO SON TODOS
                    If l_listperfnro_aux = "" Then
                        StrSql = "SELECT cyslfirmantes_det_perf.listperfnro "
                        StrSql = StrSql & " FROM cyslfirmantes_det "
                        StrSql = StrSql & " INNER JOIN cyslfirmantes_det_perf ON cyslfirmantes_det_perf.cyslfirmdetperfnro = cyslfirmantes_det.detorigen "
                        StrSql = StrSql & " WHERE cyslfirmantes_det.cyslfirmnro = " & cyslfirmnro
                        StrSql = StrSql & " AND cyslfirmantes_det.orden = " & lista_orden
                        OpenRecordset StrSql, rs_firmas
                        If Not rs_firmas.EOF Then
                            l_listperfnro_aux = rs_firmas!listperfnro
                        Else
                            l_listperfnro_aux = ""
                        End If
                        rs_firmas.Close
                    End If
                    
                    If l_listperfnro_aux <> "" Then
                        l_listperfnro = Split(l_listperfnro_aux, ",")
                        '----Crea la lista de perfiles
                        StrSql = "SELECT iduser,listperfnro "
                        StrSql = StrSql & " FROM  user_perfil "
                        StrSql = StrSql & " WHERE (',' + listperfnro + ',' like '%," & l_listperfnro(0) & ",%'"
                        If UBound(l_listperfnro) > 0 Then
                            For a = 1 To UBound(l_listperfnro)
                                StrSql_aux = StrSql_aux & " OR ',' + listperfnro + ',' like '%," & l_listperfnro(a) & ",%' "
                            Next
                        End If
                        StrSql = StrSql & StrSql_aux
                        StrSql = StrSql & ") AND iduser <> '" & iduser & "'"
                        OpenRecordset StrSql, rs_firmas
            
                        If Not rs_firmas.EOF Then
                        'Guarda el primer usuario de la lista
                            cysfirusuario = iduser
                            cysfirautoriza = iduser
                            cysfirdestino = rs_firmas!iduser
                            
                            cysfirfin = 0
                            cysfiryaaut = 0
                            cysfirrecha = 0
                        Else
                            Firmas_lista = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                            Texto = ": " & "No se encuentran usuarios con perfil "
                            NroColumna = 3
                            Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                            Call InsertaError(3, 120)
                            HuboError = True
                            Exit Function
                        
                        End If
                        rs_firmas.Close
                    End If
                 Case 2: 'Usuarios ---------------------------------------------
                    StrSql = "SELECT detorigen FROM cyslfirmantes_det"
                    StrSql = StrSql & " WHERE cyslfirmnro = " & cyslfirmnro
                    OpenRecordset StrSql, rs_firmas2
                    'Si son todos los usuarios
                    If Not rs_firmas2.EOF And rs_firmas2!detorigen = 0 Then
                        StrSql = "SELECT iduser FROM user_ter"
                        StrSql = StrSql & " WHERE iduser <> '" & iduser & "'"
                        OpenRecordset StrSql, rs_firmas3
                        If Not rs_firmas3.EOF Then
                              cysfirusuario = iduser
                              cysfirautoriza = iduser
                              cysfirdestino = rs_firmas3!iduser
                              cysfirfin = 0
                              cysfiryaaut = 0
                              cysfirrecha = 0
                        Else
                            Firmas_lista = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                            Texto = ": " & "No hay usuarios de sistema existentes"
                            NroColumna = 3
                            Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                            Call InsertaError(3, 120)
                            HuboError = True
                            Exit Function
                        
                        End If
                        rs_firmas3.Close
                    Else
                        'Si NO son todos los usuarios
                         StrSql = "SELECT iduser FROM cyslfirmantes "
                         StrSql = StrSql & " INNER JOIN cyslfirmantes_det ON cyslfirmantes_det.cyslfirmnro = cyslfirmantes.cyslfirmnro "
                         StrSql = StrSql & " INNER JOIN cyslfirmantes_det_usr ON cyslfirmantes_det_usr.cyslfirmdetnro = cyslfirmantes_det.cyslfirmdetnro "
                         StrSql = StrSql & " WHERE cyslfirmantes.cyslfirmnro = " & cyslfirmnro
                         StrSql = StrSql & " AND orden = " & lista_orden
                         StrSql = StrSql & " AND iduser <> '" & iduser & "'"
                         OpenRecordset StrSql, rs_firmas
                         If Not rs_firmas.EOF Then
                             'Guarda el primer usuario de la lista
                              cysfirusuario = iduser
                              cysfirautoriza = iduser
                              cysfirdestino = rs_firmas!iduser
                                 
                              cysfirfin = 0
                              cysfiryaaut = 0
                              cysfirrecha = 0
            
                         Else
                             Firmas_lista = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                             Texto = ": " & "No se encuentran usuarios en la lista "
                             NroColumna = 3
                             Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                             Call InsertaError(3, 121)
                             HuboError = True
                             Exit Function
                         End If
                         rs_firmas.Close
   
                    End If
                    rs_firmas2.Close
                   

                Case 3: 'Reporta A: -------------------------------------------------------------
                    StrSql = "SELECT urepo.iduser "
                    StrSql = StrSql & " FROM user_ter "
                    StrSql = StrSql & " INNER JOIN v_empleado ON v_empleado.ternro = user_ter.ternro "
                    StrSql = StrSql & " AND user_ter.iduser = '" & iduser & "'"
                    StrSql = StrSql & " AND v_empleado.empest= -1 "
                    StrSql = StrSql & " INNER JOIN v_empleado repo ON repo.ternro = v_empleado.empreporta "
                    StrSql = StrSql & " AND repo.empest = -1"
                    StrSql = StrSql & " INNER JOIN user_ter urepo ON urepo.ternro = repo.ternro "
                    StrSql = StrSql & " AND urepo.iduser <> '" & iduser & "'"
                    OpenRecordset StrSql, rs_firmas
                    
                    If Not rs_firmas.EOF Then
                        'Guarda el primer usuario de la lista
                         cysfirusuario = iduser
                         cysfirautoriza = iduser
                         cysfirdestino = rs_firmas!iduser
                            
                         cysfirfin = 0
                         cysfiryaaut = 0
                         cysfirrecha = 0
       
                    Else
                        Firmas_lista = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                        Texto = ": " & "No se encuentran usuarios Reporta a: en la lista "
                        NroColumna = 3
                        Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                        Call InsertaError(3, 122)
                        HuboError = True
                        Exit Function
                    End If
                    rs_firmas.Close
                Case 4: 'Busco lista de estructuras --------------------------------
                    
                    StrSql = "SELECT listestrnro,tenro FROM cyslfirmantes_det_estr "
                    StrSql = StrSql & " WHERE cyslfirmdetnro = " & cyslfirmdetnro
                    OpenRecordset StrSql, rs_firmas
                    If Not rs_firmas.EOF Then
                        listestrnro = rs_firmas!listestrnro
                        Tenro = rs_firmas!Tenro
                    End If
                    rs_firmas.Close
                    
                    'Busco lista de usuarios
                    StrSql = "SELECT user_ter.iduser FROM v_empleado "
                    StrSql = StrSql & " INNER JOIN his_estructura ON his_estructura.ternro = v_empleado.ternro "
                    StrSql = StrSql & " AND his_estructura.tenro = " & Tenro
                    StrSql = StrSql & " AND (his_estructura.htethasta IS NULL OR his_estructura.htethasta >= " & ConvFecha(Date) & ") "
                    StrSql = StrSql & " AND his_estructura.htetdesde <= " & ConvFecha(Date) & ""
                    StrSql = StrSql & " AND his_estructura.estrnro IN (" & listestrnro & ") AND v_empleado.empest= -1"
                    StrSql = StrSql & " INNER JOIN user_ter ON user_ter.ternro = v_empleado.ternro"
                    StrSql = StrSql & " AND user_ter.iduser <> '" & iduser & "'"
                    OpenRecordset StrSql, rs_firmas
                    If Not rs_firmas.EOF Then
                        'Guarda el primer usuario de la lista
                         cysfirusuario = iduser
                         cysfirautoriza = iduser
                         cysfirdestino = rs_firmas!iduser
                            
                         cysfirfin = 0
                         cysfiryaaut = 0
                         cysfirrecha = 0
                     Else
                        Firmas_lista = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha
                        Texto = ": " & "No se encuentran usuarios por Estructura en la lista "
                        NroColumna = 3
                        Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, strLinea)
                        Call InsertaError(3, 123)
                        HuboError = True
                        Exit Function
                    End If
                    rs_firmas.Close
        
            End Select
    
    End If 'Cierra tipoorigen <> ""
    

End If
Firmas_lista = cysfirusuario & "," & cysfirautoriza & "," & cysfirdestino & "," & cysfirfin & "," & cysfiryaaut & "," & cysfirrecha

If rs_firmas.State = adStateOpen Then rs_firmas.Close
Set rs_firmas = Nothing
'rs_firmas.Close


End Function



Public Sub Cargar_Confrep(ByVal Reporte As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga un arreglo de configuracion para posteriormente ser utilizado por el modelo que corresponda.
'              Ante cualquier Dato mal cargado se para la ejecucion.
' Autor      : FGZ
' Fecha      : 19/06/2012
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim rs_Confrep As New ADODB.Recordset
Dim rs_con_for_tpa As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_TipoPar As New ADODB.Recordset

Dim Max_Cols As Long
Dim fornro As Long
Dim Encontro As Boolean

    'Busco las columnas configuradas en el Confrep, y dependiendo de la columna que sea inserto las novedades
    StrSql = "SELECT confnrocol, conftipo, confval, confval2  FROM confrep "
    StrSql = StrSql & " WHERE repnro = " & Reporte & " ORDER BY confnrocol desc "
    OpenRecordset StrSql, rs_Confrep
    
    If Not rs_Confrep.EOF Then
        Max_Cols = rs_Confrep!confnrocol + 1
    Else
        Max_Cols = 0
    End If
    ReDim Preserve ArrConfrep(Max_Cols) As Tconfrep

    Do While Not rs_Confrep.EOF
    

        
        Select Case UCase(rs_Confrep!conftipo)
            Case "DOC" 'Cargo el tipo de documento
                TDoc = rs_Confrep!Confval
            
            Case "TE" 'Cargo el tipo de estructura revendedor
                TERevendedor = rs_Confrep!Confval
            
            Case "NOV" 'Cargo las columnas de las novedades
            
                ArrConfrep(rs_Confrep!confnrocol).columna = rs_Confrep!confnrocol
                ArrConfrep(rs_Confrep!confnrocol).ConcCod = IIf(EsNulo(rs_Confrep!confval2), "", rs_Confrep!confval2)
                ArrConfrep(rs_Confrep!confnrocol).tpanro = rs_Confrep!Confval
                
                'Que exista el concepto
                StrSql = "SELECT concnro, fornro FROM concepto WHERE conccod = '" & ArrConfrep(rs_Confrep!confnrocol).ConcCod & "'"
                OpenRecordset StrSql, rs_Concepto
                If rs_Concepto.EOF Then
                    Texto = "No se encontro el Concepto: " & ArrConfrep(rs_Confrep!confnrocol).ConcCod & " configurado en Confrep "
                    NroColumna = 2
                    Call Escribir_Log("flog", NroLinea, NroColumna, Texto, 1, "")
                    HuboError = True
                    Exit Sub
        
                Else
                    ArrConfrep(rs_Confrep!confnrocol).ConcNro = rs_Concepto!ConcNro
                    fornro = rs_Concepto!fornro
                End If
                
                'Que exista el tipo de Parametro
                StrSql = "SELECT * FROM tipopar WHERE tpanro = " & ArrConfrep(rs_Confrep!confnrocol).tpanro
                OpenRecordset StrSql, rs_TipoPar
                
                If rs_TipoPar.EOF Then
                    Texto = "No se encontro el Tipo de Parametro " & ArrConfrep(rs_Confrep!confnrocol).tpanro
                    NroColumna = 3
                    Call Escribir_Log("flog", NroLinea, NroColumna, Texto, 1, "")
                    HuboError = True
                    Exit Sub
        
                End If
        
                'Que exista el par concepto-parametro y se resuelva por novedad
                StrSql = "SELECT * FROM con_for_tpa "
                StrSql = StrSql & " WHERE concnro = " & ArrConfrep(rs_Confrep!confnrocol).ConcNro
                StrSql = StrSql & " AND fornro = " & fornro
                StrSql = StrSql & " AND tpanro = " & ArrConfrep(rs_Confrep!confnrocol).tpanro
                'StrSql = StrSql & " AND cftauto = 0 "
                OpenRecordset StrSql, rs_con_for_tpa
                
                If rs_con_for_tpa.EOF Then
                    Texto = "El parametro " & ArrConfrep(rs_Confrep!confnrocol).tpanro & " no esta asociado a la formula del concepto " & ArrConfrep(rs_Confrep!confnrocol).ConcCod
                    NroColumna = 4
                    Call Escribir_Log("flog", NroLinea, NroColumna, Texto, 1, "")
                    HuboError = True
                    Exit Sub
                Else
                    Encontro = False
                    
                    Do While Not Encontro And Not rs_con_for_tpa.EOF
                        If Not CBool(rs_con_for_tpa!cftauto) Then
                            Encontro = True
                        End If
                        rs_con_for_tpa.MoveNext
                    Loop
                    
                      If Not Encontro Then
                        Texto = ": " & "El parametro " & ArrConfrep(rs_Confrep!confnrocol).tpanro & " del concepto " & ArrConfrep(rs_Confrep!confnrocol).ConcCod & " no se resuelve por novedad "
                        NroColumna = 3
                        Call Escribir_Log("floge", NroLinea, NroColumna, Texto, Tabs, "")
                        HuboError = True
                        Exit Sub
                    End If
                    
                End If
            
        End Select
        
        rs_Confrep.MoveNext
    Loop

'Cierro y libero
If rs_Confrep.State = adStateOpen Then rs_Confrep.Close
If rs_con_for_tpa.State = adStateOpen Then rs_con_for_tpa.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_TipoPar.State = adStateOpen Then rs_TipoPar.Close

Set rs_Confrep = Nothing
Set rs_con_for_tpa = Nothing
Set rs_Concepto = Nothing
Set rs_TipoPar = Nothing

End Sub
