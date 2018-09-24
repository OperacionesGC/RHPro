Attribute VB_Name = "mdlLeerRegistraciones"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "29/09/2005"
'Global Const UltimaModificacion = " " 'Customizacion Formato 7 SCHERING.

'Global Const Version = "1.02"
'Global Const FechaModificacion = "19/12/2005"
'Global Const UltimaModificacion = " " 'Procesamiento ON-LINE y + Log en todos los formatos

'Global Const Version = "1.03"
'Global Const FechaModificacion = "12/09/2006"
'Global Const UltimaModificacion = " " 'Se agrego InsertaFormato9

'Global Const Version = "1.04"
'Global Const FechaModificacion = "18/09/2006"
'Global Const UltimaModificacion = " " 'Se agrego InsertaFormato10 del modelo 200 para Sidersa

'Global Const Version = "1.05"
'Global Const FechaModificacion = "25/09/2006"
'Global Const UltimaModificacion = " " 'Se corrijio un error en InsertaFormato10 del modelo 200 para Sidersa. El legajo va desde la posicion 21.

'Global Const Version = "1.06"
'Global Const FechaModificacion = "04/10/2006"
'Global Const UltimaModificacion = " " 'formato InsertaFormato9, se le sacó los espacios en blanco en el codigo externo del rejoj
''                                                               se agregó las lineas para el procesamiento ON-LINE

'Global Const Version = "1.07"
'Global Const FechaModificacion = "06/03/2007"
'Global Const UltimaModificacion = " " 'agregado de log formato 1

'Global Const Version = "1.08"
'Global Const FechaModificacion = "20/03/2007" 'G. Bauer - N. Trillo
'Global Const UltimaModificacion = " " 'Se modifico el proceso para que el procesamiento online para que se pueda utilizar en para turnos nocturnos,
''                                      se procesa el dia anterior y el actual del HC y AD

'Global Const Version = "1.09"
'Global Const FechaModificacion = "10/04/2007" 'FGZ
'Global Const UltimaModificacion = " " 'Se modifico el sub InsertarWF_Lecturas. Estaba mal la definicion del parametro formal ternro

'Global Const Version = "1.10"
'Global Const FechaModificacion = "20/04/2007" 'FGZ
'Global Const UltimaModificacion = " " 'Se agregó un var para el usuario del proceso, ademas se usa para cuando crear los procesos del proc ONLINE porque siempre ponia super

'Global Const Version = "1.11"
'Global Const FechaModificacion = "18/05/2007" 'FAF
'Global Const UltimaModificacion = " " 'Se agregó el formato InsertaFormato11 del modelo 199 para Tabacal (Legajo(empleg) Fecha Hora Reloj Descarte) Ej:00320 22/11/02 09:23 03 20

'Global Const Version = "1.12"
'Global Const FechaModificacion = "21/05/2007" 'FAF
'Global Const UltimaModificacion = " " ' Al modelo 199, la fecha se informa dd/mm/yyyy.
                                      ' A la funcion InsertarWF_Lecturas, al parametro ternro se definio como Long

'Global Const Version = "1.13"
'Global Const FechaModificacion = "03/06/2008" 'FAF
'Global Const UltimaModificacion = " " ' Se agrego el modelo 198 para Cargil

'Global Const Version = "1.14"
'Global Const FechaModificacion = "04/06/2008" 'FAF
'Global Const UltimaModificacion = " " ' Se agrego el modelo 197 para Andreani

'Global Const Version = "1.15"
'Global Const FechaModificacion = "10/06/2008" 'FAF
'Global Const UltimaModificacion = " " ' Se modifico el modelo 198, de forma que identifica la sucursal informada en el nombre del archivo

'Global Const Version = "1.16"
'Global Const FechaModificacion = "05/09/2008" 'Lisandro Moro
'Global Const UltimaModificacion = " " ' Se agrego el modelo 196, ActionLine


'Global Const Version = "1.17"
'Global Const FechaModificacion = "22/09/2008" 'FGZ
'Global Const UltimaModificacion = " " ' Se modifico el Main para que no inserte los procesos directamente en Pendiente
                                      ' Se modifico el modelo 205 (sub InsertaFormato2)

'Global Const Version = "1.18"
'Global Const FechaModificacion = "23/09/2008" 'Cesar Stankunas
'Global Const UltimaModificacion = "24/09/2008" ' Se agrego el modelo 195 para Repsa


'Global Const Version = "1.19"
'Global Const FechaModificacion = "16/12/2008" 'FGZ
'Global Const UltimaModificacion = " " '
'                                      ' Se modifico el modelo 197 (sub InsertaFormato12) Andreani
'                                      ' Si el reloj no tiene la marca de control de acceso ==>
'                                      '     las registraciones se insertan con estado 'X' y sin reloj asociado para que no las tome el PRC30
'
'Global Const Version = "1.20"
'Global Const FechaModificacion = "13/01/2009" 'Cesar Stankunas
'Global Const UltimaModificacion = "14/01/2009" ' Se agrego el modelo 187 para Canal9


'Global Const Version = "1.21"
'Global Const FechaModificacion = "21/01/2009" 'FGZ
'Global Const UltimaModificacion = " " '
'                                      'Encriptacion de string de conexion

'Global Const Version = "1.22"
'Global Const FechaModificacion = "15/05/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Cambio de Formato para Modelo 201 (InsertaFormato9)
                                      
                                      
'Global Const Version = "1.23"
'Global Const FechaModificacion = "10/08/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 186 (InsertaFormato16)
                                      '
'Global Const Version = "1.24"
'Global Const FechaModificacion = "19/08/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 185 (InsertaFormato17)


'Global Const Version = "1.25"
'Global Const FechaModificacion = "25/08/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Cambio de Formato para Modelo 201 (InsertaFormato9)

'Global Const Version = "1.26"
'Global Const FechaModificacion = "07/09/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Cambio de Formato para Modelo 187 (InsertaFormato15) - Canal9 (TELEARTE)


'Global Const Version = "1.26"
'Global Const FechaModificacion = "08/09/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Cambio de Formato para Modelo 186 (InsertaFormato16) - AMIA

'Global Const Version = "1.27"
'Global Const FechaModificacion = "14/09/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Modelo 186 (InsertaFormato16) - AMIA
''                                       Se cambió la validacion de tarjeta por el nro de legajo


'Global Const Version = "1.28"
'Global Const FechaModificacion = "01/10/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Cambio de Formato para Modelo 201 (InsertaFormato9) - AGD
''                                       Se cambió el tipo de codigo asociado a JDE. Antes 130 ahora 140 dado que el 130 ya lo estaban utilizando


'Global Const Version = "1.29"
'Global Const FechaModificacion = "08/10/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Modelo 201 (InsertaFormato9) - AGD
''                                      Cambio de las librerias para Produccion


'Global Const Version = "1.30"
'Global Const FechaModificacion = "09/11/2009" 'FGZ
'Global Const UltimaModificacion = " " 'Modelo 201 (InsertaFormato9) - AGD
''                                      Cambio de las librerias para Produccion
''                                      cambiaron los nombres de algunos campos en la libreria nueva PD812DTA.F554312

'Global Const Version = "1.31"
'Global Const FechaModificacion = "04/01/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Modelo 201 (InsertaFormato9) - AGD
''                                      Cambio de las librerias para Produccion PY812DTA por PD812DTA


'Global Const Version = "1.32"
'Global Const FechaModificacion = "02/03/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 184 (InsertaFormato18) - Supercanal


'Global Const Version = "1.33"
'Global Const FechaModificacion = "05/03/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 183 (InsertaFormato19) - Servicentro la Estrella

'Global Const Version = "1.34"
'Global Const FechaModificacion = "22/03/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Formato de registraciones Modelo 183 (InsertaFormato19) - Servicentro la Estrella
''                                       La ultima linea no debe leerse, es un registro de cantidad de registraciones leidas por el reloj
'

'Global Const Version = "1.35"
'Global Const FechaModificacion = "30/04/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Formato de registraciones Modelo 185 (InsertaFormato17) - Schering Plough
''                                       Cuando se lee de un reloj sin la marca de control de acceso las registraciones quedan en estado X para que los procesos no las tomen.


'Global Const Version = "1.36"
'Global Const FechaModificacion = "27/05/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Formato de registraciones Modelo 183 (InsertaFormato19) - Servicentro la Estrella
''                                      Habia problemas de validacion cuando la hora era 12 p.m

'Global Const Version = "1.37"
'Global Const FechaModificacion = "02/06/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Formato de registraciones Modelo 183 (InsertaFormato19) - Servicentro la Estrella
'                                      Habia problemas de validacion cuando la hora era 12 a.m.

'Global Const Version = "1.38"
'Global Const FechaModificacion = "11/06/2010" 'EGO
'Global Const UltimaModificacion = " " 'Se Agrego Modelo 182 (InsertaFormato20)- Andreani
''                                      Se ingresan las registraciones a traves de una tabla wc_bajada_reg

'Global Const Version = "1.39"
'Global Const FechaModificacion = "05/08/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Se Agrego Modelo 181 (InsertaFormato21)- Monresa
''                                      Formato condicional

'Global Const Version = "1.40"
'Global Const FechaModificacion = "13/09/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Se modificó el Modelo 181 (InsertaFormato21)- Monresa
''                                      El Formato 1(marcas HP.txt) estaba tomando solo un digito para los relojes. Ahora toma 2

'Global Const Version = "1.41"
'Global Const FechaModificacion = "03/12/2010" 'FGZ
'Global Const UltimaModificacion = " " 'Se modificó el Modelo 200 (InsertaFormato10)- Sidersa
''                                      los nros de tarjeta pasan de 6 a 8

'Global Const Version = "1.42"
'Global Const FechaModificacion = "18/03/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Se modificó el Modelo 180 (InsertaFormato22)- SYKES
''                                      el campo empleg de la tabla wc_bajada_reg es string

'Global Const Version = "1.43"
'Global Const FechaModificacion = "06/04/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Se modificó el Modelo 180 (InsertaFormato22)- SYKES
''                                      se agregó un campo identity a la tabla wc_bajada_reg (regnro) para utilizarlo para actualizar la tabla

'Global Const Version = "1.44"
'Global Const FechaModificacion = "13/04/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Se modificó el Modelo 180 (InsertaFormato22)- SYKES
''                                      se agregó actualizacion del progreso del proceso

'Global Const Version = "1.45"
'Global Const FechaModificacion = "30/04/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Formato de registraciones Modelo 185 (InsertaFormato17) - Schering Plough
''                                       se quitaron espacios en blanco en algunos campos(reloj y marca de entrada salida)

'Global Const Version = "1.46"
'Global Const FechaModificacion = "13/05/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Se Agrego Modelo 179 (InsertaFormato23)- Mundo Maipú

'Global Const Version = "1.47"
'Global Const FechaModificacion = "02/06/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Se modificó el Modelo 180 (InsertaFormato22)- SYKES
''                                      Cuando no encuentre los empleados asociados a una tarjeta las marque y no las vuelva a tomar
''                                       Marca es estado con X y el pronro con el nro de proceso de lectura

'Global Const Version = "1.48"
'Global Const FechaModificacion = "03/06/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Se modificó el Modelo 180 (InsertaFormato22)- SYKES
''                                       Habia quedado un error en el update.

'Global Const Version = "1.49"
'Global Const FechaModificacion = "06/09/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Se modificó el Modelo 184 (InsertaFormato18)- BAPRO
''                                       Valida que no haya mas de un empleado con el mismo DNI

'Global Const Version = "1.50"
'Global Const FechaModificacion = "21/09/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Se modificó el Modelo 184 (InsertaFormato18)- BAPRO
'                                       Cuando hay mas de un empleado con el mismo DNI ...se queda con el empleado activo.
'                                       Si hubiese mas de uno activo dará un mensaje de error y no insertará registraciones.

'Global Const Version = "1.51"
'Global Const FechaModificacion = "30/09/2011" 'Gonzalez Nicolás
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 178 (InsertaFormato24)- CAJA DE ODONTOLOGOS

'Global Const Version = "1.52"
'Global Const FechaModificacion = "30/09/2011" 'Gonzalez Nicolás
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 177 (InsertaFormato25)- Laboratorios SL- QA
'                                      '---------------------FALTAN CORROBORAR DATOS PARA TERMINARLO----------------------'

'Global Const Version = "1.53"
'Global Const FechaModificacion = "06/10/2011" 'Gonzalez Nicolás
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 176 - SYKES :
'                                      Anexo del modelo 180 formato InsertaFormato26 (que es una copia del modelo 207).

'Global Const Version = "1.54"
'Global Const FechaModificacion = "17/10/2011" 'Gonzalez Nicolás
'Global Const UltimaModificacion = " " 'Modelo 177 - (InsertaFormato25)- Laboratorios SL- QA:
''                                      Se corroboraron datos faltantes. Id de persona = Legajo | N° de reloj, se levanta desde la base


'Global Const Version = "1.55"
'Global Const FechaModificacion = "15/11/2011" 'Gonzalez Nicolás
'Global Const UltimaModificacion = " " 'Modelo 176 - (InsertaFormato26)- Sykes
'                                      15/11/2011 - Gonzalez Nicolás - Se cambio la forma de guardar el registro EntradaSalida

'Global Const Version = "1.56"
'Global Const FechaModificacion = "30/11/2011" 'Zamarbide Juan
'Global Const UltimaModificacion = " " 'Modelo 348 - (InsertaFormato27)- Laboratorio LS
'' CAS -14391- LABORATORIO LS -  Modelo de reloj V2

'Global Const Version = "1.57"
'Global Const FechaModificacion = "29/12/2011" 'FGZ
'Global Const UltimaModificacion = " " 'Modificaciones generales sobre el control de instancia previa corriendo en memoria

'Global Const Version = "1.58"
'Global Const FechaModificacion = "27/04/2012" 'Gonzalez Nicolás
'Global Const UltimaModificacion = " " 'Modelo 208 - (InsertaFormato5) - Se valida que el reloj se de CONTROL DE ACCESO
'                                      CAS-15711 - Farmografica - modelo de lectura registraciones


'Global Const Version = "1.59"
'Global Const FechaModificacion = "02/07/2012" 'Deluchi Ezequiel
'Global Const UltimaModificacion = " " 'Modelo 204 - (InsertaFormato1) - Se valida que el codext del reloj con valor alfanumerico
'                                      CAS-16311 - TABACAL - GTI - Carga de Registraciones - modificación

'Global Const Version = "1.60"
'Global Const FechaModificacion = "03/07/2012" 'Gonzalez Nicolás
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 191 (InsertaFormato28)
'                                      CAS-16329 - MIMO - CUSTOM RELOJ GALSYS

'Global Const Version = "1.61"
'Global Const FechaModificacion = "17/09/2012" 'Manterola Maria Magdalena
'Global Const UltimaModificacion = " " 'Modificación - Modelo 179 - InsertaFormato23
''                                      CAS-15535 - Mundo Maipu - Inhabilitacion de Relojes

'Global Const Version = "1.62"
'Global Const FechaModificacion = "26/10/2012" 'FGZ
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 175 (InsertaFormato29)
''                                      CAS-17037 - TIMBO - Interfaz para SPEC

'Global Const Version = "1.63"
'Global Const FechaModificacion = "10/01/2013" 'FGZ
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 174 (InsertaFormato30)
'                                      CAS-17915 - HB - (VISION) - Custom en Interfaz de Registraciones

'Global Const Version = "1.64"
'Global Const FechaModificacion = "22/02/2013" 'Sebastian Stremel
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 173 (InsertaFormatoSpec)
'                                      CAS-18151 - Akzo - Interfaces Spec

'Global Const Version = "1.65"
'Global Const FechaModificacion = "21/05/2013" 'Sebastian Stremel
'Global Const UltimaModificacion = " " 'Formato de registraciones Modelo 173 (InsertaFormatoSpec)
''                                      Se modifico que si el resultado no es E,S o "" entonces ponga la entrada/salida en vacio, ya que rhpro no acepta otros codigos.
''                                      CAS-18151 - Akzo - Interfaces Spec


'Global Const Version = "1.66"
'Global Const FechaModificacion = "03/06/2013" 'FGZ
'Global Const UltimaModificacion = " " 'Formato de registraciones Modelo 192 (InsertaFormatoSpec2)
''                                      CAS-19918 - HORWTH LITORAL - GEMPLAST - Formato de Reloj

'Global Const Version = "1.67"
'Global Const FechaModificacion = "04/06/2013" 'FGZ
'Global Const UltimaModificacion = " " 'Formato de registraciones Modelo 193 (InsertaFormatoSpec2)
'                                      CAS-19919 - AMR - Nuevo formato de Reloj

'Global Const Version = "1.68"
'Global Const FechaModificacion = "18/06/2013" 'FAF
'Global Const UltimaModificacion = " " 'Formato de registraciones Modelo 194 (InsertaFormatoCargillGOSC)
'                                      CAS-11908 - Punto 10 - Este caso contiene varios puntos, referido a registraciones es el 10, donde a raiz del pedido se genera
'                                      un nuevo Modelo de Lectura para GOSC - Cargill. Se mantiene el 198 para Flour - Harina - Trigaglia

'Global Const Version = "1.69"
'Global Const FechaModificacion = "23/07/2013" 'Sebastian Stremel
'Global Const UltimaModificacion = " " 'Se creo el modelo 172 --> InsertarFormatoRaffo - CAS-18990 - Raffo - Adecuaciones GTI - Intefaces Registraciones


'Global Const Version = "1.70"
'Global Const FechaModificacion = "17/09/2013" 'Mauricio Zwenger
'Global Const UltimaModificacion = " " 'Se creo el modelo 171 --> InsertarFormato31 - CAS-21338 - Owens - Formato de reloj

'Global Const Version = "1.71"
'Global Const FechaModificacion = "18/09/2013" 'Mauricio Zwenger
'Global Const UltimaModificacion = " " 'Se creo el modelo 170 --> InsertarFormato32 - CAS-21337 - SAN CAMILO - GTI-QA-Formato de Reloj

'Global Const Version = "1.72"
'Global Const FechaModificacion = "11/10/2013" 'Mauricio Zwenger
'Global Const UltimaModificacion = " " 'Se creo el modelo 169 --> InsertarFormato33 - CAS-21170 - SGS - Interfase de levantamiento de registraciones con parte de movilidad

'Global Const Version = "1.73"
'Global Const FechaModificacion = "05/12/2013" 'Dimatz Rafael
'Global Const UltimaModificacion = " " 'Se creo el modelo 167 --> InsertarFormato34 - CAS-22139 - VSO - LAB ETICOS

'Global Const Version = "1.74"
'Global Const FechaModificacion = "10/12/2013" 'Dimatz Rafael
'Global Const UltimaModificacion = " " 'Se creo el modelo 166 --> InsertarFormatoSpec3 - CAS-22668 - 5CA - Nuevo formato de reloj

'Global Const Version = "1.75"
'Global Const FechaModificacion = "09/01/2014" 'Dimatz Rafael
'Global Const UltimaModificacion = " " 'Se creo el modelo 167 --> Se Modifico InsertarFormato34 para obtener los datos de la registracion correctamente

'Global Const Version = "1.76"
'Global Const FechaModificacion = "13/01/2014" 'Dimatz Rafael
'Global Const UltimaModificacion = " " 'Se modifico en la Registracion la I por E y la O por S en insertaFormato34

'Global Const Version = "1.77"
'Global Const FechaModificacion = "17/03/2014" 'Mauricio Zwenger
'Global Const UltimaModificacion = " " 'Se creo el modelo 169 --> InsertarFormato33 - CAS-21170 - SGS - Interfase de levantamiento de registraciones con parte de movilidad


'Global Const Version = "1.78"
'Global Const FechaModificacion = "28/04/2014" 'FGZ
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 165 (InsertaFormatoMarkovations)
''                                      CAS-24771 - VILLA MARIA - ARCHIVO LECTURA DE REGISTRACIONES

'
'Global Const Version = "1.79"
'Global Const FechaModificacion = "17/06/2014" 'EAM
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 164
''                                      CAS-25421 - Lucaioli

'Global Const Version = "1.80"
'Global Const FechaModificacion = "19/06/2014" 'EAM
'Global Const UltimaModificacion = " " 'Nuevo Formato de registraciones Modelo 164
''                                      CAS-25421 - Lucaioli
'
'
'Global Const Version = "1.81"
'Global Const FechaModificacion = "14/07/2014" 'EAM
'Global Const UltimaModificacion = " " 'Se modificó el modelo 205 para validar alcance por estructura de los relojes (si al menos hay un reloj con alcance configurado)
'                                      CAS-26142 - HIRSCH - Error en configuracion de relojes

'Global Const Version = "1.82"
'Global Const FechaModificacion = "06/08/2014" 'MDZ
'Global Const UltimaModificacion = " " ' Se corrigio bug en lectura de ultimo campo para el modelo 170 (InsertaFormato32())
                                        'CAS-21337 - SAN CAMILO - GTI-QA-Formato de Reloj


'Global Const Version = "1.83"
'Global Const FechaModificacion = "05/09/2014" 'MDZ
'Global Const UltimaModificacion = " " ' Se modifico InsertaFormato2 correspondiente al modelo 205 para que el campo correspondiente a E/S acepte 1 o 2 correspondientes a 1 = Salida Intermedia y 2 = Entrada Intermedia
'                                      ' CAS-26995 - Union Personal - Nuevo formato de registraciones

'Global Const Version = "1.84"
'Global Const FechaModificacion = "03/11/2014" 'LED
'Global Const UltimaModificacion = " " ' se creo nuevo modelo de registracion (163) InsertaFormatoPollPar()
                                      ' CAS-27481 - VSO - POLLPAR - Nuevo modelo lectura registraciones


'Global Const Version = "1.85"
'Global Const FechaModificacion = "07/11/2014"
'Global Const UltimaModificacion = " "  ' Matias Fernandez
                                       ' CAS-26028 - H&A - Modificaciones R4 - Agregar modelo de reloj al log de lectura de registraciones
                                       ' se agrego al log el numero de modelo que se ejecuto.

'Global Const Version = "1.86"
'Global Const FechaModificacion = "12/11/2014"
'Global Const UltimaModificacion = " "  ' Matias Fernandez
                                       ' CAS-26028 - H&A - Modificaciones R4 - Agregar modelo de reloj al log de lectura de registraciones
                                       ' se agrego al log el numero de modelo que se ejecuto.

'Global Const Version = "1.87"
'Global Const FechaModificacion = "02/12/2014"
'Global Const UltimaModificacion = " "  ' Mauricio Zwenger
                                       ' CAS-27487 - ALTA PLASTICA - CUSTOM FORMATO DE REGISTRACIONES
                                       ' Se modifico funcion InsertaFormato31 para que valide tambien, una variante del formato antes levantado

'Global Const Version = "1.88"
'Global Const FechaModificacion = "18/03/2015"
'Global Const UltimaModificacion = " "  ' Sebastian Stremel
                                       ' CAS-29890 - Claxson - Custom Nuevo Formato de reloj
                                       ' Se creo el modelo 162 de lectura de registraciones


'Global Const Version = "1.89"
'Global Const FechaModificacion = "26/03/2015"
'Global Const UltimaModificacion = " "  ' Sebastian Stremel
                                       ' CAS-29890 - Claxson - Custom Nuevo Formato de reloj [Entrega 2]
                                       ' Se agrego el reloj por defecto al modelo 162
                                       
'Global Const Version = "1.90"
'Global Const FechaModificacion = "27/04/2015"
'Global Const UltimaModificacion = " "  ' Sebastian Stremel
                                       ' CAS-30518 - SUIZO - Nuevo Formato de Reloj
                                       ' Se creo el modelo 162 de lectura de registraciones - se corrige comentario el modelo es 161
                                       
                                       
'Global Const Version = "1.91"
'Global Const FechaModificacion = "27/04/2015"
'Global Const UltimaModificacion = " "  ' Sebastian Stremel
                                       ' CAS-29890 - Claxson - Custom Nuevo Formato de reloj
                                       ' Se creo el modelo 160 de lectura de registraciones
                                       
'Global Const Version = "1.92"
'Global Const FechaModificacion = "04/05/2015"
'Global Const UltimaModificacion = " "  ' Sebastian Stremel
                                       ' CAS-29890 - Claxson - Custom Nuevo Formato de reloj Modelo 2 [Entrega 2]
                                       ' Correccion en el modelo 160 de lectura de registraciones, error en el insert.
                                       
                                        
'Global Const Version = "1.93"
'Global Const FechaModificacion = "08/05/2015"
'Global Const UltimaModificacion = " "  ' Miriam Ruiz
                                       ' CAS-28352 - Salto Grande - Custom GTI - Archivo de registraciones
                                       ' Se creo el modelo 159 de lectura de registraciones para Salto Grande
                                       
'Global Const Version = "1.94"
'Global Const FechaModificacion = "19/05/2015"
'Global Const UltimaModificacion = " "  ' Carmen Quintero
                                       ' CAS-30743 - POLLPAR - Error en relojes por estructuras
                                       ' Se modificaron todos los modelos excepto el 205, para que validen alcance por estructura de los relojes (si al menos hay un reloj con alcance configurado)
                                       
'Global Const Version = "1.95"
'Global Const FechaModificacion = "05/06/2015"
'Global Const UltimaModificacion = " "  ' LED
                                       ' CAS-30945 - ASM - Nuevo formato de reloj
                                       ' nuevo modelo 158 de lectura de registraciones para ASM
                                       
                                       
'Global Const Version = "1.96"
'Global Const FechaModificacion = "09/09/2015"
'Global Const UltimaModificacion = " " 'MDF
                                      'CAS-32778 - Monasterio base 1 - Bug en procesamiento
                                      'Se abre el objeto de conexion de progreso

'Global Const Version = "1.97"
'Global Const FechaModificacion = "06/11/2015"
'Global Const UltimaModificacion = " " 'LED
                                      'CAS-33933 - Salto Grande - Archivos de registraciones
                                      'Cambio en el modelo 175
                                      'Se amplio el campo de reloj a 25 caracteres, primero busca en cod externo, si no existe lo busca en la descripcion abreviada.

'Global Const Version = "1.98"
'Global Const FechaModificacion = "12/01/2016"
'Global Const UltimaModificacion = " " 'Gonzalez Nicolás - CAS-33995 - SOLAR - Nuevo formato de reloj
                                      'Nuevo modelo : 157
                                      
'Global Const Version = "1.99"
'Global Const FechaModificacion = "12/02/2016"
'Global Const UltimaModificacion = " " 'Dimatz Rafael - CAS-35560 - IQ FARMA - FORMATO PARA ARCHIVO DE REGISTRACIONES
'                                      'Nuevo modelo : 156

'Global Const Version = "2.00"
'Global Const FechaModificacion = "10/03/2016"
'Global Const UltimaModificacion = " " 'Dimatz Rafael - CAS-35560 - IQ FARMA - FORMATO PARA ARCHIVO DE REGISTRACIONES
'                                      'Se deja Fija la E en el campo Entrada Salida ya que no distingue entre E/S

'Global Const Version = "2.01"
'Global Const FechaModificacion = "11/03/2016"
'Global Const UltimaModificacion = " " 'FGZ - CAS-35345 - IQ FARMA - QA Bug Nro Tarjeta
                                        ' Se modificó la validación del nro de tarjeta para que complee a 10 digitos

'Global Const Version = "2.02"
'Global Const FechaModificacion = "21/04/2016"
'Global Const UltimaModificacion = " " 'Gonzalez Nicolás - CAS-36794- TNPlatex - Nuevo formato de reloj
                                       'Nuevo modelo: 155

Global Const Version = "2.03"
Global Const FechaModificacion = "25/04/2016"
Global Const UltimaModificacion = " " 'Gonzalez Nicolás - CAS-36794- TNPlatex - Nuevo formato de reloj
                                       'Modificado: 155 - Se busca el reloj solamente por el cód. externo (nodo)

'------------------------------------------------------------------------
'------------------------------------------------------------------------
Dim fs, f
Global Flog
'FGZ - 16/12/2008 - le saqué esta definicion por la global de abajo
'Dim objFechasHoras As New FechasHoras
Global objFechasHoras As New FechasHoras

Dim NroLinea As Long
Dim crpNro As Long
Dim RegLeidos As Long
Dim RegError As Long
Dim regfecha As Date
Dim NroProceso As Long

Global Path As String
Global NArchivo As String
Global Rta
Global ObjetoVentana As Object
Global HuboError As Boolean
Global Nro_Modelo As Integer
Global Etiqueta
Global separador As String

'FGZ - 26/03/2007
Global Cantidad_de_OpenRecordset As Long
Global Cantidad_Call_Politicas As Long

'FGZ - 20/04/2007
Global Usuario As String
Global objConn2 As New ADODB.Connection
Global Version_Valida As Boolean

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Private Sub InsertaFormato5(strreg As String)
'Modificado : 27/04/2012 -  Gonzalez Nicolás - Se valida que el reloj se de CONTROL DE ACCESO | CAS-15711 - Farmografica - modelo de lectura registraciones
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean
Dim regestado

    Reg_Valida = False
    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Fecha = Mid(strreg, pos1, pos2 - pos1)
    regfecha = Fecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline "Error. Hora no valida " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    nroreloj = Mid(strreg, pos1, pos2 - pos1)
    nrorelojtxt = Mid(strreg, pos1, pos2 - pos1)
    
    '27/04/2012 - NG - Ahora tengo en cuenta si el reloj asociado tiene la marca de control de acceso
    StrSql = "SELECT relnro, tptrnro,relvalestado FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro,relvalestado FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. No se encontro el Reloj: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
            Reg_Valida = CBool(objRs!relvalestado)
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
        Reg_Valida = CBool(objRs!relvalestado)
    End If
     '27/04/2012 - NG - Ahora tengo en cuenta si el reloj asociado tiene la marca de control de acceso
    If Reg_Valida Then
        regestado = "I"
    Else
        regestado = "X"
    End If
    
    

    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = IIf(Trim(Mid(strreg, pos1)) = "20", "E", "S")
       
    ' 15/07/2003
    ' no poner comillas al nroLegajo porque no toma bien los legajos que comienzan con 000...
    'If codReloj = 13 Then
    '    StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = " & nroLegajo & " AND tptrnro = 2 AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    'Else
    '    StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = " & nroLegajo & " AND tptrnro = 1 AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    'End If

    ' ----------------------------------------------------
    'FZG 06/08/2003
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      OpenRecordset StrSql, objRs
      If Not objRs.EOF Then
         Ternro = objRs!Ternro
      Else
         Flog.writeline "Error. No se encontro la tarjeta para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj: " & codReloj
         Flog.writeline "SQL: " & StrSql
         InsertaError 1, 33
         Exit Sub
      End If
    End If
    ' Primero se busca con el numero de tarjeta y el tipo asociado al reloj. Si no se encuentra,
    ' se busca solo con el número de la tarjeta. O.D.A. 06/08/2003
    ' ----------------------------------------------------
        
    
    'Carmen Quintero - 15/05/2015
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(Fecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'" & regestado & "')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        ' FGZ - 09/10/2003
        ' Inserto el par (Ternro,Fecha)
        If Reg_Valida Then ' NG - 27/04/2012
            Call InsertarWF_Lecturas(Ternro, Fecha)
        End If
    Else
        Flog.writeline " Registracion ya Existente "
        Flog.writeline " Hora: " & Hora & " - Fecha: " & regfecha
        InsertaError 1, 92
    End If
End Sub




Private Sub InsertaFormato4_old(strreg As String)

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer

    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Fecha = Mid(strreg, pos1, pos2 - pos1)
    regfecha = Fecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    nroreloj = Mid(strreg, pos1, pos2 - pos1)
    nrorelojtxt = Mid(strreg, pos1, pos2 - pos1)
    
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline " Error Reloj: " & nroreloj
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    

    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = IIf(Trim(Mid(strreg, pos1)) = "20", "E", "S")
       
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      OpenRecordset StrSql, objRs
      If Not objRs.EOF Then
         Ternro = objRs!Ternro
      Else
         Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
         InsertaError 1, 33
         Exit Sub
      End If
    End If
        
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        ' FGZ - 09/10/2003
        ' Inserto el par (Ternro,Fecha)
        Call InsertarWF_Lecturas(Ternro, Fecha)
        
    Else
        Flog.writeline " Registracion ya Existente "
        InsertaError 1, 92
    End If
        
End Sub


Private Sub InsertaFormato3(strreg As String)

Dim NroLegajo As String

Dim Ternro As Long
Dim Fecha As String
Dim Hora As String
Dim entradasalida As String

Dim nroreloj As String

Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean

    RegLeidos = RegLeidos + 1
    
    nroreloj = Mid(strreg, 1, 5)
    
    entradasalida = IIf(Mid(strreg, 6, 2) = "20", "E", "S")
    
    Fecha = (Mid(strreg, 12, 2) & "/" & Mid(strreg, 10, 2) & "/20" & Mid(strreg, 8, 2))
    
    Hora = Mid(strreg, 14, 4)
    
    NroLegajo = Mid(strreg, 18, 5)
    
    Flog.writeline NroLegajo & ";" & Fecha & ";" & Hora & ";" & nroreloj & ";" & entradasalida

    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error. Hora no valida " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Error no se encontro el Reloj: " & nroreloj
        Flog.writeline "SQL: " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj
    StrSql = StrSql & " AND hstjnrotar = '" & NroLegajo
    StrSql = StrSql & "' AND (hstjfecdes <= " & ConvFecha(Fecha)
    StrSql = StrSql & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
        Ternro = objRs!Ternro
    Else
        Flog.writeline " Error. No se encontro tarjeta para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj:" & codReloj
        Flog.writeline "SQL: " & StrSql
        InsertaError 1, 33
        Exit Sub
    End If
        
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(Fecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
        
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & Hora
    StrSql = StrSql & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, Fecha)
        
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: " & ConvFecha(regfecha)
        InsertaError 1, 92
    End If
        
End Sub


Private Sub InsertaFormato2(strreg As String)

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean
Dim TipoReg As String

    On Error GoTo ME_Local
    
    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Fecha = Mid(strreg, pos1, pos2 - pos1)
    regfecha = Fecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = UCase(Trim(Mid(strreg, pos1, pos2 - pos1)))
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    nroreloj = Trim(Mid(strreg, pos1))
    nrorelojtxt = Trim(Mid(strreg, pos1))
    
    'FGZ - 14/07/2014 --------------------------------
    'StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    'OpenRecordset StrSql, objRs
    'If objRs.EOF Then
    '    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
    '    OpenRecordset StrSql, objRs
    '    If objRs.EOF Then
    '        Flog.writeline "Error. Reloj no encontrado " & nroreloj
    '        Flog.writeline "SQL: " & StrSql
    '        InsertaError 4, 32
    '        Exit Sub
    '    Else
    '        codReloj = objRs!relnro
    '        tipotarj = objRs!tptrnro
    '    End If
    'Else
    '    codReloj = objRs!relnro
    '    tipotarj = objRs!tptrnro
    'End If
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro,relvalestado FROM gti_reloj WHERE relcodext = '" & Trim(nroreloj) & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro,relvalestado FROM gti_reloj WHERE relcodext = '" & Trim(nrorelojtxt) & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
            Reg_Valida = CBool(objRs!relvalestado)
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
        Reg_Valida = CBool(objRs!relvalestado)
    End If
    'FGZ - 14/07/2014 --------------------------------
    
    
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      OpenRecordset StrSql, objRs
      If Not objRs.EOF Then
         Ternro = objRs!Ternro
      Else
        StrSql = "SELECT ternro FROM gti_histarjeta WHERE convert(bigint,hstjnrotar) = " & CLng(NroLegajo) & " AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
           Ternro = objRs!Ternro
        Else
           Flog.writeline "Error. Nro de tarjeta no encontrado para el legajo: " & NroLegajo & ", Tipo de tarjeta: " & tipotarj & " y codigo de reloj:  " & codReloj
           Flog.writeline "SQL: " & StrSql
           InsertaError 1, 33
           Exit Sub
        End If
      End If
    End If
        
    'FGZ - 14/07/2014 --------------------------------
    StrSql = "SELECT relnro FROM gti_rel_estr "
    'StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'FGZ - 14/07/2014 --------------------------------
    
    'MDZ - 05/09/2014 --------------------------------
    Select Case entradasalida
        Case "E"
            TipoReg = "00"
        
        Case "S"
            TipoReg = "01"
        
        Case "1"
            TipoReg = "02"
            entradasalida = "S"
        
        Case "2"
            TipoReg = "03"
            entradasalida = "E"
    End Select
    
    'MDZ - 05/09/2014 --------------------------------
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado, tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I','" & TipoReg & "')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado, tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X','" & TipoReg & "')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, Fecha)
    Else
        Flog.writeline " Registracion ya Existente "
        Flog.writeline " Hora: " & Hora & " - Fecha: " & regfecha
        InsertaError 1, 92
    End If
        
Fin:
Exit Sub
ME_Local:
    HuboError = True
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    GoTo Fin
End Sub


Private Sub InsertaFormato1(strreg As String)

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean

    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Legajo:  " & NroLegajo
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    regfecha = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Fecha:  " & regfecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    'NroReloj = Mid(strReg, pos1, pos2 - pos1)
    nrorelojtxt = Trim(Mid(strreg, pos1, pos2 - pos1))
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    Flog.writeline "Busco el reloj"
    'StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & NroReloj & "'"
    'OpenRecordset StrSql, objRs
    'If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nrorelojtxt
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    'Else
        'codReloj = objRs!relnro
        'TipoTarj = objRs!tptrnro
    'End If
    

    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = IIf(Trim(Mid(strreg, pos1)) = "20", "E", "S")
       
    Flog.writeline "Busco el nro de tarjeta "
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      'StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & nroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      'OpenRecordset StrSql, objRs
      'If Not objRs.EOF Then
      '   Ternro = objRs!Ternro
      'Else
         Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " , Reloj: " & codReloj
         Flog.writeline "SQL: " & StrSql
         InsertaError 1, 33
         Exit Sub
      'End If
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
End Sub

Private Sub InsertaError(NroCampo As Byte, nroError As Long)

    'Flog.writeline "antes de insertar error en car_err" & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'StrSql = "INSERT INTO Car_err(crpnnro,inerrnro,nrolinea,campnro) VALUES (" & _
    '         crpNro & "," & nroError & "," & NroLinea & "," & NroCampo & ")"
    'objConn.Execute StrSql, , adExecuteNoRecords
    
    'Flog.writeline "insertó en car_err" & Format(Now, "dd/mm/yyyy hh:mm:ss")
        
    RegError = RegError + 1
    
End Sub


Private Sub Main()
Dim rs_btp As New ADODB.Recordset
Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim strCmdLine As String
Dim Nombre_Arch As String
Dim Rs_WF_Lec_Fechas As New ADODB.Recordset
Dim Rs_WF_Lec_Terceros As New ADODB.Recordset
Dim NroProcesoHC As Long
Dim NroProcesoAD As Long

Dim rs_ONLINE As New ADODB.Recordset
Dim Proc_ONLINE As Boolean 'Si el procesamiento On Line está o no activo
Dim HC_ONLINE As Boolean 'Si genera o no procesos de Horario Cumplido por procesamiento On Line
Dim AD_ONLINE As Boolean 'Si genera o no procesos de Acumulado Diario por procesamiento On Line



Dim PID As String
Dim ArrParametros
Dim strreg, NroLegajo, regfecha, Hora, nroreloj, nrorelojtxt, entradasalida, pos1, pos2, pos3
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
    
    
    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas

    'If App.PrevInstance Then End

'    'Abro la conexion
'    OpenConnection strconexion, objConn
    
    'Crea el archivo de log
    ' ----------
    Nombre_Arch = PathFLog & "LecturaReg " & "-" & NroProceso & "-" & Format(Date, "dd-mm-yyyy") & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    On Error Resume Next
    OpenConnection strconexion, objConn
    OpenConnection strconexion, objConnProgreso 'MDF
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion"
        Exit Sub
    End If
    
    
    On Error GoTo ME_Local
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    'FGZ - 28/11/2011 ------
    
    If App.PrevInstance Then
        Flog.writeline "Ya hay una instancia del proceso corriendo. Queda Pendiente." & Format(Now, "dd/mm/yyyy hh:mm:ss")
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Pendiente', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Terminar
    End If
        
    'FGZ - 04/08/2010 --------- Control de versiones ------
    Version_Valida = ValidarV(Version, 22, TipoBD)
    If Not Version_Valida Then
        'Actualizo el progreso
        MyBeginTrans
            StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Error de Version', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
            objConnProgreso.Execute StrSql, , adExecuteNoRecords
        MyCommitTrans
        Flog.writeline
        GoTo Terminar
    End If
    'FGZ - 04/08/2010 --------- Control de versiones ------
        
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprcpid = " & PID & ", bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcprogreso = 0 WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'FGZ - 20/04/2007 --------------------------------------
    StrSql = "SELECT iduser FROM batch_proceso WHERE bpronro = " & NroProceso
    OpenRecordset StrSql, rs_btp
    If Not rs_btp.EOF Then
        Usuario = rs_btp!IdUser
    End If
    'FGZ - 20/04/2007 --------------------------------------
    Flog.writeline "Inicio Transferencia " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    'FGZ  - 09/10/2003
    Call CargarNombresTablasTemporales
    'Creo la tabla temporal
    Call CreateTempTable(TTempWFLecturas)
    
    Call ComenzarTransferencia
    
    Flog.writeline
    Flog.writeline "Procesamiento ON-LINE"
    StrSql = "SELECT * FROM GTI_puntos_proc " & _
             " INNER JOIN GTI_proc_online ON GTI_puntos_proc.ptoprcnro = GTI_proc_online.ptoprcnro " & _
             " WHERE GTI_puntos_proc.ptoprcid = 19 AND GTI_puntos_proc.ptoprcact = -1 "
    OpenRecordset StrSql, rs_ONLINE
    
    Proc_ONLINE = False
    HC_ONLINE = False
    AD_ONLINE = False
    
    If rs_ONLINE.EOF Then
        Proc_ONLINE = False
        Flog.writeline "Procesamiento ON-LINE, Lectura de Registraciones, punto de procesamiento inactivo."
    Else
        Flog.writeline "Procesamiento ON-LINE, hay puntos de procesamiento activos ==>"
        Proc_ONLINE = True
        Do While Not rs_ONLINE.EOF
            Select Case rs_ONLINE!btprcnro
            Case 1:
                HC_ONLINE = True
                Flog.writeline "Puntos de procesamiento: Horario Cumplido"
            Case 2:
                AD_ONLINE = True
                Flog.writeline "Puntos de procesamiento: Acumulado Diario"
            Case Else
                Flog.writeline "Puntos de procesamiento desconocido. " & rs_ONLINE!btprcnro
            End Select
        
            rs_ONLINE.MoveNext
        Loop
    End If
    
    If Proc_ONLINE Then
        StrSql = "SELECT DISTINCT fecha FROM " & TTempWFLecturas
        OpenRecordset StrSql, Rs_WF_Lec_Fechas
    
        Do While Not Rs_WF_Lec_Fechas.EOF
        
        'G. Bauer y N. Trillo - Se Modifico el rango de fecha para que procese el dia anterior y el actual
            If HC_ONLINE Then
                Flog.writeline "genero HC para el " & ConvFecha(Rs_WF_Lec_Fechas!Fecha)
                'Inserto en batch_proceso un HC
                StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, " & _
                         "bprcestado, empnro) " & _
                         "VALUES (" & 1 & "," & ConvFecha(Date) & ",'" & Usuario & "'" & ",'" & Format(Now, "hh:mm:ss ") & "' " & _
                         ", " & ConvFecha(DateAdd("d", -1, Rs_WF_Lec_Fechas!Fecha)) & ", " & ConvFecha(Rs_WF_Lec_Fechas!Fecha) & _
                         ", 'Temp', 0)"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'recupero el numero de proceso generado
                NroProcesoHC = getLastIdentity(objConn, "Batch_Proceso")
                Flog.writeline "Disparo HC. Nro de proceso: " & NroProcesoHC
            End If
            
        'G. Bauer y N. Trillo - Se Modifico el rango de fecha para que procese el dia anterior y el actual
            If AD_ONLINE Then
                Flog.writeline "genero AD para el " & ConvFecha(Rs_WF_Lec_Fechas!Fecha)
                'Inserto en batch_proceso un AD
                StrSql = "INSERT INTO Batch_Proceso (btprcnro, bprcfecha, iduser, bprchora, bprcfecdesde, bprcfechasta, " & _
                         "bprcestado, empnro) " & _
                         "VALUES (" & 2 & "," & ConvFecha(Date) & ",'" & Usuario & "'" & ",'" & Format(Now, "hh:mm:ss ") & "' " & _
                         ", " & ConvFecha(DateAdd("d", -1, Rs_WF_Lec_Fechas!Fecha)) & ", " & ConvFecha(Rs_WF_Lec_Fechas!Fecha) & _
                         ", 'Temp', 0)"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'recupero el numero de proceso generado
                NroProcesoAD = getLastIdentity(objConn, "Batch_Proceso")
                Flog.writeline "Disparo AD. Nro de proceso: " & NroProcesoAD
            End If
            
            Flog.writeline
            Flog.writeline "Inserto en batch_empleados los empleados de los procesos generados"
            'Inserto en batch_empleados los empleados de los procesos generados
            StrSql = "SELECT DISTINCT ternro FROM " & TTempWFLecturas & _
                     " WHERE fecha = " & ConvFecha(Rs_WF_Lec_Fechas!Fecha)
            OpenRecordset StrSql, Rs_WF_Lec_Terceros
            Do While Not Rs_WF_Lec_Terceros.EOF
                ' para HC
                If HC_ONLINE Then
                    'Flog.writeline "    para el HC"
                    StrSql = "INSERT INTO batch_empleado (bpronro, ternro, estado) VALUES (" & _
                             NroProcesoHC & "," & Rs_WF_Lec_Terceros!Ternro & ", NULL )"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                'para AD
                If AD_ONLINE Then
                    'Flog.writeline "    para el AD"
                    StrSql = "INSERT INTO batch_empleado (bpronro, ternro, estado) VALUES (" & _
                             NroProcesoAD & "," & Rs_WF_Lec_Terceros!Ternro & ", NULL )"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                Rs_WF_Lec_Terceros.MoveNext
            Loop
        
            'Actualizo el estado de los procesos a Pendiente
            StrSql = "UPDATE Batch_Proceso "
            StrSql = StrSql & " SET bprcestado ='Pendiente'"
            StrSql = StrSql & " WHERE bpronro = " & NroProcesoHC
            objConn.Execute StrSql, , adExecuteNoRecords
        
            'Actualizo el estado de los procesos a Pendiente
            StrSql = "UPDATE Batch_Proceso "
            StrSql = StrSql & " SET bprcestado ='Pendiente'"
            StrSql = StrSql & " WHERE bpronro = " & NroProcesoAD
            objConn.Execute StrSql, , adExecuteNoRecords
        
        
            Rs_WF_Lec_Fechas.MoveNext
        Loop
    End If
    
    
    If Not HuboError Then
        'StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcprogreso =100, bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    
    Call BorrarTempTable(TTempWFLecturas)
    
Terminar:
    ' eliminar el proceso de la tabla batch_proceso si es que termino correctamente
    Call TerminarTransferencia
    Flog.writeline "Lectura completa. Fin " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.Close
    If rs_btp.State = adStateOpen Then rs_btp.Close
    Set rs_btp = Nothing
Exit Sub
ME_Local:
    Flog.writeline "Importacion de Registraciones Abortada."
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "SQL: " & StrSql
    Flog.writeline "Decripcion: " & Err.Description
End Sub


Public Sub ComenzarTransferencia()

Dim ObjCrp As New ADODB.Recordset

Dim NombreArchivo As String
Dim Directorio As String
Dim CArchivos
Dim archivo
Dim Folder
Dim IncPorc As Single
Dim Progreso As Single
Dim Carpeta As String
'FGZ - 05/10/2011 - cambié la definicion de lugar
Dim fc, F1, S2

    
    StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Directorio = Trim(objRs!sis_dirsalidas)
    Else
        Exit Sub
    End If
    
    StrSql = "SELECT * FROM modelo WHERE modtipo = 3 and modestado = -1"
    OpenRecordset StrSql, ObjCrp
    
    Do While Not ObjCrp.EOF
        
        Select Case ObjCrp!Modnro
            Case 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 165, 166, 167, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 181, 183, 184, 185, 186, 187, 189, 191, 192, 193, 194, 195, 197, 198, 199, 200, 201, 203, 204, 205, 206, 207, 208:
                If Not EsNulo(ObjCrp!modseparador) Then
                    separador = ObjCrp!modseparador
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Separador Default ;"
                    separador = ";"
                End If
                
                Carpeta = Directorio & Trim(ObjCrp!modarchdefault)
                Flog.writeline "Directorio de Registraciones:  " & Carpeta
            
                Set fs = CreateObject("Scripting.FileSystemObject")
            
                Path = Carpeta
            
                'Dim fc, F1, S2
                Set Folder = fs.GetFolder(Carpeta)
                Set CArchivos = Folder.Files
            
                Progreso = 0
                If Not CArchivos.Count = 0 Then
                    Flog.writeline CArchivos.Count & " Archivos de registraciones encontrados " & Format(Now, "dd/mm/yyyy hh:mm:ss")
                    IncPorc = 100 / CArchivos.Count
                End If
            
                HuboError = False
                
                For Each archivo In CArchivos
                    Nro_Modelo = ObjCrp!Modnro
                    NArchivo = archivo.Name
                    Flog.writeline "Archivo:  " & archivo.Name
                    Flog.writeline "Numero de modelo a ejecutar: " & Nro_Modelo
                    Call LeeRegistraciones(Carpeta & "\" & archivo.Name)
                    
                    Progreso = Progreso + IncPorc
                    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
                    objConn.Execute StrSql, , adExecuteNoRecords
                Next
            
            Case 209:
                Call InsertaFormato6
            Case 202:
                Call InsertaFormato8
            Case 196:
                Call InsertaFormato13
            Case 182:
                Call InsertaFormato20
            Case 180:   'Este formato es para Sykes
                Call InsertaFormato22
               
        End Select
    
        ObjCrp.MoveNext
        
    Loop
    
End Sub

Public Sub TerminarTransferencia()

    If Not HuboError Then
        StrSql = "DELETE FROM batch_proceso WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    If objConn.State = adStateOpen Then objConn.Close
    
End Sub


Public Sub TerminarTransferencia_old()
Dim rs_Batch_Proceso As New ADODB.Recordset
Dim rs_His_Batch_Proceso As New ADODB.Recordset


    On Error GoTo ME_Local
    
    If Not HuboError Then
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
    End If
    
    If objConn.State = adStateOpen Then objConn.Close
    
Fin:
Exit Sub

ME_Local:
    Flog.writeline "Copia en el historico abortada."
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    GoTo Fin
End Sub



Private Sub LeeRegistraciones(NombreArchivo As String)
Const ForReading = 1
Const TristateFalse = 0
Dim strlinea As String
Dim Archivo_Aux As String
Dim Carpeta
Dim Intentos As Integer
Dim pos_ini As Integer
Dim pos_fin As Integer
Dim Nro_Sucursal As String
    

    If App.PrevInstance Then Exit Sub

    On Error Resume Next
    Err.Number = 1
    Intentos = 20
    Do Until Err.Number = 0 Or Intentos = 0
        Err.Number = 0
        Set f = fs.getfile(NombreArchivo)
        If f.Size = 0 Then
            Err.Number = 1
            Intentos = Intentos - 1
        End If
    Loop
    On Error GoTo 0
    
    
    On Error GoTo ce
    Set f = fs.OpenTextFile(NombreArchivo, ForReading, TristateFalse)
    
    NroLinea = 0
    RegLeidos = 0
    RegError = 0
    If Not f.AtEndOfStream Then
        StrSql = "INSERT INTO car_pin(modnro,crpnarchivo,crpnregleidos,crpnregerr,crpnfecha,crpndesc,crpnestado) VALUES ( " & _
                                      Nro_Modelo & ",'" & Left(NombreArchivo, 60) & "',0,0," & ConvFecha(C_Date(Date)) & ",'" & Left("Carga : " & Now, 30) & "','I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        crpNro = getLastIdentity(objConn, "car_pin")
        
    End If
    
    Flog.writeline "Procesando archivo " & NombreArchivo
    
    If Nro_Modelo = 198 Or Nro_Modelo = 194 Then
        ' Determino la sucursal que se informa en el nombre del archivo, entre los signos _ y .
        pos_ini = InStr(1, NombreArchivo, "_")
        pos_fin = InStr(1, NombreArchivo, ".")
        Nro_Sucursal = Mid(NombreArchivo, pos_ini + 1, pos_fin - pos_ini - 1)
        If Not IsNumeric(Nro_Sucursal) Then
            Flog.writeline " ***** Error - Nro Sucursal no es numérico --> " & Nro_Sucursal
            GoTo seguir
        Else
            StrSql = "SELECT * FROM estructura WHERE estrnro = " & Nro_Sucursal & " AND tenro=1"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Flog.writeline " ***** Error - La sucursal: " & Nro_Sucursal & " no existe o no es del tipo Sucursal"
                GoTo seguir
            End If
        End If
    End If
    
    Do While Not f.AtEndOfStream
        strlinea = f.ReadLine
        NroLinea = NroLinea + 1
        If Trim(strlinea) <> "" Then
            Select Case Nro_Modelo
                Case 155 'TNPlatex
                    Call InsertaFormato155(strlinea)
                Case 156 'CAS-35560 - IQ FARMA - FORMATO PARA ARCHIVO DE REGISTRACIONES - Dimatz Rafael - 12/02/2016
                        Call InsertaFormatoIQFarma(strlinea)
                Case 157 'SOLAR -> Inicia en la linea N° 3
                    If NroLinea > 2 Then
                        Call InsertaFormato157(strlinea)
                    End If
                Case 158 'CAS-30945 - ASM - Nuevo formato de reloj - LED - 05/06/2015
                    Call InsertaFormato158(strlinea)
                Case 159 'Salto Grande
                    Call InsertaFormatoSalto(strlinea)
                Case 160 'Claxon
                     Call InsertaFormatoClaxonV2(strlinea)
                Case 161 'Suizo
                     Call InsertaFormatoSuizo(strlinea)
                Case 162 'Claxon
                     Call InsertaFormatoClaxon(strlinea)
                Case 163 'POLLPAR
                     Call InsertaFormatoPollPar(strlinea)
                Case 164 'Markovations
                     Call InsertaFormato35(strlinea)
                Case 165 'Markovations
                     Call InsertaFormatoMarkovations(strlinea)
                Case 166
                     Call InsertaFormatoSpec3(strlinea)
                Case 167
                     Call InsertaFormato34(strlinea)
                Case 169
                    Call InsertaFormato33(strlinea)
                Case 170
                    Call InsertaFormato32(strlinea)
                Case 171
                    Call InsertaFormato31(strlinea)
                Case 172
                    Call InsertarFormatoRaffo(strlinea)
                Case 173
                    Call InsertaFormatoSpec(strlinea)
                Case 174
                    Call InsertaFormato30(strlinea)
                Case 175
                    Call InsertaFormato29(strlinea)
                Case 176
                    Call InsertaFormato26(strlinea)
                Case 177
                    Call InsertaFormato25(strlinea)
                Case 178
                    Call InsertaFormato24(strlinea)
                Case 179
                    Call InsertaFormato23(strlinea)
                Case 180    'Tambien hay un modelo 180 que lee de tabla temporal
                    Call InsertaFormato26(strlinea)
                Case 181
                    Call InsertaFormatoMonresa(strlinea)
                Case 183
                    If NroLinea > 1 Then
                        Call InsertaFormato19(strlinea)
                    End If
                Case 184
                    Call InsertaFormato18(strlinea)
                Case 185
                    Call InsertaFormato17(strlinea)
                Case 186
                    Call InsertaFormato16(strlinea)
                Case 187
                    Call InsertaFormato15(strlinea)
                Case 189
                    Call InsertaFormato27(strlinea)
                Case 191 'MIMO - CUSTOM RELOJ GALSYS
                    Call InsertaFormato28(strlinea)
                Case 192    'Formato para SPEC pero cliente Gemplast
                    Call InsertaFormatoSpec2(strlinea)
                Case 193    'Formato para AMR
                    Call InsertaFormatoM193(strlinea)
                Case 194
                    Call InsertaFormatoCargillGOSC(strlinea, Nro_Sucursal)
                Case 195
                    Call InsertaFormato14(strlinea)
                Case 197
                    Call InsertaFormato12(strlinea)
                Case 198
                    Call InsertaFormatoCargil(strlinea, Nro_Sucursal)
                Case 199
                    Call InsertaFormato11(strlinea)
                Case 200
                    Call InsertaFormato10(strlinea)
                Case 201
                    Call InsertaFormato9(strlinea)
                Case 203
                    Call InsertaFormato7(strlinea)
                Case 204
                    Call InsertaFormato1(strlinea)
                Case 205
                    Call InsertaFormato2(strlinea)
                Case 206
                    Call InsertaFormato3(strlinea)
                Case 207
                    Call InsertaFormato4(strlinea)
                Case 208
                    Call InsertaFormato5(strlinea)
            End Select
        End If
    Loop
    
seguir:
    Flog.writeline "Actualizo registros Leidos"
    StrSql = "UPDATE car_pin SET crpnregleidos = " & RegLeidos & _
             ",crpnregerr = " & RegError & _
             " WHERE crpnnro = " & crpNro
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Actualizado. SQL --> " & StrSql
    
    f.Close
    Flog.writeline "Archivo procesado: " & NombreArchivo & " " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    
    Set f = fs.getfile(NombreArchivo)
    Archivo_Aux = Replace(Format(Now, "yyyy-mm-dd hh:mm:ss"), ":", "-") & " " & NArchivo
    
    On Error Resume Next
    f.Move Path & "\bk\" & Mid(Archivo_Aux, 1, Len(Archivo_Aux) - 3) & "bk"
    If Err.Number <> 0 Then
        Flog.writeline Espacios(Tabulador * 0) & "La carpeta Destino no existe. Se creará."
        Set Carpeta = fs.CreateFolder(Path & "\bk")
        f.Move Path & "\bk\" & Mid(Archivo_Aux, 1, Len(Archivo_Aux) - 3) & "bk"
    End If
    'desactivo el manejador de errores
    On Error GoTo 0
    
    Flog.writeline "archivo movido " & Format(Now, "dd/mm/yyyy hh:mm:ss")
Fin:
    Exit Sub
    
ce:
    StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    HuboError = True
    Flog.writeline "Error. " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    Flog.writeline "Ultimo SQL: " & StrSql
    GoTo Fin
End Sub

Private Sub InsertaFormato12(strreg As String)
'------------------------------------------------------------------
'Es necesario un modelo que soporte la lectura de registraciones desde
'un archivo con la siguiente especificacion:
'Linea Ejemplo:        085468,"08","04","30","1725","A1"
'Detalle de Campos: Legajo  aa    mm  dd    hora   nro. Reloj
'------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim NroReloj_aux As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim NroTarj As Integer
Dim NroTarj_aux As String
Dim descarte  As String
Dim Reg_Valida As Boolean

    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    Flog.writeline "   - Registración --> " & strreg
    
    'Legajo
    pos1 = 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    NroLegajo = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")

    'Fecha YY/MM/DD --> convertir a DD/MM/YYYY
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    Anio = "20" & Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    Mes = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    Dia = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    fecha_aux = Dia & "/" & Mes & "/" & Anio
    
    'Hora HHMM
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    Hora = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    
    'Nro reloj
    pos1 = pos2 + 1
    pos2 = Len(strreg)
    NroReloj_aux = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    

'====================================================================
' Validar los parametros Levantados
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & NroLegajo
        GoTo Fin
    Else
        Ternro = objRs!Ternro
    End If
    
    'Que la fecha sea válida
    Dia = Mid(fecha_aux, 1, 2)
    Mes = Mid(fecha_aux, 4, 2)
    Anio = Mid(fecha_aux, 7, 4)
    If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
        Flog.writeline "       ****** Fecha no válida --> " & fecha_aux
        GoTo Fin
    Else
        regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline "       ****** Hora no válida --> " & Hora
        GoTo Fin
    End If
    
'    'Busco el Reloj
'    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & NroReloj_aux & "'"
'    OpenRecordset StrSql, objRs
'    If objRs.EOF Then
'        Flog.writeline "       ****** No se encontro el Reloj. SQL --> " & StrSql
'        Exit Sub
'    Else
'        codReloj = objRs!relnro
'        TipoTarj = objRs!tptrnro
'    End If
    
    
    
    'Busco el Reloj
    'Ahora tengo en cuenta si el reloj asociado tiene la marca de control de acceso
    StrSql = "SELECT relnro, tptrnro,relvalestado FROM gti_reloj WHERE relcodext = '" & NroReloj_aux & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el Reloj. SQL --> " & StrSql
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
        Reg_Valida = CBool(objRs!relvalestado)
    End If
    
    'Carmen Quintero - 15/05/2015
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Reloj: " & codReloj
    
'        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
'                Ternro & "," & crpNro & "," & ConvFecha(RegFecha) & ",'" & Hora & "',''," & codReloj & ",'I')"
'        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 16/12/2008  - si el reloj no tiene la marca de control de acceso ==> se inserta en un estado NN para que el proceso no la tenga en cuenta en el procesamiento
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal"
        'If Reg_Valida Then
            StrSql = StrSql & ",relnro"
        'End If
        StrSql = StrSql & ",regestado) VALUES ( "
        StrSql = StrSql & Ternro & ","
        StrSql = StrSql & crpNro & ","
        StrSql = StrSql & ConvFecha(regfecha) & ","
        StrSql = StrSql & "'" & Hora & "','',"
        'If Reg_Valida Then
            StrSql = StrSql & codReloj & ","
        'End If
        If Reg_Valida Then
            StrSql = StrSql & "'I'"
        Else
            StrSql = StrSql & "'X'"
        End If
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        If Reg_Valida Then
            Call InsertarWF_Lecturas(Ternro, regfecha)
        End If
        'FGZ - 16/12/2008  - si el reloj no tiene la marca de control de acceso ==> se inserta en un estado NN para que el proceso no la tenga en cuenta en el procesamiento
    Else
        Flog.writeline "       ****** Registracion ya Existente"
        Flog.writeline "         Error Legajo: " & NroLegajo & " y Reloj: " & codReloj
        Flog.writeline "         Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub

Private Sub InsertaFormatoCargil(strreg As String, NroSuc As String)

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim NroTarj As Long
Dim NroTarj_aux As String
Dim descarte  As String
Dim Validar As Boolean
Dim tipo_reloj As Integer
Dim Origen As String
Dim comedor As Boolean
Dim Legajo As String
Dim Reg_Valida As Boolean

    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    
    Validar = False
    
    comedor = False
    
    Flog.writeline "   - Registración --> " & strreg
    
    tipo_reloj = 4 ' Por default Tehknosur
    Origen = "Por defecto Tehknosur"
    If Len(strreg) = 17 Then
        tipo_reloj = 1 ' Cargil
        Origen = "Cargil"
    End If
    If (UCase(Mid(strreg, 1, 3)) = "COM") Then
        tipo_reloj = 2 ' Trigalia Formato 1
        Origen = "Trigalia"
    End If
    If (UCase(Mid(strreg, 1, 1)) = "*") Then
        tipo_reloj = 3 ' Trigalia Formato 2
        Origen = "Trigalia"
    End If
    Flog.writeline "    - ORIGEN RELOJ --> " & Origen

    Select Case tipo_reloj
        Case 1: 'Por default Cargil
            'Fecha MMDDYYYY
            pos1 = 4
            pos2 = 8
            Anio = Year(Date)
            If (CInt(Mid(strreg, pos1, pos2 - pos1 - 2)) > Month(Date)) Then
                Anio = Anio - 1
            End If
            fecha_aux = Mid(strreg, pos1, pos2 - pos1) & CStr(Anio)

            'Hora HHMM
            pos1 = 8
            pos2 = 12
            Hora = Mid(strreg, pos1, pos2 - pos1)
            
            'Nro Tarjeta
            pos1 = 12
            pos2 = 18
            NroTarj_aux = "0" & Mid(strreg, pos1, pos2 - pos1)
            
            'Nro Reloj
            nroreloj = 4
            
            Validar = True
        Case 2: ' Trigalia Formato 2
            If UCase(Mid(strreg, 1, 3)) = "COM" Then
                'Fecha YYMMDD --> convertir a MMDDYYYY
                pos1 = 21
                pos2 = 27
                fecha_aux = Mid(strreg, pos1 + 2, pos2 - pos1 - 4) & Mid(strreg, pos1 + 4, pos2 - pos1 - 4) & "20" & Mid(strreg, pos1, pos2 - pos1 - 4)
                
                'Hora HHMM
                pos1 = 27
                pos2 = 31
                Hora = Mid(strreg, pos1, pos2 - pos1)
                
                'Nro Tarjeta
                pos1 = 13
                pos2 = 17
                NroTarj_aux = "0" & Mid(strreg, pos1, pos2 - pos1)
                
                'Nro Reloj
                nroreloj = 1
                
                Validar = True
            End If
        Case 3: ' Trigalia Formato 1
            If (UCase(Mid(strreg, 2, 1)) <> "A" And UCase(Mid(strreg, 2, 1)) <> "B") Then
                'Fecha MMDDYYYY
                pos1 = 4
                pos2 = 8
                Anio = Year(Date)
                If (CInt(Mid(strreg, pos1, pos2 - pos1 - 2)) > Month(Date)) Then
                    Anio = Anio - 1
                End If
                fecha_aux = Mid(strreg, pos1, pos2 - pos1) & CStr(Anio)
    
                'Hora HHMM
                pos1 = 8
                pos2 = 12
                Hora = Mid(strreg, pos1, pos2 - pos1)
                
                'Nro Tarjeta
                pos1 = 12
                pos2 = 18
                NroTarj_aux = "0" & Mid(strreg, pos1, pos2 - pos1)
                
                'Nro Reloj
                nroreloj = 1
                
                Validar = True
            End If
        Case 4: ' Tehknosur
            'Nro Tarjeta
            pos1 = 1
            pos2 = 6
            NroTarj_aux = "0" & Mid(strreg, pos1, pos2 - pos1)
            
            'Fecha DD/MM/YYYY --> convertir a MMDDYYYY
            pos1 = 7
            pos2 = 17
            fecha_aux = Mid(strreg, pos1 + 3, pos2 - pos1 - 8) & Mid(strreg, pos1, pos2 - pos1 - 8) & Mid(strreg, pos1 + 6, pos2 - pos1 - 6)

            'Hora HHMM
            pos1 = 18
            pos2 = 23
            Hora = Mid(strreg, pos1, pos2 - pos1 - 3) & Mid(strreg, pos1 + 3, pos2 - pos1 - 3)
            
            'Nro Reloj
            nroreloj = 1
            
            If Mid(strreg, 24, 2) = "02" Then
                comedor = True
            End If
            
            Validar = True
            
    End Select
    

'====================================================================
' Validar los parametros Levantados
    
    If Validar Then
        'Busco el Reloj
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "       ****** Error. Reloj no encontrado --> " & nroreloj
            GoTo Fin
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
        
        'Que la fecha sea válida
        Dia = Mid(fecha_aux, 3, 2)
        Mes = Mid(fecha_aux, 1, 2)
        Anio = Mid(fecha_aux, 5, 4)
        If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
            Flog.writeline "       ****** Fecha no válida (MMDDYYYY) --> " & fecha_aux
            GoTo Fin
        Else
            regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
        End If
        
        'Que la hora sea válida
        If Not objFechasHoras.ValidarHora(Hora) Then
            Flog.writeline "       ****** Hora no válida --> " & Hora
            GoTo Fin
        End If
        
        'Que la tarjeta sea numérico
        If Not IsNumeric(NroTarj_aux) Then
            Flog.writeline "       ****** La Tarjeta no es numérica --> " & NroTarj_aux
            GoTo Fin
        Else
            NroTarj = CDbl(NroTarj_aux)
        End If
        
    
    '    Flog.writeline "     ****** Busco que el nro de tarjeta sea válido"
        StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroTarj & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Ternro = objRs!Ternro
        Else
            Flog.writeline "       ****** Error. Empleado no encontrado asociado al Nro. tarjeta: '" & NroTarj & "' y tipo: " & tipotarj & " "
            Flog.writeline "         SQL: " & StrSql
            GoTo Fin
        End If
        
        
        'Que el empleado este activo
        StrSql = "SELECT * FROM empleado where ternro = " & Ternro & " AND empest = -1"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            StrSql = "SELECT * FROM empleado where ternro = " & Ternro
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                Flog.writeline "       ****** Error. El empleado esta inactivo. Legajo --> " & objRs!EmpLeg
            Else
                Flog.writeline "       ****** Error. El empleado no se encuentra. Nro. Tarjeta --> " & NroTarj
            End If
            GoTo Fin
        Else
            Legajo = CStr(objRs!EmpLeg)
            StrSql = "SELECT * FROM his_estructura WHERE ternro = " & Ternro
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & ")"
            StrSql = StrSql & " AND estrnro = " & CInt(NroSuc) & " AND tenro = 1"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Flog.writeline "       ****** Error. El Empleado: " & Legajo & " no pertenece a la estructura: " & NroSuc
                GoTo Fin
            End If
        End If
         
        'Carmen Quintero - 15/05/2015
        Reg_Valida = True
        StrSql = "SELECT relnro FROM gti_rel_estr "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
            'Valido que el reloj sea de control de acceso para el empleado
            StrSql = "SELECT ternro FROM his_estructura H "
            StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
            StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
            StrSql = StrSql & " AND ( h.ternro = " & Ternro
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Reg_Valida = False
                Flog.writeline "    El reloj No está habilitado para el empleado "
            End If
        End If
        'Fin Carmen Quintero - 15/05/2015
         
        If comedor Then
            StrSql = "SELECT * FROM gti_regcome WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
            
                Flog.writeline "               INSERTO REGISTRACION COMEDOR - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
            
                If Reg_Valida Then
                    StrSql = " INSERT INTO gti_regcome(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'I')"
                Else
                    StrSql = " INSERT INTO gti_regcome(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'X')"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Call InsertarWF_Lecturas(Ternro, regfecha)
                
            Else
                Flog.writeline "       ****** Registracion de Comedor ya Existente"
                Flog.writeline "         Error Legajo: " & NroLegajo & " y Reloj: " & codReloj
                Flog.writeline "         Hora: " & Hora & " - Fecha: '" & regfecha & "'"
            End If
        Else
            StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
            
                Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
            
                If Reg_Valida Then
                    StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'I')"
                Else
                    StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'X')"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Call InsertarWF_Lecturas(Ternro, regfecha)
                
            Else
                Flog.writeline "       ****** Registracion ya Existente"
                Flog.writeline "         Error Legajo: " & NroLegajo & " y Reloj: " & codReloj
                Flog.writeline "         Hora: " & Hora & " - Fecha: '" & regfecha & "'"
            End If
        End If
    End If

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub

Private Sub InsertaFormatoCargillGOSC(strreg As String, NroSuc As String)

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim NroTarj As Long
Dim NroTarj_aux As String
Dim descarte  As String
Dim Validar As Boolean
Dim tipo_reloj As Integer
Dim Origen As String
Dim comedor As Boolean
Dim Legajo As String
Dim Reg_Valida As Boolean

    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    
    Validar = False
    
    comedor = False
    
    Flog.writeline "   - Registración --> " & strreg
    
    tipo_reloj = 4 ' Por default Tehknosur
    Origen = "Por defecto Tehknosur"
    If Len(strreg) = 17 Then
        tipo_reloj = 1 ' Cargil
        Origen = "Cargil"
    End If
    If (UCase(Mid(strreg, 1, 3)) = "COM") Then
        tipo_reloj = 2 ' Trigalia Formato 1
        Origen = "Trigalia"
    End If
    If (UCase(Mid(strreg, 1, 1)) = "*") Then
        tipo_reloj = 3 ' Trigalia Formato 2
        Origen = "Trigalia"
    End If
    Flog.writeline "    - ORIGEN RELOJ --> " & Origen

    Select Case tipo_reloj
        Case 1: 'Por default Cargil
            'Fecha MMDDYYYY
            pos1 = 4
            pos2 = 8
            Anio = Year(Date)
            If (CInt(Mid(strreg, pos1, pos2 - pos1 - 2)) > Month(Date)) Then
                Anio = Anio - 1
            End If
            fecha_aux = Mid(strreg, pos1, pos2 - pos1) & CStr(Anio)

            'Hora HHMM
            pos1 = 8
            pos2 = 12
            Hora = Mid(strreg, pos1, pos2 - pos1)
            
            'Nro Tarjeta
            pos1 = 12
            pos2 = 18
            NroTarj_aux = "0" & Mid(strreg, pos1, pos2 - pos1)
            
            'Nro Reloj
            nroreloj = NroSuc
            
            Validar = True
        Case 2: ' Trigalia Formato 2
            If UCase(Mid(strreg, 1, 3)) = "COM" Then
                'Fecha YYMMDD --> convertir a MMDDYYYY
                pos1 = 21
                pos2 = 27
                fecha_aux = Mid(strreg, pos1 + 2, pos2 - pos1 - 4) & Mid(strreg, pos1 + 4, pos2 - pos1 - 4) & "20" & Mid(strreg, pos1, pos2 - pos1 - 4)
                
                'Hora HHMM
                pos1 = 27
                pos2 = 31
                Hora = Mid(strreg, pos1, pos2 - pos1)
                
                'Nro Tarjeta
                pos1 = 13
                pos2 = 17
                NroTarj_aux = "0" & Mid(strreg, pos1, pos2 - pos1)
                
                'Nro Reloj
                nroreloj = NroSuc
                
                Validar = True
            End If
        Case 3: ' Trigalia Formato 1
            If (UCase(Mid(strreg, 2, 1)) <> "A" And UCase(Mid(strreg, 2, 1)) <> "B") Then
                'Fecha MMDDYYYY
                pos1 = 4
                pos2 = 8
                Anio = Year(Date)
                If (CInt(Mid(strreg, pos1, pos2 - pos1 - 2)) > Month(Date)) Then
                    Anio = Anio - 1
                End If
                fecha_aux = Mid(strreg, pos1, pos2 - pos1) & CStr(Anio)
    
                'Hora HHMM
                pos1 = 8
                pos2 = 12
                Hora = Mid(strreg, pos1, pos2 - pos1)
                
                'Nro Tarjeta
                pos1 = 12
                pos2 = 18
                NroTarj_aux = "0" & Mid(strreg, pos1, pos2 - pos1)
                
                'Nro Reloj
                nroreloj = NroSuc
                
                Validar = True
            End If
        Case 4: ' Tehknosur
            'Nro Tarjeta
            pos1 = 1
            pos2 = 6
            NroTarj_aux = "0" & Mid(strreg, pos1, pos2 - pos1)
            
            'Fecha DD/MM/YYYY --> convertir a MMDDYYYY
            pos1 = 7
            pos2 = 17
            fecha_aux = Mid(strreg, pos1 + 3, pos2 - pos1 - 8) & Mid(strreg, pos1, pos2 - pos1 - 8) & Mid(strreg, pos1 + 6, pos2 - pos1 - 6)

            'Hora HHMM
            pos1 = 18
            pos2 = 23
            Hora = Mid(strreg, pos1, pos2 - pos1 - 3) & Mid(strreg, pos1 + 3, pos2 - pos1 - 3)
            
            'Nro Reloj
            nroreloj = NroSuc
            
            If Mid(strreg, 24, 2) = "02" Then
                comedor = True
            End If
            
            Validar = True
            
    End Select
    

'====================================================================
' Validar los parametros Levantados
    
    If Validar Then
        'Busco el Reloj
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "       ****** Error. Reloj no encontrado --> " & nroreloj
            GoTo Fin
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
        
        'Que la fecha sea válida
        Dia = Mid(fecha_aux, 3, 2)
        Mes = Mid(fecha_aux, 1, 2)
        Anio = Mid(fecha_aux, 5, 4)
        If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
            Flog.writeline "       ****** Fecha no válida (MMDDYYYY) --> " & fecha_aux
            GoTo Fin
        Else
            regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
        End If
        
        'Que la hora sea válida
        If Not objFechasHoras.ValidarHora(Hora) Then
            Flog.writeline "       ****** Hora no válida --> " & Hora
            GoTo Fin
        End If
        
        'Que la tarjeta sea numérico
        If Not IsNumeric(NroTarj_aux) Then
            Flog.writeline "       ****** La Tarjeta no es numérica --> " & NroTarj_aux
            GoTo Fin
        Else
            NroTarj = CDbl(NroTarj_aux)
        End If
        
    
    '    Flog.writeline "     ****** Busco que el nro de tarjeta sea válido"
        StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroTarj & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Ternro = objRs!Ternro
        Else
            Flog.writeline "       ****** Error. Empleado no encontrado asociado al Nro. tarjeta: '" & NroTarj & "' y tipo: " & tipotarj & " "
            Flog.writeline "         SQL: " & StrSql
            GoTo Fin
        End If
        
        
        'Que el empleado este activo
        StrSql = "SELECT * FROM empleado where ternro = " & Ternro & " AND empest = -1"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            StrSql = "SELECT * FROM empleado where ternro = " & Ternro
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                Flog.writeline "       ****** Error. El empleado esta inactivo. Legajo --> " & objRs!EmpLeg
            Else
                Flog.writeline "       ****** Error. El empleado no se encuentra. Nro. Tarjeta --> " & NroTarj
            End If
            GoTo Fin
        Else
            Legajo = CStr(objRs!EmpLeg)
            StrSql = "SELECT * FROM his_estructura WHERE ternro = " & Ternro
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & ")"
            StrSql = StrSql & " AND estrnro = " & CInt(NroSuc) & " AND tenro = 1"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Flog.writeline "       ****** Error. El Empleado: " & Legajo & " no pertenece a la estructura: " & NroSuc
                GoTo Fin
            End If
        End If
         
         
        'Carmen Quintero - 15/05/2015
        Reg_Valida = True
        StrSql = "SELECT relnro FROM gti_rel_estr "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
            'Valido que el reloj sea de control de acceso para el empleado
            StrSql = "SELECT ternro FROM his_estructura H "
            StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
            StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
            StrSql = StrSql & " AND ( h.ternro = " & Ternro
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Reg_Valida = False
                Flog.writeline "    El reloj No está habilitado para el empleado "
            End If
        End If
        'Fin Carmen Quintero - 15/05/2015
         
        If comedor Then
            StrSql = "SELECT * FROM gti_regcome WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
            
                Flog.writeline "               INSERTO REGISTRACION COMEDOR - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
                
                If Reg_Valida Then
                    StrSql = " INSERT INTO gti_regcome(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'I')"
                Else
                    StrSql = " INSERT INTO gti_regcome(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'X')"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Call InsertarWF_Lecturas(Ternro, regfecha)
                
            Else
                Flog.writeline "       ****** Registracion de Comedor ya Existente"
                Flog.writeline "         Error Legajo: " & NroLegajo & " y Reloj: " & codReloj
                Flog.writeline "         Hora: " & Hora & " - Fecha: '" & regfecha & "'"
            End If
        Else
            StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
            
                Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
            
                If Reg_Valida Then
                    StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'I')"
                Else
                    StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'X')"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Call InsertarWF_Lecturas(Ternro, regfecha)
                
            Else
                Flog.writeline "       ****** Registracion ya Existente"
                Flog.writeline "         Error Legajo: " & NroLegajo & " y Reloj: " & codReloj
                Flog.writeline "         Hora: " & Hora & " - Fecha: '" & regfecha & "'"
            End If
        End If
    End If

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub


Private Sub InsertaFormato6()

Dim Nroter As Long
Dim codReloj As Integer

Dim rs_TMK As New ADODB.Recordset
Dim rs_GTI_Registracion As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset

codReloj = 1
StrSql = "SELECT * FROM TMK_BAJADA_REG where trim(upper(legajo)) <> 'GUARD'"
OpenRecordset StrSql, rs_TMK

Do While Not rs_TMK.EOF
    
    Nroter = 0

    StrSql = "SELECT * FROM empleado "
    StrSql = StrSql & " WHERE empleg =" & rs_TMK!Legajo
    OpenRecordset StrSql, rs_Empleado

    If Not rs_Empleado.EOF Then
        Nroter = rs_Empleado!Ternro
        
        StrSql = "SELECT * FROM gti_registracion "
        StrSql = StrSql & " WHERE ternro =" & Nroter
        StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_TMK!regfecha)
        StrSql = StrSql & " AND reghora ='" & rs_TMK!Reghora & "'"
        OpenRecordset StrSql, rs_GTI_Registracion
        
        If rs_GTI_Registracion.EOF Then
            StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,regentsal,relnro,regestado,regmanual) VALUES ("
            StrSql = StrSql & Nroter & ","
            StrSql = StrSql & ConvFecha(rs_TMK!regfecha) & ",'"
            StrSql = StrSql & Replace(rs_TMK!Reghora, ":", "") & "','"
            StrSql = StrSql & " ',"
            StrSql = StrSql & codReloj
            StrSql = StrSql & ",'I',"
            StrSql = StrSql & CInt(False) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Call InsertarWF_Lecturas(Nroter, rs_TMK!regfecha)
        Else
            Flog.writeline " esa registracion ya existe. Legajo: " & rs_TMK!Legajo & " Fecha: " & rs_TMK!regfecha & " Hora: " & rs_TMK!Reghora
        End If
        
        'Borro
        StrSql = "DELETE TMK_BAJADA_REG "
        StrSql = StrSql & " WHERE regfecha =" & ConvFecha(rs_TMK!regfecha)
        StrSql = StrSql & " AND reghora ='" & rs_TMK!Reghora & "'"
        StrSql = StrSql & " AND legajo ='" & rs_TMK!Legajo & "'"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        Flog.writeline " No se encontro el legajo " & rs_TMK!Legajo
    End If
    
    rs_TMK.MoveNext
    
Loop

Fin:
    
    If rs_GTI_Registracion.State = adStateOpen Then rs_GTI_Registracion.Close
    If rs_TMK.State = adStateOpen Then rs_TMK.Close
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    
    Set rs_GTI_Registracion = Nothing
    Set rs_TMK = Nothing
    Set rs_Empleado = Nothing
    
End Sub

Private Sub InsertaFormato8()
' ---------------------------------------------------------------------------------------------
' Descripcion: HALLIBURTON
' Autor      : JMH
' Fecha      : 25/01/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Ternro As Long
Dim codReloj As Integer

Dim rs_WC As New ADODB.Recordset
Dim rs_GTI_Registracion As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim HayError As Boolean
Dim tipotarj As Integer

codReloj = 1
StrSql = "SELECT * FROM WC_BAJADA_REG WHERE bpronro = 0"
OpenRecordset StrSql, rs_WC

Do While Not rs_WC.EOF
    
    Ternro = 0
    HayError = False

    Flog.writeline "Busco la hora"
    If Not objFechasHoras.ValidarHora(rs_WC!Reghora) Then
        Flog.writeline " Error Hora: " & rs_WC!Reghora
        HayError = True
    End If
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & rs_WC!relcodext & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
       Flog.writeline "Error. Reloj no encontrado: " & rs_WC!relcodext
       HayError = True
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    Flog.writeline "Busco el nro de tarjeta "
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & rs_WC!EmpLeg & "' AND (hstjfecdes <= " & ConvFecha(rs_WC!regfecha) & ") AND ( (" & ConvFecha(rs_WC!regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
       Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & rs_WC!EmpLeg & ", tipo de tarjeta: " & tipotarj & " , Reloj: " & codReloj
       HayError = True
    End If
    
    Flog.writeline "Busco la entrada/salida"
    If rs_WC!regentsal <> "E" And rs_WC!regentsal <> "S" Then
       Flog.writeline "Error. La entrada/salida debe ser E o S "
       HayError = True
    End If
    
    If HayError = False Then
        'Nroter = rs_Empleado!Ternro
        
        StrSql = "SELECT * FROM gti_registracion "
        StrSql = StrSql & " WHERE ternro =" & Ternro
        StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_WC!regfecha)
        StrSql = StrSql & " AND reghora ='" & rs_WC!Reghora & "'"
        OpenRecordset StrSql, rs_GTI_Registracion
        
        If rs_GTI_Registracion.EOF Then
            StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,regentsal,relnro,regestado, regmanual) VALUES ("
            StrSql = StrSql & Ternro & ","
            StrSql = StrSql & ConvFecha(rs_WC!regfecha) & ",'"
            StrSql = StrSql & Replace(rs_WC!Reghora, ":", "") & "','"
            StrSql = StrSql & rs_WC!regentsal & "',"
            StrSql = StrSql & codReloj
            StrSql = StrSql & ",'" & rs_WC!regestado & "',"
            StrSql = StrSql & rs_WC!regmanual & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            Flog.writeline " esa registracion ya existe. Legajo: " & rs_WC!EmpLeg & " Fecha: " & rs_WC!regfecha & " Hora: " & rs_WC!Reghora
        End If
        
        'Actualizo la tabla poniendo el proceso
        StrSql = "UPDATE WC_BAJADA_REG SET bpronro =" & NroProceso
        StrSql = StrSql & " WHERE empleg = " & rs_WC!EmpLeg
        StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_WC!regfecha)
        StrSql = StrSql & " AND reghora ='" & rs_WC!Reghora & "'"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    rs_WC.MoveNext
    
Loop

Fin:
    
    If rs_GTI_Registracion.State = adStateOpen Then rs_GTI_Registracion.Close
    If rs_WC.State = adStateOpen Then rs_WC.Close
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    
    Set rs_GTI_Registracion = Nothing
    Set rs_WC = Nothing
    Set rs_Empleado = Nothing
    
End Sub

'Martin Ferraro - 12/09/2006 - Modelo 201 Macronet para AGD
'Reloj-Tarjeta-0-fecha-001-hora-E/s (B=Entrada E=Salida)
'4100 tttttttt0dd -mm - yy001hhmmc
'ej
'04100 00008295 005-09-060010812B
'0410000008283005-09-060010812B
'0410000008439005-09-060010813B
Private Sub InsertaFormato9_old(strreg As String)

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Dia As String
Dim Mes As String
Dim Anio As String

    On Error GoTo ME_Local
    
    RegLeidos = RegLeidos + 1
    
    
    'RELOJ------------------------------------------------------------
    nroreloj = Mid(strreg, 1, 5)
    nrorelojtxt = Mid(strreg, 1, 5)
    
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & Trim(nroreloj) & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & Trim(nrorelojtxt) & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    
    'FECHA------------------------------------------------------------
    Dia = Mid(strreg, 15, 2)
    Mes = Mid(strreg, 18, 2)
    Anio = Mid(strreg, 21, 2)
    Anio = "20" & Anio
    Fecha = CDate(Dia & "/" & Mes & "/" & Anio)
    regfecha = Fecha
    
    
    'TARJETA----------------------------------------------------------
    NroLegajo = Mid(strreg, 6, 8)
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & Trim(NroLegajo) & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & Trim(NroLegajo) & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      OpenRecordset StrSql, objRs
      If Not objRs.EOF Then
         Ternro = objRs!Ternro
      Else
         Flog.writeline "Error. Nro de tarjeta: " & NroLegajo & " no encontrado, Tipo de tarjeta: " & tipotarj & " y codigo de reloj:  " & codReloj
         Flog.writeline "SQL: " & StrSql
         InsertaError 1, 33
         Exit Sub
      End If
    End If
      
   
    'HORA-------------------------------------------------------------
    Hora = Mid(strreg, 26, 4)
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    
    'ENTRADA/SALIDA---------------------------------------------------
    entradasalida = Trim(Mid(strreg, 30, 1))
    
    Select Case UCase(entradasalida)
    Case "B":
        entradasalida = "E"
    Case "E":
        entradasalida = "S"
    Case "F":
        Exit Sub
    Case Else
        Flog.writeline "Error. La entrada/salida debe ser E o S "
        Exit Sub
    End Select
    
   
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & Trim(Hora) & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        objConn.Execute StrSql, , adExecuteNoRecords
        Call InsertarWF_Lecturas(Ternro, Fecha)
        
    Else
        Flog.writeline " Registracion ya Existente "
        Flog.writeline " Hora: " & Hora & " - Fecha: " & regfecha
        InsertaError 1, 92
    End If
        
Fin:
Exit Sub
ME_Local:
    HuboError = True
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    GoTo Fin
End Sub



Private Sub InsertaFormato9(strreg As String)
'Martin Ferraro - 12/09/2006 - Modelo 201 Macronet para AGD
'Reloj-Tarjeta-0-fecha-001-hora-E/s (B=Entrada E=Salida)
'4100 tttttttt0dd -mm - yy001hhmmc
'ej
'04100 00008295 005-09-060010812B
'0410000008283005-09-060010812B
'0410000008439005-09-060010813B
'------------------------------------------------------------------
'Se redefinió el formato del archivo de fichadas
'En principio es igual que anes solo que se agregan 2 campos:

'rrrrrtttttttt0dd -mm - yy001hhnnXYZ
'Donde:      rrrrr = Reloj (sector)              tttttttt = tarjeta
'        X = E o S dependiendo si es Entrada o Salida
'        Y = C, S, M dependiendo si es Con contrato, Sin contrato o tiempo Muerto
'        Z = N o A dependiendo si es Normal o Automático
'------------------------------------------------------------------
'Ultimas Modificaciones
'   Modificar el nombre de la biblioteca CLTDTA73 por PD812DTA.
'   La búsqueda que hace en la tabla  FBS4312 (de la biblioteca AGDPRODTA) se debe reemplazar por: PD812DTA.F554312
'   y los campos que se utilizan en el Select debe ser los siguientes:
'------------------------------------------
'Campo Anterior             Campo Nuevo
'------------------------------------------
'$TKCO                       ROKCOO
'$TAN8                       ROAN8
'$TRTO                       ROZDCT
'$TCM01                      ROZRTO
'------------------------------------------------------------------
'------------------------------------------------------------------

'Para Probar Localmente -----------
'Const c_sql_db_F4209 = " F4209 "
'Const c_sql_db_nueordene = " nueordene "
'Const c_sql_db_F4311 = " F4311 "
'Const c_sql_db_F0411 = " F0411 "
'Const c_sql_db_FBS4312 = " FBS4312 "

'Produccion -----------
'Const c_sql_db_F4209 = " CLTDTA73.F4209 "
'Const c_sql_db_nueordene = " MANTECDATO.nueordene "
'Const c_sql_db_F4311 = " CLTDTA73.F4311 "
'Const c_sql_db_F0411 = " CLTDTA73.F0411 "
'Const c_sql_db_FBS4312 = " AGDPRODTA.FBS4312 "
'Nueva produccion
'Const c_sql_db_F4209 = " PD812DTA.F4209 "
'Const c_sql_db_nueordene = " MANTECDATO.nueordene "
'Const c_sql_db_F4311 = " PY812DTA.F4311 "
'Const c_sql_db_F0411 = " PY812DTA.F0411 "
'Const c_sql_db_FBS4312 = " PD812DTA.F554312 "
 
'FGZ - 04/01/2010
'Nueva produccion
Const c_sql_db_F4209 = " PD812DTA.F4209 "
Const c_sql_db_nueordene = " MANTECDATO.nueordene "
Const c_sql_db_F4311 = " PD812DTA.F4311 "
Const c_sql_db_F0411 = " PD812DTA.F0411 "
Const c_sql_db_FBS4312 = " PD812DTA.F554312 "
  
'Nuevo Desarrollo -----------
'Const c_sql_db_F4209 = " CRPDTA73.F4209 "
'Const c_sql_db_nueordene = " MANTECDATO.nueordene "
'Const c_sql_db_F4311 = " CRPDTA73.F4311 "
'Const c_sql_db_F0411 = " CRPDTA73.F0411 "
'Const c_sql_db_FBS4312 = " CRPPRODTA.FBS4312 "


Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Dia As String
Dim Mes As String
Dim Anio As String

'Nuevos Variables
Dim TipoRegistro As String
Dim TipoCierre As String
Dim TipoContrato As String
Dim LineaContrato As String
Dim Contrato As String
Dim Anormalidad As Long
Dim Empresa_Cod As String
Dim Regnro As Long
Dim Empresa As Long
Dim Reg_Valida As Boolean

Dim rs As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Estr_Cod As New ADODB.Recordset
Dim rs_ContJDE As New ADODB.Recordset
Dim rs_Con As New ADODB.Recordset
Dim rs_F4311 As New ADODB.Recordset

    On Error GoTo ME_Local
    
    RegLeidos = RegLeidos + 1
    
    'Inicializo
    TipoContrato = ""
    Contrato = ""
    Anormalidad = 0
    Empresa_Cod = "00000"
    Regnro = 0
    Empresa = 0
    
    Flog.writeline "    Linea " & strreg
    
    
    'RELOJ------------------------------------------------------------
    If IsNumeric(Mid(strreg, 1, 5)) Then
        nroreloj = Mid(strreg, 1, 5)
        nrorelojtxt = Mid(strreg, 1, 5)
    Else
        nroreloj = 0
        nrorelojtxt = Mid(strreg, 1, 5)
    End If
    
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & Trim(nroreloj) & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & Trim(nrorelojtxt) & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    
    'FECHA------------------------------------------------------------
    Dia = Mid(strreg, 15, 2)
    Mes = Mid(strreg, 18, 2)
    Anio = Mid(strreg, 21, 2)
    Anio = "20" & Anio
    Fecha = CDate(Dia & "/" & Mes & "/" & Anio)
    regfecha = Fecha
    
    
    'TARJETA----------------------------------------------------------
    NroLegajo = Mid(strreg, 6, 8)
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & Trim(NroLegajo) & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & Trim(NroLegajo) & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      OpenRecordset StrSql, objRs
      If Not objRs.EOF Then
         Ternro = objRs!Ternro
      Else
         Flog.writeline "Error. Nro de tarjeta: " & NroLegajo & " no encontrado, Tipo de tarjeta: " & tipotarj & " y codigo de reloj:  " & codReloj
         Flog.writeline "SQL: " & StrSql
         InsertaError 1, 33
         Exit Sub
      End If
    End If
      
   
    'HORA-------------------------------------------------------------
    Hora = Mid(strreg, 26, 4)
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    
    'ENTRADA/SALIDA---------------------------------------------------
    entradasalida = Trim(Mid(strreg, 30, 1))
    
    
    'TIPO DE REGISTRO Y TIPO DE CIERRE--------------------------------
    'FGZ - 13/05/2009 - los nuevos campos no siempre van a venir
    If Len(strreg) > 30 Then
        TipoRegistro = Trim(Mid(strreg, 31, 1))
        TipoCierre = Trim(Mid(strreg, 32, 1))
    Else
        TipoRegistro = ""
        TipoCierre = ""
    End If
    
    Select Case UCase(entradasalida)
    Case "B":
        entradasalida = "E"
    Case "E":
        entradasalida = "S"
    Case "F":
        Flog.writeline "Error. La entrada/salida debe ser E o S "
        Exit Sub
    Case Else
        Flog.writeline "Error. La entrada/salida debe ser E o S "
        Exit Sub
    End Select
   
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "  Datos asociados..."
    Flog.writeline Espacios(Tabulador * 2) & " Legajo(Tarjeta)  : " & NroLegajo
    Flog.writeline Espacios(Tabulador * 2) & " Reloj            : " & nrorelojtxt
    Flog.writeline Espacios(Tabulador * 2) & " Fecha            : " & Fecha
    Flog.writeline Espacios(Tabulador * 2) & " Hora             : " & Hora
    Flog.writeline Espacios(Tabulador * 2) & " Entrada/Salida   : " & entradasalida
    Flog.writeline Espacios(Tabulador * 2) & " Tipo de registro : " & TipoRegistro
    Flog.writeline Espacios(Tabulador * 2) & " Tipo de Cierre   : " & TipoCierre
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    'FGZ - 13/05/2009 ------------------------------
    'SI es una registracion de agencia ==>
    If Not EsNulo(TipoRegistro) Then
        Anormalidad = 0
        Select Case TipoRegistro
        Case "C":   'Con Contrato Asociado
            Flog.writeline Espacios(Tabulador * 1) & "  Registracion Con contrato asociado."
            'Al leer una registración, si esta viene asociada a un contrato (Indicado por un estado),
            'la misma se deberá asociar al contrato asociado al legajo
            '   y validar contra JDE que ese contrato este activo para la fecha de la registración
            
            'Busco el contrato asociado al empleado para la fecha de la registracion(me quedo con el primero que encuentre)
            StrSql = "SELECT tipocont, cont FROM his_contratos "
            StrSql = StrSql & " WHERE ternro = " & Ternro
            StrSql = StrSql & " AND (hdesde <= " & ConvFecha(regfecha) & ") AND "
            StrSql = StrSql & " ((" & ConvFecha(regfecha) & " <= hhasta) or (hhasta IS NULL))"
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                TipoContrato = rs!tipocont
                Contrato = rs!cont
                Flog.writeline "    Contrato asignado " & TipoContrato & "-" & Contrato
                Anormalidad = 0
            Else
                Contrato = ""
                TipoContrato = ""
                Anormalidad = 16 'Empleado sin Contrato
                Flog.writeline "    Empleado sin Contrato asignado "
            End If
            
            'Ahora debo validar que
            '   el contrato esta activo en JDE y ese contrato esta ligado a la empresa del reloj
            
            'El reloj está asociado a una empresa y la empresa iene un tipo de codigo JDE asociado que utilizo para buscar en JDE
            '==>
            If Anormalidad = 0 Then
                'Busco la empresa ligada al reloj
            
                StrSql = "SELECT estructura.estrnro FROM gti_rel_estr "
                StrSql = StrSql & " INNER JOIN estructura ON gti_rel_estr.estrnro = estructura.estrnro AND estructura.tenro = 10"
                StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
                OpenRecordset StrSql, rs_Estructura
                If rs_Estructura.EOF Then
                    'El reloj No tiene empresa asociada
                    Anormalidad = 15 'Reloj Sin Empresa Asociada
                    Flog.writeline "    El reloj No tiene empresa asociada "
                Else
                    Empresa = rs_Estructura!estrnro
                    Flog.writeline "    El reloj tiene asociada la empresa " & Empresa
                    
                    'busco el ipo de codigo JDE asociado a la empresa asociada al reloj
                    StrSql = "SELECT * FROM estr_cod WHERE estrnro =" & rs_Estructura!estrnro
                    StrSql = StrSql & " AND tcodnro = 140"
                    OpenRecordset StrSql, rs_Estr_Cod
                    If Not rs_Estr_Cod.EOF Then
                        Empresa_Cod = Left(CStr(rs_Estr_Cod!nrocod), 5)
                        Flog.writeline "    codigo JDE asociado a la empresa asociada al reloj " & Empresa_Cod
                    Else
                        Flog.writeline "    No se encontró el codigo interno para la Empresa"
                        Empresa_Cod = "00000"
                    End If
                End If
            
                'Establecer la conexion a la BD JDE
                StrSql = " SELECT cnnro, cnstring FROM conexion "
                StrSql = StrSql & " ORDER BY cnnro DESC "
                OpenRecordset StrSql, rs_Con
                If rs_Con.EOF Then
                    Flog.writeline Espacios(Tabulador * 0) & "  No se encuentra la conexion con JDE "
                    Anormalidad = 20 'Contrato Sin Validar
                Else
                    On Error Resume Next
                    'Abro la conexion
                    OpenConnection rs_Con!cnstring, objConn2
                    If Err.Number <> 0 Then
                        Flog.writeline Espacios(Tabulador * 0) & "Problemas en la conexion. Debe Configurar bien la conexion con JDE"
                        Anormalidad = 20 'Contrato Sin Validar
                        On Error GoTo ME_Local
                    Else
                        Flog.writeline Espacios(Tabulador * 1) & "  Validando contrato.."
                        On Error GoTo ME_Local
                        
                        'Validar contrato
                        'StrSql = "SELECT * FROM " & c_sql_db_F4209
                        'StrSql = StrSql & " WHERE HOKCOO = '" & Empresa_Cod & "'"
                        'StrSql = StrSql & " AND HODCTO = '" & UCase(TipoContrato) & "'"
                        'StrSql = StrSql & " AND HODOCO = " & Contrato
                        'StrSql = StrSql & " AND HOASTS = '3N' "
                        
                        StrSql = "SELECT * FROM " & c_sql_db_F4209
                        StrSql = StrSql & " WHERE HOKCOO = '" & Empresa_Cod & "'"
                        StrSql = StrSql & " AND HOASTS = '3N' "
                        StrSql = StrSql & " AND HODCTO = '" & UCase(TipoContrato) & "'"
                        StrSql = StrSql & " AND HODOCO = " & Contrato
                        OpenRecordsetWithConn StrSql, rs_ContJDE, objConn2
                        If rs_ContJDE.EOF Then
                            Anormalidad = 17 'Sin Contrato
                            Flog.writeline Espacios(Tabulador * 2) & "  No se encuentra Contrato."
                            Flog.writeline StrSql
                        Else
                            'En caso de existir el contrato, se deberá establecer un control adicional
                            '   de Porcentaje de Recepción
                            'archivo F4311(PDDCTO, PDKCOO, PDDOCO, PDLNID)
                            'PDLNID ' CONTROL DE VALORES (es necesario que el valor de línea a validar sea multiplicado por 1000)
                            'PDLNID = 1000 ' 1
                            'PDLNID = 2000 ' 2
                            'PDLNID = 3000 ' 3
                            'PDLNID = 4000 ' 4
                                                        
                            'Select * From CLTDTA73.F4311
                            'Where PDKCOO = '00001'
                            'And PDDCTO = NP', 'N3', 'NX', segun se tenga que validar
                            'And PDLNID = Nro de linea según se este por validar
                            'And PDDOCO = Nro. de Contrato a validar
                            '
                            'SI PDUOPN = 0
                            '    SI PDUREC < PDUORG
                            '        Anormalidad "CONTRATO CANCELADO"
                            '    Else
                            '        Anormalidad "CONTRATO CERRADO"
                            '    End If
                            'Else
                            'NO GENERA ANORMALIDAD
                            'End If
                            
                            Flog.writeline Espacios(Tabulador * 1) & "  Validando linea y porcentaje del contrato."
                            StrSql = "SELECT PDLNID,PDUOPN,PDUREC,PDUORG FROM " & c_sql_db_F4311
                            StrSql = StrSql & " WHERE PDKCOO = '" & Empresa_Cod & "'"
                            StrSql = StrSql & " AND PDDCTO = '" & UCase(TipoContrato) & "'"
                            StrSql = StrSql & " AND PDDOCO = " & Contrato
                            StrSql = StrSql & " AND (PDLNID = 1000 OR PDLNID = 2000 OR PDLNID = 3000 OR PDLNID = 4000) "
                            OpenRecordsetWithConn StrSql, rs_F4311, objConn2
                            If rs_F4311.EOF Then
                                Anormalidad = 17 'Sin Contrato
                                LineaContrato = ""
                                Flog.writeline Espacios(Tabulador * 2) & "  Contrato no tiene Linea asociada."
                            Else
                                
                                LineaContrato = CStr(CLng(rs_F4311!PDLNID / 1000))
                                If rs_F4311!PDUOPN = 0 Then
                                    If rs_F4311!PDUREC < rs_F4311!PDUORG Then
                                        Anormalidad = 18 'CONTRATO CANCELADO
                                        Flog.writeline Espacios(Tabulador * 2) & "  Contrato Cancelado."
                                    Else
                                        Anormalidad = 19 'CONTRATO CERRADO
                                        Flog.writeline Espacios(Tabulador * 2) & "  Contrato Cerrado."
                                    End If
                                Else
                                    'NO GENERA ANORMALIDAD
                                    Anormalidad = 0
                                    Flog.writeline Espacios(Tabulador * 2) & "  Contrato VALIDO."
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                'No tiene sentido revisar mas dado que el empleado no tiene contrato
                Flog.writeline Espacios(Tabulador * 1) & "  Empleado no tiene contrato."
            End If
        Case "S":   'Sin Contrato Asociado (Por Orden de Trabajo)
            LineaContrato = ""
            Contrato = ""
            TipoContrato = ""
            Anormalidad = 20 'Sin Contrato u Orden de Trabajo Asociado
            Flog.writeline Espacios(Tabulador * 1) & "  Registracion Sin Contrato u Orden de Trabajo Asociado."
        Case "M":   'Tiempo Muerto
            LineaContrato = ""
            Contrato = ""
            TipoContrato = ""
            Anormalidad = 20 'Sin Contrato u Orden de Trabajo Asociado
            Flog.writeline Espacios(Tabulador * 1) & "  Tiempo Muerto. Registracion Sin Contrato u Orden de Trabajo Asociado."
        Case Else:
            Flog.writeline Espacios(Tabulador * 1) & "  Tipo de registro incorrecto " & TipoRegistro
        End Select
        
        
        StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & Trim(Hora) & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            If Reg_Valida Then
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
            Else
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Inserto los datos complementarios
            Regnro = getLastIdentity(objConn, "gti_registracion")
            Flog.writeline Espacios(Tabulador * 1) & "  Registracion insertada " & Regnro
            
            Flog.writeline Espacios(Tabulador * 1) & "  insertando complemento... "
            StrSql = " INSERT INTO gti_reg_comp(regnro,tiporeg,tipocierre"
            If Not (EsNulo(Contrato) And EsNulo(TipoContrato)) Then
                StrSql = StrSql & ",tipocont,cont"
            End If
            StrSql = StrSql & ", linea, empresa,empresacod,normnro "
            'If Anormalidad <> 0 Then
            '    StrSql = StrSql & ",normnro"
            'End If
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & Regnro
            StrSql = StrSql & "," & "'" & TipoRegistro & "'"
            StrSql = StrSql & "," & "'" & TipoCierre & "'"
            If Not (EsNulo(Contrato) And EsNulo(TipoContrato)) Then
                StrSql = StrSql & "," & "'" & TipoContrato & "'"
                StrSql = StrSql & "," & Contrato
            End If
            StrSql = StrSql & ",'" & LineaContrato & "'"
            StrSql = StrSql & "," & Empresa
            StrSql = StrSql & "," & "'" & Empresa_Cod & "'"
            'If Anormalidad <> 0 Then
                StrSql = StrSql & "," & Anormalidad
            'End If
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline Espacios(Tabulador * 1) & "  complemento insertado."
            
            
            Call InsertarWF_Lecturas(Ternro, Fecha)
            
        Else
            Flog.writeline " Registracion ya Existente "
            Flog.writeline " Hora: " & Hora & " - Fecha: " & regfecha
            InsertaError 1, 92
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & "  Tipo de registro NULO. Registracion estandar."
        StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & Trim(Hora) & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            If Reg_Valida Then
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
            Else
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            Call InsertarWF_Lecturas(Ternro, Fecha)
            
        Else
            Flog.writeline " Registracion ya Existente "
            Flog.writeline " Hora: " & Hora & " - Fecha: " & regfecha
            InsertaError 1, 92
        End If
    End If
    'FGZ - 13/05/2009 ------------------------------
        
        
    'Cierro y libero todo
    If objConn2.State = adStateOpen Then objConn2.Close
    If rs.State = adStateOpen Then rs.Close
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    If rs_Estr_Cod.State = adStateOpen Then rs_Estr_Cod.Close
    If rs_ContJDE.State = adStateOpen Then rs_ContJDE.Close
    If rs_Con.State = adStateOpen Then rs_Con.Close
    If rs_F4311.State = adStateOpen Then rs_F4311.Close
    
Fin:
Exit Sub
ME_Local:
    HuboError = True
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    GoTo Fin
End Sub




Private Sub InsertaFormato10(strreg As String)
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim pos1 As Byte
Dim pos2 As Byte
'Dim codReloj As Integer
'Dim TipoTarj As Integer
'Dim NroTarj As Integer
Dim codReloj As Long
Dim tipotarj As Long
Dim NroTarj As Long

Dim NroTarj_aux As String
Dim descarte  As String
Dim Reg_Valida As Boolean

    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    
    Flog.writeline "   - Registración --> " & strreg
    'Descartar
    pos1 = 1
    pos2 = 3
    descarte = Mid(strreg, pos1, pos2 - pos1)
    
    'Fecha YYMMDD
    pos1 = 3
    pos2 = 9
    fecha_aux = Mid(strreg, pos1, pos2 - pos1)
    
    'Hora HHMM
    pos1 = 9
    pos2 = 13
    Hora = Mid(strreg, pos1, pos2 - pos1)
    
    'Entrada/Salida
    pos1 = 13
    pos2 = 15
    entradasalida = Mid(strreg, pos1, pos2 - pos1)
    
    'Nro Tarjeta
    pos1 = 15
    'pos2 = 21
    pos2 = 23
    NroTarj_aux = Mid(strreg, pos1, pos2 - pos1)
    
    'Legajo
    'pos1 = 21
    pos1 = 23
    NroLegajo = Mid(strreg, pos1, 5)
    

'====================================================================
' Validar los parametros Levantados
    
    'Busco el Reloj que este definido como por defecto
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE reldefault = -1"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro un Reloj definido por defecto. SQL --> " & StrSql
        InsertaError 1, 32
        GoTo Fin
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    'Que la fecha sea válida
    Dia = Mid(fecha_aux, 5, 2)
    Mes = Mid(fecha_aux, 3, 2)
    Anio = "20" & Mid(fecha_aux, 1, 2)
    If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
        Flog.writeline "       ****** Fecha no válida --> " & fecha_aux
        InsertaError 2, 4
        GoTo Fin
    Else
        regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline "       ****** Hora no válida --> " & Hora
        InsertaError 3, 38
        GoTo Fin
    End If
    
    'Entrada/Salida sea válido
    If UCase(entradasalida) <> "EN" And UCase(entradasalida) <> "SA" Then
        Flog.writeline "       ****** Entrada/Salida no válido --> " & entradasalida & ". Debe ser EN o SA"
        InsertaError 4, 7
        GoTo Fin
    Else
        entradasalida = Mid(entradasalida, 1, 1)
    End If
    
    'Que la tarjeta sea numérico
    If Not IsNumeric(NroTarj_aux) Then
        Flog.writeline "       ****** La Tarjeta no es numérica --> " & NroTarj_aux
        InsertaError 5, 3
        GoTo Fin
    Else
        'NroTarj = CDbl(NroTarj_aux)
        NroTarj = CLng(NroTarj_aux)
    End If
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & NroLegajo
        InsertaError 6, 8
        GoTo Fin
    End If
    
    
'    Flog.writeline "     ****** Busco que el nro de tarjeta sea válido"
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroTarj & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Ternro = objRs!Ternro
    Else
        Flog.writeline "       ****** Error. Tarjeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " y Nro. tarjeta: " & NroTarj
        Flog.writeline "         SQL: " & StrSql
        InsertaError 1, 33
        GoTo Fin
    End If
    
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        
    Else
        Flog.writeline "       ****** Registracion ya Existente"
        Flog.writeline "         Error Legajo: " & NroLegajo & " y Reloj: " & codReloj
        Flog.writeline "         Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub

Private Sub InsertaFormato11(strreg As String)

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim NroReloj_aux As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim NroTarj As Integer
Dim NroTarj_aux As String
Dim descarte  As String
Dim Reg_Valida As Boolean

    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    Flog.writeline "   - Registración --> " & strreg
    
    'Legajo
    pos1 = 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    
    'Fecha DD/MM/YYYY
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    fecha_aux = Mid(strreg, pos1, pos2 - pos1)
    
    'Hora HH:MM
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    Hora = Mid(strreg, pos1, pos2 - pos1)
    Hora = Replace(Hora, ":", "")
    
    'Nro reloj
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    NroReloj_aux = Mid(strreg, pos1, pos2 - pos1)
    
    
    'Descartar (Entrada/Salida)
    pos1 = pos2 + 1
    pos2 = Len(strreg)
    descarte = Mid(strreg, pos1, pos2)
    

'====================================================================
' Validar los parametros Levantados
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & NroLegajo
        InsertaError 1, 8
        GoTo Fin
    Else
        Ternro = objRs!Ternro
    End If
    
    'Que la fecha sea válida
    Dia = Mid(fecha_aux, 1, 2)
    Mes = Mid(fecha_aux, 4, 2)
    Anio = Mid(fecha_aux, 7, 4)
    If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
        Flog.writeline "       ****** Fecha no válida --> " & fecha_aux
        InsertaError 2, 4
        GoTo Fin
    Else
        regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline "       ****** Hora no válida --> " & Hora
        InsertaError 3, 38
        GoTo Fin
    End If
    
    'Busco el Reloj
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & NroReloj_aux & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el Reloj. SQL --> " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Reloj: " & codReloj
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "',''," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        
    Else
        Flog.writeline "       ****** Registracion ya Existente"
        Flog.writeline "         Error Legajo: " & NroLegajo & " y Reloj: " & codReloj
        Flog.writeline "         Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub

Private Sub InsertaFormato7(strreg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato para CARSA.
' Autor      : FGZ
' Fecha      : 05/09/2005
' Ultima Mod.: 29/09/2005
' Descripcion: Se le agregó el reloj(nombre de la puerta)
' ---------------------------------------------------------------------------------------------
Dim Datos
Dim NroLegajo As String
Dim Ternro As Long

Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim codReloj As Long
Dim tipotarj As Integer


Dim FechaYHora As String    'dd/mm/yyyy  hh:mm:ss
Dim tarjeta
Dim Puerta
Dim Reg_Valida As Boolean

Dim rs As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset

    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1

    Datos = Split(strreg, separador)
    FechaYHora = Trim(Datos(29))
    tarjeta = Mid(Datos(36), 2, Len(Datos(36)) - 2)
    Puerta = Trim(Datos(41))
        
    regfecha = C_Date(Left(FechaYHora, 10))
    Hora = Trim(Mid(FechaYHora, 11, 7))
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 5, 38
        Exit Sub
    End If
    
    'Busco el nro de tarjeta
    StrSql = "SELECT ternro FROM gti_histarjeta "
    StrSql = StrSql & " WHERE hstjnrotar = " & tarjeta
    StrSql = StrSql & " AND (hstjfecdes <= " & ConvFecha(regfecha) & ")"
    StrSql = StrSql & " AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) "
    StrSql = StrSql & " OR ( hstjfechas is null ))"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Ternro = rs!Ternro
       
        StrSql = "SELECT empleg FROM empleado "
        StrSql = StrSql & " WHERE ternro = " & Ternro
        If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
        OpenRecordset StrSql, rs_Empleado
        If Not rs_Empleado.EOF Then
            NroLegajo = rs_Empleado!EmpLeg
        Else
            Flog.writeline Espacios(Tabulador * 1) & " no existe el tercero " & Ternro
            InsertaError 13, 33
            Exit Sub
        End If
    Else
        Flog.writeline Espacios(Tabulador * 1) & " Error. Tarjeta " & tarjeta & " inexistente o inactiva "
        Flog.writeline "SQL: " & StrSql
        InsertaError 13, 33
        Exit Sub
    End If
    
    If InStr(1, UCase(Trim(Puerta)), "ENTRADA") Then
        entradasalida = "E"
    Else
        entradasalida = "S"
    End If
    
    'StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & Mid(Puerta, 2, Len(Puerta) - 2) & "'"
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE upper(reldext) = '" & UCase(Trim(Mid(Puerta, 2, Len(Puerta) - 2))) & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Error. No se encontro el Reloj: " & nroreloj
        Flog.writeline "SQL: " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = IIf(Not EsNulo(objRs!relnro), objRs!relnro, 0)
        tipotarj = IIf(Not EsNulo(objRs!tptrnro), objRs!tptrnro, 0)
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    'Procesando
    Debug.Print "Legajo: " & NroLegajo & " - E/S: " & entradasalida & "-  Fecha y Hora: " & FechaYHora & " - Puerta: " & Puerta
    Flog.writeline Espacios(Tabulador * 1) & "Legajo: " & NroLegajo & " - E/S: " & entradasalida & "-  Fecha y Hora: " & FechaYHora & " - Puerta: " & Puerta
    
    'Busco la registracion
    StrSql = "SELECT * FROM gti_registracion "
    StrSql = StrSql & " WHERE regfecha = " & ConvFecha(regfecha)
    StrSql = StrSql & " AND reghora = '" & Hora & "'"
    StrSql = StrSql & " AND ternro = " & Ternro
    StrSql = StrSql & " AND regentsal = '" & entradasalida & "'"
    StrSql = StrSql & " AND relnro = " & codReloj
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If rs.EOF Then
        'Flog.writeline Espacios(Tabulador * 1) & "Insertando registracion - " & NroLegajo & "  ;  '" & RegFecha & "'    ;    " & Hora
         If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES ("
            StrSql = StrSql & Ternro & ","
            StrSql = StrSql & crpNro & ","
            StrSql = StrSql & ConvFecha(regfecha) & ","
            StrSql = StrSql & "'" & Hora & "',"
            StrSql = StrSql & "'" & entradasalida & "',"
            StrSql = StrSql & codReloj & ","
            StrSql = StrSql & "'I'"
            StrSql = StrSql & ")"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        'Flog.writeline Espacios(Tabulador * 1) & "Insertada"
        Call InsertarWF_Lecturas(Ternro, regfecha)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Registracion ya Existente"
        Flog.writeline Espacios(Tabulador * 1) & " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        'InsertaError 1, 92
    End If
        
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
Set rs_Empleado = Nothing

Exit Sub
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    Set rs_Empleado = Nothing
End Sub

Private Sub InsertaFormato4(strreg As String)
'-------------------------------------------------------------------
'FGZ - 22/11-2005
'Formato:
'   Legajo, fecha, hora, nro de reloj, E/S
'Ejemplo:
'   000001, 22/11/2005, 1000, 0010, E
'   000001, 22/11/2005, 1800, 0010, S
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean

    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Legajo:  " & NroLegajo
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    regfecha = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Fecha:  " & regfecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    nroreloj = Mid(strreg, pos1, pos2 - pos1)
    nrorelojtxt = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. No se encontro el Reloj: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    

    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = IIf(UCase(Trim(Mid(strreg, pos1))) = "E", "E", "S")
       
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      'StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & nroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      'OpenRecordset StrSql, objRs
      'If Not objRs.EOF Then
      '   Ternro = objRs!Ternro
      'Else
         Flog.writeline "Error. No se encontro la terjeta para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj: " & codReloj
         InsertaError 1, 33
         Exit Sub
      'End If
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
        
End Sub




Public Function EsNulo(ByVal Objeto) As Boolean
    If IsNull(Objeto) Then
        EsNulo = True
    Else
        If UCase(Objeto) = "NULL" Or UCase(Objeto) = "" Then
            EsNulo = True
        Else
            EsNulo = False
        End If
    End If
End Function

Public Function Espacios(ByVal Cantidad As Integer) As String
    Espacios = Space(Cantidad)
End Function

Private Sub InsertaFormato13()
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato para ActionLine.
' Autor      : Lisandro Moro
' Fecha      : 05/09/2008
' Ultima Mod.: 29/09/2008
' Descripcion: Lee de la tabla cta_emp_dist
' ---------------------------------------------------------------------------------------------

Dim Nroter As Long
Dim EmpLeg As Long
'Dim codReloj As Integer

Dim rs_AL As New ADODB.Recordset
Dim rs_GTI_Registracion As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset

'codReloj = 1
StrSql = "SELECT * FROM cta_emp_dist WHERE aprob = -1 AND bpronro is null "
OpenRecordset StrSql, rs_AL

Do While Not rs_AL.EOF
    
    Nroter = 0
    EmpLeg = 0

    StrSql = "SELECT * FROM empleado "
    StrSql = StrSql & " WHERE ternro =" & rs_AL!Ternro
    OpenRecordset StrSql, rs_Empleado

    If Not rs_Empleado.EOF Then
        Nroter = rs_Empleado!Ternro
        EmpLeg = rs_Empleado!EmpLeg
        
        
        'Entrada
        StrSql = "SELECT * FROM gti_registracion "
        StrSql = StrSql & " WHERE ternro =" & Nroter
        StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_AL!Fecha)
        StrSql = StrSql & " AND reghora ='" & rs_AL!hordesde & "'"
        StrSql = StrSql & " AND regentsal = 'E' "
        OpenRecordset StrSql, rs_GTI_Registracion
        
        If rs_GTI_Registracion.EOF Then
            StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,regentsal,relnro,regestado,regmanual) VALUES ("
            StrSql = StrSql & Nroter
            StrSql = StrSql & "," & ConvFecha(rs_AL!Fecha)
            StrSql = StrSql & ",'" & rs_AL!hordesde & "'"
            StrSql = StrSql & ",'E' "
            StrSql = StrSql & ", null "
            StrSql = StrSql & ",'I',"
            StrSql = StrSql & CInt(False) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            Flog.writeline " esa registracion ya existe. Legajo: " & rs_AL!Legajo & " Fecha: " & rs_AL!Fecha & " Hora: " & rs_AL!Hora & " : Entrada"
        End If
        
        
        'Salida
        StrSql = "SELECT * FROM gti_registracion "
        StrSql = StrSql & " WHERE ternro =" & Nroter
        StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_AL!Fecha)
        StrSql = StrSql & " AND reghora ='" & rs_AL!horhasta & "'"
        StrSql = StrSql & " AND regentsal = 'S' "
        OpenRecordset StrSql, rs_GTI_Registracion
        
        If rs_GTI_Registracion.EOF Then
            StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,regentsal,relnro,regestado,regmanual) VALUES ("
            StrSql = StrSql & Nroter
            StrSql = StrSql & "," & ConvFecha(rs_AL!Fecha)
            StrSql = StrSql & ",'" & rs_AL!horhasta & "'"
            StrSql = StrSql & ",'S' "
            StrSql = StrSql & ", null "
            StrSql = StrSql & ",'I',"
            StrSql = StrSql & CInt(False) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            Flog.writeline " esa registracion ya existe. Legajo: " & rs_AL!Legajo & " Fecha: " & rs_AL!Fecha & " Hora: " & rs_AL!Hora & " : Entrada"
        End If
        
        'Marco como pocesado
        StrSql = " UPDATE cta_emp_dist SET bpronro  = " & NroProceso
        StrSql = StrSql & " WHERE aprob = -1 "
        StrSql = StrSql & " AND ternro = " & Nroter
        StrSql = StrSql & " AND fecha = " & ConvFecha(rs_AL!Fecha)
        StrSql = StrSql & " AND hordesde = '" & rs_AL!hordesde & "'"
        StrSql = StrSql & " AND horhasta = '" & rs_AL!horhasta & "'"
        objConn.Execute StrSql, , adExecuteNoRecords

    Else
        Flog.writeline " No se encontro el legajo " & EmpLeg
    End If
    
    rs_AL.MoveNext
    
Loop

Fin:
    
    If rs_GTI_Registracion.State = adStateOpen Then rs_GTI_Registracion.Close
    If rs_AL.State = adStateOpen Then rs_AL.Close
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    
    Set rs_GTI_Registracion = Nothing
    Set rs_AL = Nothing
    Set rs_Empleado = Nothing
    
End Sub

Private Sub InsertaFormato14(strreg As String)
'------------------------------------------------------------------
'Es necesario un modelo que soporte la lectura de registraciones desde
'un archivo con la siguiente especificacion:
'Linea Ejemplo:      5213 , TESMA,    0    ,    999    , 15/04/2008, 08:30, INI0000002,    8DF6
'Detalle de Campos: LoteId, Reloj, CodRegis, CodRegisEx,    Fecha  , Hora ,   Tarjeta , Checksum
'------------------------------------------------------------------
Dim reloj As String
Dim CodRegis As Long
Dim Fecha As Date
Dim Hora As String
Dim tarjeta As String
Dim Lote As String
Dim Ternro As Long
Dim NroLegajo As String
Dim codReloj As String
Dim tipotarj As Long
Dim entradasalida As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim Reg_Valida As Boolean


    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    Flog.writeline "   - Registración --> " & strreg
    
    pos1 = 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    
    'Reloj
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    reloj = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    reloj = Trim(reloj)
    
    'CodRegis
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    CodRegis = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    CodRegis = Trim(CodRegis)
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    
    'Fecha
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    Fecha = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    Fecha = Trim(Fecha)
    
    'Hora
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    Hora = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    Hora = Trim(Hora)
    
    'Tarjeta
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    tarjeta = Replace(Mid(strreg, pos1, pos2 - pos1), """", "")
    tarjeta = Trim(tarjeta)
    
'====================================================================
' Validar los parametros Levantados
    
    Lote = UCase(Left(tarjeta, 3))
    If Lote = "INI" Then
        Flog.writeline "   - No se toma la registracion (Inicio de Lote)"
    Else
        If Lote = "FIN" Then
            Flog.writeline "   - No se toma la registracion (Fin de Lote)"
        Else
            'Busco nro de legajo
            StrSql = "SELECT Ternro FROM gti_histarjeta "
            StrSql = StrSql & "WHERE hstjnrotar = " & tarjeta & " "
            StrSql = StrSql & "AND (hstjfecdes <= " & ConvFecha(Fecha) & ") "
            StrSql = StrSql & "AND ((" & ConvFecha(Fecha) & " <= hstjfechas) OR (hstjfechas is null)) "
            OpenRecordset StrSql, objRs
            Ternro = 0
            If objRs.EOF Then
                Flog.writeline "       ****** No se encontro el empleado con tarjeta --> " & tarjeta
                GoTo Fin
            Else
                Ternro = objRs!Ternro
            End If
            
            'Numero de Legajo
            NroLegajo = 0
            If Ternro <> 0 Then
                StrSql = "SELECT empleg FROM empleado "
                StrSql = StrSql & "WHERE ternro = " & Ternro
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    NroLegajo = objRs!EmpLeg
                End If
            End If
            
            'Que la hora sea válida
            If Not objFechasHoras.ValidarHora(Hora) Then
                Flog.writeline "       ****** Hora no válida --> " & Hora
                GoTo Fin
            End If
            
            'Mira si es entrada o salida
            If CodRegis Mod 2 = 0 Then
                entradasalida = "E"
            Else
                entradasalida = "S"
            End If
            
            'Busco el Reloj
            StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE reldabr = '" & reloj & "'"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Flog.writeline "       ****** No se encontro el Reloj. SQL --> " & StrSql
                Exit Sub
            Else
                codReloj = objRs!relnro
                tipotarj = objRs!tptrnro
            End If
            
            'Carmen Quintero - 15/05/2015
            Reg_Valida = True
            StrSql = "SELECT relnro FROM gti_rel_estr "
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
                'Valido que el reloj sea de control de acceso para el empleado
                StrSql = "SELECT ternro FROM his_estructura H "
                StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
                StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
                StrSql = StrSql & " AND ( h.ternro = " & Ternro
                StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(Fecha) & "))"
                OpenRecordset StrSql, objRs
                If objRs.EOF Then
                    Reg_Valida = False
                    Flog.writeline "    El reloj No está habilitado para el empleado "
                End If
            End If
            'Fin Carmen Quintero - 15/05/2015
            
             StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & Fecha & "'  ; Hora: " & Hora & "  ; Nro. Reloj: " & codReloj
                If Reg_Valida Then
                    StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES "
                    StrSql = StrSql & "(" & Ternro
                    StrSql = StrSql & "," & crpNro
                    StrSql = StrSql & "," & ConvFecha(Fecha)
                    StrSql = StrSql & ",'" & Hora & "'"
                    StrSql = StrSql & ",'" & entradasalida & "'"
                    StrSql = StrSql & "," & codReloj
                    StrSql = StrSql & ",'I')"
                Else
                    StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                    Ternro & "," & crpNro & ", " & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                
                Call InsertarWF_Lecturas(Ternro, Fecha)
            Else
                Flog.writeline "       ****** Registracion ya Existente"
                Flog.writeline "         Error Legajo: " & NroLegajo & " y Reloj: " & codReloj
                Flog.writeline "         Hora: " & Hora & " - Fecha: '" & Fecha & "'"
                InsertaError 1, 92
            End If
        End If
    End If

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub

Private Sub InsertaFormato15(strreg As String)
'------------------------------------------------------------------
'Es necesario un modelo que soporte la lectura de registraciones desde
'un archivo con la siguiente especificacion:
'Linea Ejemplo:     000000002029200302271510EAE
'Linea Ejemplo:     000000002029     20030227       1510             E          AE
'Detalle de Campos:    Legajo    Fecha(aaaammdd) Hora(hhmm) E=Entrada/S=Salida Grupo
'Ultima Modificacion:
'               FGZ - 07/09/2009 - la longitud del nro de legajo (tarjeta) pasó a ser 11
'                                       y por ende se desplaza todo el resto de la linea
'------------------------------------------------------------------
Dim reloj As String
Dim Fecha As Date
Dim FecD As String
Dim FecM As String
Dim FecA As String
Dim Hora As String
Dim Ternro As Long
Dim NroLegajo As String
Dim entradasalida As String
Dim codReloj As Long
Dim Reg_Valida As Boolean

    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    Flog.writeline "   - Registración --> " & strreg
    
'    'Legajo
'    NroLegajo = Mid(strReg, 1, 12)
'
'    'Fecha(aaaammdd)
'    FecA = Mid(strReg, 13, 4)
'    FecM = Mid(strReg, 17, 2)
'    FecD = Mid(strReg, 19, 2)
'    Fecha = FecD & "/" & FecM & "/" & FecA
'
'    'Hora(hhmm)
'    Hora = Mid(strReg, 21, 2) & ":" & Mid(strReg, 23, 2)
'
'    'E=Entrada/S=Salida
'    EntradaSalida = Mid(strReg, 25, 1)
    
'FGZ - 07/09/2009 - la longitud del nro de legajo (tarjeta) pasó a ser 11 y por ende se desplaza todo el resto de la linea
    NroLegajo = Mid(strreg, 1, 11)
    'Fecha(aaaammdd)
    FecA = Mid(strreg, 12, 4)
    FecM = Mid(strreg, 16, 2)
    FecD = Mid(strreg, 18, 2)
    Fecha = FecD & "/" & FecM & "/" & FecA
    'Hora(hhmm)
    Hora = Mid(strreg, 20, 2) & ":" & Mid(strreg, 22, 2)
    'E=Entrada/S=Salida
    entradasalida = Mid(strreg, 24, 1)
    
    
'====================================================================
' Validar los parametros Levantados

    'Numero de Legajo
    StrSql = "SELECT Ternro FROM empleado "
    StrSql = StrSql & "WHERE empleg = " & NroLegajo & " "
    OpenRecordset StrSql, objRs
    Ternro = 0
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el empleado con Legajo --> " & NroLegajo
        GoTo Fin
    Else
        Ternro = objRs!Ternro
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline "       ****** Hora no válida --> " & Hora
        GoTo Fin
    End If
    
    'Mira si es entrada o salida
    If entradasalida <> "E" And entradasalida <> "S" Then
        Flog.writeline "       ****** La registración no especifíca si es de entrada o de Salida."
        GoTo Fin
    End If
    
    
    'FGZ - 07/09/2009 - Agregué esto
    StrSql = "SELECT relnro FROM gti_reloj"
    StrSql = StrSql & " ORDER BY relnro"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        codReloj = objRs!relnro
    Else
        codReloj = 0
    End If
    
    If codReloj <> 0 Then
    
        'Carmen Quintero - 15/05/2015
        Reg_Valida = True
        StrSql = "SELECT relnro FROM gti_rel_estr "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
            'Valido que el reloj sea de control de acceso para el empleado
            StrSql = "SELECT ternro FROM his_estructura H "
            StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
            StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
            StrSql = StrSql & " AND ( h.ternro = " & Ternro
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(Fecha) & "))"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Reg_Valida = False
                Flog.writeline "    El reloj No está habilitado para el empleado "
            End If
        End If
        'Fin Carmen Quintero - 15/05/2015
    
        StrSql = " SELECT * FROM gti_registracion WHERE"
        StrSql = StrSql & " regfecha = " & ConvFecha(Fecha)
        StrSql = StrSql & " AND reghora = '" & Hora & "'"
        StrSql = StrSql & " AND ternro = " & Ternro
        StrSql = StrSql & " AND regentsal = '" & entradasalida & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & Fecha & "'  ; Hora: " & Hora & " ; Entrada-Salida (E-S): " & entradasalida
            
            If Reg_Valida Then
                StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,regentsal,relnro,regestado) VALUES "
                StrSql = StrSql & "(" & Ternro
                StrSql = StrSql & "," & ConvFecha(Fecha)
                StrSql = StrSql & ",'" & Hora & "'"
                StrSql = StrSql & ",'" & entradasalida & "'"
                StrSql = StrSql & "," & codReloj
                StrSql = StrSql & ",'I')"
            Else
               StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
 
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            Call InsertarWF_Lecturas(Ternro, Fecha)
        Else
            Flog.writeline "       ****** Registracion ya Existente"
            Flog.writeline "         Error Legajo: " & NroLegajo & " - (E-S): " & entradasalida
            Flog.writeline "         Hora: " & Hora & " - Fecha: '" & Fecha & "'"
            InsertaError 1, 92
        End If
    Else
            Flog.writeline "       No hay ningun reloj configurado. No se inserta la registracion."
    End If
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub


Private Sub InsertaFormato16(strreg As String)
'-------------------------------------------------------------------
'FGZ - 10/08/2009
'Formato:
'   Reloj, fecha, hora, Legajo, E/S (siempre E)
'Ejemplo:
'   03 31/01/08 17:06 0000003077 E
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Long
Dim tipotarj As Long
Dim Legajo As Long
Dim Reg_Valida As Boolean

    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    nroreloj = Mid(strreg, pos1, pos2 - pos1)
    nrorelojtxt = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    regfecha = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Fecha:  " & regfecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Legajo:  " & NroLegajo
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    

'    pos1 = pos2
'    pos2 = InStr(pos1 + 1, strReg, " ")
'    'EntradaSalida = IIf(Trim(Mid(strReg, pos1)) = "20", "E", "S")
'    EntradaSalida = Mid(strReg, pos1, pos2 - pos1)
    
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = Trim(Mid(strreg, pos1))
    
    
'    Flog.writeline "Busco el nro de tarjeta "
'    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & TipoTarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(RegFecha) & ") AND ( (" & ConvFecha(RegFecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
'    OpenRecordset StrSql, objRs
'    If Not objRs.EOF Then
'       Ternro = objRs!Ternro
'    Else
'      'StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & nroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
'      'OpenRecordset StrSql, objRs
'      'If Not objRs.EOF Then
'      '   Ternro = objRs!Ternro
'      'Else
'         Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & TipoTarj & " , Reloj: " & codReloj
'         Flog.writeline "SQL: " & StrSql
'         InsertaError 1, 33
'         Exit Sub
'      'End If
'    End If
    
    Flog.writeline "Busco el nro de legajo"
    Legajo = CLng(Trim(NroLegajo))
    
    StrSql = "SELECT Ternro FROM empleado "
    StrSql = StrSql & " WHERE empleg = " & Legajo & " "
    OpenRecordset StrSql, objRs
    Ternro = 0
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el empleado con Legajo --> " & NroLegajo
        Exit Sub
    Else
        Ternro = objRs!Ternro
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
End Sub


Private Sub InsertaFormato17(strreg As String)
'-------------------------------------------------------------------
'FGZ - 10/08/2009
'Formato:
'   Legajo, fecha, hora, Reloj, Entrada/Salida
'Ejemplo:
'   27001 28/7/2008 17:16 L05 Entrada Sala de Computos
' Ultima Modificacion: FGZ - 360/04/2010 - Cuando se lee de un reloj sin la marca de control de acceso
'                           La registracion queda en estado X y para que los procesos no la tomen.
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim Horastr
Dim HH As String
Dim MM As String
Dim entradasalida As String
Dim nroreloj As String
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean

    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Legajo:  " & NroLegajo
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    regfecha = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Fecha:  " & regfecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    Horastr = Split(Hora, ":")
    HH = Format(Horastr(0), "00")
    MM = Format(Horastr(1), "00")
    Hora = HH & ":" & MM
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    nroreloj = Mid(strreg, pos1, pos2 - pos1)
    nrorelojtxt = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro,relvalestado FROM gti_reloj WHERE relcodext = '" & Trim(nroreloj) & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro,relvalestado FROM gti_reloj WHERE relcodext = '" & Trim(nrorelojtxt) & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
            Reg_Valida = CBool(objRs!relvalestado)
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
        Reg_Valida = CBool(objRs!relvalestado)
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    'EntradaSalida = IIf(Trim(Mid(strReg, pos1)) = "20", "E", "S")
    entradasalida = Mid(strreg, pos1, pos2 - pos1)
    If UCase(Trim(entradasalida)) = "ENTRADA" Then
        entradasalida = "E"
    Else
        entradasalida = "S"
    End If
    
    Flog.writeline "Busco el nro de tarjeta "
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      'StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & nroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      'OpenRecordset StrSql, objRs
      'If Not objRs.EOF Then
      '   Ternro = objRs!Ternro
      'Else
         Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " , Reloj: " & codReloj
         Flog.writeline "SQL: " & StrSql
         InsertaError 1, 33
         Exit Sub
      'End If
    End If
    
    
    'Carmen Quintero - 15/05/2015
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES ("
        StrSql = StrSql & Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj
        'StrSql = StrSql & ",'I')"
        If Reg_Valida Then
            StrSql = StrSql & ",'I'"
        Else
            StrSql = StrSql & ",'X'"
        End If
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
End Sub


Private Sub InsertaFormato18(strreg As String)
'-------------------------------------------------------------------
'FGZ - 02/03/2010 Formato para Supercanal
'Formato:
'   DNI, Reloj, Entrada/Salida, fecha, hora

'campo 1: DNI
'campo 2: Fijo "01" ---> Este codigo lo vamos a tomar como el codigo externo del reloj
'campo 3: Entrada = "00", Salida = "01"
'campo 4: Fecha
'campo 5: Hora

'Ejemplo:
'14169571,01,01,2010-02-11 00:06:39
'22536144,01,01,2010-02-11 00:06:47
'22536144,01,01,2010-02-11 00:06:49
'29326686,01,01,2010-02-11 00:07:29
'28847342,01,01,2010-02-11 00:08:06
'32352693,01,01,2010-02-11 00:08:11
'13540019,01,01,2010-02-11 00:11:46

'-------------------------------------------------------------------
Dim DNI As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Fecha_Hora As String
Dim Hora As String
Dim Horastr
Dim HH As String
Dim MM As String
Dim entradasalida As String
Dim nroreloj As String
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean

    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, ",")
    DNI = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "DNI:  " & DNI
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, ",")
    nroreloj = Mid(strreg, pos1, pos2 - pos1)
    nrorelojtxt = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
        
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, ",")
    entradasalida = IIf(Trim(Mid(strreg, pos1, pos2 - pos1)) = "00", "E", "S")
        
        
    pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, ",")
    pos2 = Len(strreg)
    Fecha_Hora = Mid(strreg, pos1)
    Flog.writeline "Fecha y hora (AAAA-MM-DD HH:MM:SS):  " & Fecha_Hora
            
        
    'pos1 = pos2
    'pos2 = InStr(pos1 + 1, strReg, ",")
    regfecha = Mid(Fecha_Hora, 9, 2) & "/" & Mid(Fecha_Hora, 6, 2) & "/" & Mid(Fecha_Hora, 1, 4)
    Flog.writeline "Fecha (DD/MM/AAAA):  " & regfecha
    
    Hora = Trim(Mid(Fecha_Hora, 12, 8))
    Horastr = Split(Hora, ":")
    HH = Format(Horastr(0), "00")
    MM = Format(Horastr(1), "00")
    Hora = HH & ":" & MM
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
       
    
    Flog.writeline "Busco el legajo con el DNI "
    'StrSql = "SELECT ternro FROM ter_doc WHERE nrodoc = '" & DNI & "' AND  tidnro <=5 "
    StrSql = "SELECT empleado.ternro, empleado.empleg, empleado.empest FROM ter_doc "
    StrSql = StrSql & " INNER JOIN empleado ON ter_doc.ternro = empleado.ternro "
    StrSql = StrSql & " WHERE nrodoc = '" & DNI & "' AND  tidnro <=5 "
    StrSql = StrSql & " AND empleado.empest <> 0"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        If objRs.RecordCount > 1 Then
            Flog.writeline "Error. Hay mas de un empleado ACTIVO con el mismo DNI " & DNI
            Do While Not objRs.EOF
                Flog.writeline "    Legajo " & objRs!EmpLeg
                objRs.MoveNext
            Loop
            InsertaError 1, 33
            Exit Sub
        Else
            Ternro = objRs!Ternro
        End If
    Else
         Flog.writeline "Error. El DNI " & DNI & "no se encuentra o no hay ningun empleado ACTIVO con ese DNI."
         Flog.writeline "SQL: " & StrSql
         InsertaError 1, 33
         Exit Sub
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & DNI & "  ;  '" & regfecha & "'    ;    " & Hora
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & DNI & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
End Sub



Private Sub InsertaFormato19(strreg As String)
'-------------------------------------------------------------------
'FGZ - 05/03/2010 - Formato para Servicentro la Estrella
'Formato:
    'campo 1: codigo de reloj
    'campo 2: Nombre del reloj
    'campo 3: El campo esta compuesto por fecha y hora.
    'campo 4: Event: Ignorarlo.
    'campo 5: User No.: Legajo del Usuario. Coincide con el legajo en RH Pro.
    'campo 6: User Name: Nombre del usuario.
    'campo 7: Vacio.


'Ejemplo:
'Se adjunta un ejemplo del archivo original, respetar los separadores <TAB>:

'Device No.  Device name Time    Event   User No.    User name   Card No.
'1   Estrella    23/02/2010 09:10:23 a.m.    Deny Access 15  ESCUDERO
'1   Estrella    23/02/2010 09:10:23 a.m.    Deny Access 15  ESCUDERO
'1   Estrella    23/02/2010 09:08:00 a.m.    Deny Access 168 FLORES A
'1   Estrella    23/02/2010 09:08:00 a.m.    Deny Access 168 FLORES A
'1   Estrella    23/02/2010 08:18:54 a.m.    Deny Access 36  ROSALES
'1   Estrella    23/02/2010 08:18:54 a.m.    Deny Access 36  ROSALES
'1   Estrella    23/02/2010 08:10:15 a.m.    Deny Access 214 GIMENEZ
'1   Estrella    23/02/2010 08:10:15 a.m.    Deny Access 214 GIMENEZ
'1   Estrella    23/02/2010 08:08:15 a.m.    Deny Access 292 ALBARRAC
'1   Estrella    23/02/2010 08:08:15 a.m.    Deny Access 292 ALBARRAC
'Count   100

'FGZ - 22/03/2010 - Formato para Servicentro la Estrella
'La primer linea es de encabezado y ultima linea es totalizadora de registraciones.
'    no deben leerse

'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim Fecha_Hora As String
Dim regfecha As String
Dim Hora As String
Dim Horastr
Dim HH As String
Dim MM As String
Dim AM_PM As String
Dim entradasalida As String
Dim nroreloj As String
Dim nrorelojtxt As String
Dim NombreReloj As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Aux As String
Dim par
Dim Reg_Valida As Boolean

    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1 + 1, strreg, Chr(9))
    nroreloj = Mid(strreg, pos1, pos2 - pos1)
    nrorelojtxt = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    'FGZ - 22/03/2010 - la ultima linea es totalizadora y no debe leerse
    If UCase(nroreloj) = "COUNT" Then
        Flog.writeline "Linea final."
        Exit Sub
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, Chr(9))
    NombreReloj = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, Chr(9))
    Fecha_Hora = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Fecha y hora (DD/MM/AAAA hh:mm:ss a.m.):  " & Fecha_Hora
        
    regfecha = Mid(Fecha_Hora, 1, 2) & "/" & Mid(Fecha_Hora, 4, 2) & "/" & Mid(Fecha_Hora, 7, 4)
    Flog.writeline "Fecha (DD/MM/AAAA):  " & regfecha
    
    Hora = Trim(Mid(Fecha_Hora, 12, 8))
    AM_PM = Trim(Mid(Fecha_Hora, 21, 4))
    
    Horastr = Split(Hora, ":")
    HH = Format(Horastr(0), "00")
    If UCase("p.m.") = UCase(AM_PM) Then
        If Horastr(0) = "12" Then
            HH = Format(Horastr(0), "00")
        Else
            HH = Format(Horastr(0) + 12, "00")
        End If
    End If
    If UCase("a.m.") = UCase(AM_PM) Then
        If Horastr(0) = "12" Then
            HH = Format(0, "00")
        End If
    End If
    
    MM = Format(Horastr(1), "00")
    Hora = HH & ":" & MM
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, Chr(9))
    Aux = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Fecha:  " & regfecha
    
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1, strreg, Chr(9))
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Legajo:  " & NroLegajo
    
   
    'Validaciones
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    'El reloj aparentemente no distingue entre entrada y salida
    entradasalida = "NULL"
    
    Flog.writeline "Busco el nro de tarjeta "
    StrSql = "SELECT ternro FROM empleado WHERE empleg = '" & NroLegajo & "'"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
        Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " , Reloj: " & codReloj
        Flog.writeline "SQL: " & StrSql
        InsertaError 1, 33
        Exit Sub
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    
    
        'StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
        'Ternro & "," & crpNro & "," & ConvFecha(RegFecha) & ",'" & Hora & "','" & EntradaSalida & "'," & codReloj & ",'I')"
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
End Sub

Private Sub InsertaFormato20()
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : EGO - Liz Oviedo
' Fecha      : 10/06/2010
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Ternro As Long
Dim codReloj As Integer

Dim rs_WC As New ADODB.Recordset
Dim rs_GTI_Registracion As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim HayError As Boolean
Dim tipotarj As Integer
Dim estadoInicial As String
Dim Reg_Valida As Boolean

codReloj = 1
StrSql = "SELECT * FROM WC_BAJADA_REG WHERE bpronro = 0"
OpenRecordset StrSql, rs_WC

Do While Not rs_WC.EOF
    
    Ternro = 0
    HayError = False

    'Valido hora
    Flog.writeline "Busco la hora"
    If Not objFechasHoras.ValidarHora(rs_WC!Reghora) Then
        Flog.writeline " Error Hora: " & rs_WC!Reghora
        HayError = True
    End If
    
    'Valido reloj
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro,relvalestado FROM gti_reloj WHERE relcodext = '" & rs_WC!relcodext & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
       Flog.writeline "Error. Reloj no encontrado: " & rs_WC!relcodext
       HayError = True
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
        'Ahora tengo en cuenta si el reloj asociado tiene la marca de control de acceso
        Reg_Valida = CBool(objRs!relvalestado)
    End If
    
    If Reg_Valida Then
        estadoInicial = "I"
    Else
        estadoInicial = "X"
    End If
    
    Flog.writeline "Busco el nro de tarjeta "
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & rs_WC!EmpLeg & "' AND (hstjfecdes <= " & ConvFecha(rs_WC!regfecha) & ") AND ( (" & ConvFecha(rs_WC!regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
       Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & rs_WC!EmpLeg & ", tipo de tarjeta: " & tipotarj & " , Reloj: " & codReloj
       HayError = True
    End If
    
    Flog.writeline "Busco la entrada/salida"
    If rs_WC!regentsal <> "E" And rs_WC!regentsal <> "S" Then
       Flog.writeline "Error. La entrada/salida debe ser E o S "
       HayError = True
    End If
    
    If HayError = False Then
        
        StrSql = "SELECT * FROM gti_registracion "
        StrSql = StrSql & " WHERE ternro =" & Ternro
        StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_WC!regfecha)
        StrSql = StrSql & " AND reghora ='" & rs_WC!Reghora & "'"
        OpenRecordset StrSql, rs_GTI_Registracion
        
        If rs_GTI_Registracion.EOF Then
            StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,regentsal,relnro,regestado, regmanual) VALUES ("
            StrSql = StrSql & Ternro & ","
            StrSql = StrSql & ConvFecha(rs_WC!regfecha) & ",'"
            StrSql = StrSql & Replace(rs_WC!Reghora, ":", "") & "','"
            StrSql = StrSql & rs_WC!regentsal & "',"
            StrSql = StrSql & codReloj
            StrSql = StrSql & ",'" & estadoInicial & "',"
            StrSql = StrSql & rs_WC!regmanual & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'si el reloj no tiene la marca de control de acceso ==> se inserta en un estado NN para que el proceso no la tenga en cuenta en el procesamiento
            If Reg_Valida Then
                Call InsertarWF_Lecturas(Ternro, rs_WC!regfecha)
            End If
        Else
            Flog.writeline " esa registracion ya existe. Legajo: " & rs_WC!EmpLeg & " Fecha: " & rs_WC!regfecha & " Hora: " & rs_WC!Reghora
        End If
        
        
        'Actualizo la tabla poniendo el proceso
        StrSql = "UPDATE WC_BAJADA_REG SET bpronro =" & NroProceso
        StrSql = StrSql & " WHERE empleg = " & rs_WC!EmpLeg
        StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_WC!regfecha)
        StrSql = StrSql & " AND reghora ='" & rs_WC!Reghora & "'"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    rs_WC.MoveNext
    
Loop

Fin:
    
    If rs_GTI_Registracion.State = adStateOpen Then rs_GTI_Registracion.Close
    If rs_WC.State = adStateOpen Then rs_WC.Close
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    
    Set rs_GTI_Registracion = Nothing
    Set rs_WC = Nothing
    Set rs_Empleado = Nothing
    
End Sub


Private Sub InsertaFormato22()
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : FGZ
' Fecha      : 18/03/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Ternro As Long
Dim codReloj As Integer

Dim rs_WC As New ADODB.Recordset
Dim rs_GTI_Registracion As New ADODB.Recordset
Dim rs_Empleado As New ADODB.Recordset
Dim HayError As Boolean
Dim tipotarj As Integer
Dim estadoInicial As String
Dim Reg_Valida As Boolean
Dim Cantidad As Long
Dim IncPorc As Single
Dim Progreso As Single


codReloj = 1
StrSql = "SELECT * FROM WC_BAJADA_REG WHERE bpronro = 0"
OpenRecordset StrSql, rs_WC


Progreso = 0
If rs_WC.RecordCount = 0 Then
    Cantidad = 1
Else
    Cantidad = rs_WC.RecordCount
End If
IncPorc = 100 / Cantidad

Do While Not rs_WC.EOF
    
    Ternro = 0
    HayError = False

    'Valido hora
    Flog.writeline "Busco la hora"
    If Not objFechasHoras.ValidarHora(rs_WC!Reghora) Then
        Flog.writeline " Error Hora: " & rs_WC!Reghora
        HayError = True
    End If
    
    'Valido reloj
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro,relvalestado FROM gti_reloj WHERE relcodext = '" & rs_WC!relcodext & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
       Flog.writeline "Error. Reloj no encontrado: " & rs_WC!relcodext
       HayError = True
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
        'Ahora tengo en cuenta si el reloj asociado tiene la marca de control de acceso
        Reg_Valida = CBool(objRs!relvalestado)
    End If
    
    If Reg_Valida Then
        estadoInicial = "I"
    Else
        estadoInicial = "X"
    End If
    
    Flog.writeline "Busco el nro de tarjeta "
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & rs_WC!EmpLeg & "' AND (hstjfecdes <= " & ConvFecha(rs_WC!regfecha) & ") AND ( (" & ConvFecha(rs_WC!regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
       Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & rs_WC!EmpLeg & ", tipo de tarjeta: " & tipotarj & " , Reloj: " & codReloj
       HayError = True
    
        'FGZ - 02/06/2011 ---------------------------------------------------------------------
        'no encuentre los empleados asociados a una tarjeta las marque y no las vuelva a tomar
            'Actualizo la tabla poniendo el proceso
            StrSql = "UPDATE WC_BAJADA_REG SET bpronro =" & NroProceso
            StrSql = StrSql & " ,regestado = 'X'"
            StrSql = StrSql & " WHERE regnro = '" & rs_WC!Regnro & "'"
            'FGZ - 06/04/2011 - se cambió la forma de actualizar ------
            objConn.Execute StrSql, , adExecuteNoRecords
        'FGZ - 02/06/2011 ---------------------------------------------------------------------
    End If
    
    Flog.writeline "Busco la entrada/salida"
    If rs_WC!regentsal <> "E" And rs_WC!regentsal <> "S" Then
       Flog.writeline "Error. La entrada/salida debe ser E o S "
       HayError = True
    End If
    
    If HayError = False Then
        StrSql = "SELECT * FROM gti_registracion "
        StrSql = StrSql & " WHERE ternro =" & Ternro
        StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_WC!regfecha)
        StrSql = StrSql & " AND reghora ='" & rs_WC!Reghora & "'"
        OpenRecordset StrSql, rs_GTI_Registracion
        
        If rs_GTI_Registracion.EOF Then
            StrSql = " INSERT INTO gti_registracion(ternro,regfecha,reghora,regentsal,relnro,regestado, regmanual) VALUES ("
            StrSql = StrSql & Ternro & ","
            StrSql = StrSql & ConvFecha(rs_WC!regfecha) & ",'"
            StrSql = StrSql & Replace(rs_WC!Reghora, ":", "") & "','"
            StrSql = StrSql & rs_WC!regentsal & "',"
            StrSql = StrSql & codReloj
            StrSql = StrSql & ",'" & estadoInicial & "',"
            StrSql = StrSql & rs_WC!regmanual & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'si el reloj no tiene la marca de control de acceso ==> se inserta en un estado NN para que el proceso no la tenga en cuenta en el procesamiento
            If Reg_Valida Then
                Call InsertarWF_Lecturas(Ternro, rs_WC!regfecha)
            End If
        Else
            Flog.writeline " esa registracion ya existe. Legajo: " & rs_WC!EmpLeg & " Fecha: " & rs_WC!regfecha & " Hora: " & rs_WC!Reghora
        End If
        
        
        'Actualizo la tabla poniendo el proceso
        StrSql = "UPDATE WC_BAJADA_REG SET bpronro =" & NroProceso
        'FGZ - 06/04/2011 - se cambió la forma de actualizar ------
        'StrSql = StrSql & " WHERE empleg = '" & rs_WC!EmpLeg & "'"
        'StrSql = StrSql & " AND regfecha =" & ConvFecha(rs_WC!RegFecha)
        'StrSql = StrSql & " AND reghora ='" & rs_WC!Reghora & "'"
        StrSql = StrSql & " WHERE regnro = '" & rs_WC!Regnro & "'"
        'FGZ - 06/04/2011 - se cambió la forma de actualizar ------
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    rs_WC.MoveNext
Loop

Fin:
    If rs_GTI_Registracion.State = adStateOpen Then rs_GTI_Registracion.Close
    If rs_WC.State = adStateOpen Then rs_WC.Close
    If rs_Empleado.State = adStateOpen Then rs_Empleado.Close
    
    Set rs_GTI_Registracion = Nothing
    Set rs_WC = Nothing
    Set rs_Empleado = Nothing
End Sub



Private Sub InsertaFormatoMonresa(ByVal strreg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato de Lectura Condicional (2 formatos posibles)
' Autor      : FGZ
' Fecha      : 26/07/2010
' Ultima Mod.:
' Descripcion:
'Formato 1(marcas HP.txt):
'   tipo de marca
'   número de empleado
'   Fecha
'   Hora
'   número de reloj
'
'Ejemplo
'    123456789012345678901234567890123456789
'----------------------------------------
'    01 0000008490 21-04-10  17:17:44  3
'    01 0000005223 21-04-10  17:19:30  2
'    01 0000009935 21-04-10  17:20:42  3
'    01 0000009930 21-04-10  17:21:02  2
'    01 0000009770 21-04-10  17:21:46  3
'----------
'Format 2(marcasFC.txt):
'   carácter "M"
'   número de empleado
'   Fecha
'   Hora
'   tipo de marca
'   número de reloj
'
'-------------------------
            'Tipo de Marca
                '01 es "entrada al turno"
                '02 es "salida a pausa de descanso",
                '03 es "entrada de fin de pausa de descanso"
                '04 es la "salida del turno"
'Ejemplo
'----------------------------------------
'   12345678901234567890123456789012
'   M     9420 21/04/2010 17:16 2 03
' ---------------------------------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
'Dim NroTarj As Long
Dim NroTarj As String
Dim NroTarj_aux As String
Dim descarte  As String
Dim Validar As Boolean
Dim tipo_reloj As Integer
Dim Origen As String
Dim comedor As Boolean
Dim Legajo As String
Dim Tipo_Marca As String
Dim campos
Dim Auxiliar As String
Dim Reg_Valida As Boolean


    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    Validar = False
    comedor = False
    
    Flog.writeline "   - Registración --> " & strreg
    
    separador = Chr(32)
    
    campos = Split(strreg, separador)
    Auxiliar = Mid(strreg, 1, 1)
    
    Select Case Auxiliar
        Case "M": 'Format 2(marcasFC.txt):
            'carácter "M"
            'número de empleado
            'Fecha
            'Hora
            'tipo de marca
            'número de reloj
            
            'Empleado
            NroTarj = Trim(Mid(strreg, 3, 8)) 'campos(1)
            
            'Fecha
            fecha_aux = Trim(Mid(strreg, 12, 10)) 'campos(2)
            Dia = Mid(fecha_aux, 1, 2)
            Mes = Mid(fecha_aux, 4, 2)
            Anio = Mid(fecha_aux, 7, 4)
            
            'Hora
            Hora = Mid(strreg, 23, 5) 'campos(3)
            
            'Tipo de marca
            Tipo_Marca = Mid(strreg, 29, 1) 'campos(4)
            Select Case Tipo_Marca
            Case "1", "01", "3", "03":
                entradasalida = "E"
            Case "2", "02", "4", "04"
                entradasalida = "S"
            Case Else
                entradasalida = ""
            End Select
            
            'Nro reloj
            nroreloj = Mid(strreg, 31, 2) 'campos(5)
            
            Validar = True
        Case Else: 'Formato 1(marcas HP.txt):
            'tipo de marca
            'número de empleado
            'Fecha
            'Hora
            'número de reloj

            '------------------------
            'Tipo de Marca
            Tipo_Marca = Trim(Mid(strreg, 1, 2)) 'campos(0)
            Select Case Tipo_Marca
            Case "1", "01", "3", "03":
                entradasalida = "E"
            Case "2", "02", "4", "04"
                entradasalida = "S"
            Case Else
                entradasalida = ""
            End Select
            
            'Empleado
            NroTarj = Trim(Mid(strreg, 4, 10)) 'campos(1)
            
            'Fecha
            fecha_aux = Trim(Mid(strreg, 15, 8)) 'campos(2)
            Dia = Mid(fecha_aux, 1, 2)
            Mes = Mid(fecha_aux, 4, 2)
            Anio = Mid(fecha_aux, 7, 2)
            Anio = "20" & Anio
            
            'Hora
            Hora = Trim(Mid(strreg, 25, 5)) 'campos(3)
            
            'Nro reloj
            'NroReloj = Trim(Mid(strReg, 35, 1)) 'campos(4)
            nroreloj = Trim(Mid(strreg, 35, 2)) 'campos(4)
                
            Validar = True
    End Select
    

'====================================================================
' Validar los parametros Levantados
    
    If Validar Then
        Flog.writeline "Validaciones..."
        
        Flog.writeline "Busco el reloj"
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            GoTo Fin
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
        
        'Que la fecha sea válida
        If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
            Flog.writeline " Error Fecha: " & fecha_aux
            InsertaError 4, 4
            GoTo Fin
        Else
            regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
        End If
        
        'Que la hora sea válida
        If Not objFechasHoras.ValidarHora(Hora) Then
            Flog.writeline " Error Hora: " & Hora
            InsertaError 4, 37
            GoTo Fin
        End If
        
    
        'Busco que el nro de tarjeta sea válido
        StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroTarj & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Ternro = objRs!Ternro
        Else
            Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " , Reloj: " & codReloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 1, 33
            GoTo Fin
        End If
        
        'Carmen Quintero - 15/05/2015
        Reg_Valida = True
        StrSql = "SELECT relnro FROM gti_rel_estr "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
            'Valido que el reloj sea de control de acceso para el empleado
            StrSql = "SELECT ternro FROM his_estructura H "
            StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
            StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
            StrSql = StrSql & " AND ( h.ternro = " & Ternro
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Reg_Valida = False
                Flog.writeline "    El reloj No está habilitado para el empleado "
            End If
        End If
        'Fin Carmen Quintero - 15/05/2015
         
        StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
            If Reg_Valida Then
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I','" & Tipo_Marca & "')"
            Else
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X','" & Tipo_Marca & "')"
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "SQL: -->" & StrSql
            
            Call InsertarWF_Lecturas(Ternro, regfecha)
            Flog.writeline "Inserto en temporal WF_Lecturas"
        Else
            Flog.writeline " Registracion ya Existente"
            Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
            Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
            InsertaError 1, 92
        End If
        Flog.writeline "Linea Procesada"
      End If

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub



Private Sub InsertaFormato23(ByVal strreg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato de Lectura para Cliente Mundo Maipu
' Autor      : FGZ
' Fecha      : 26/07/2010
' Ultima Mod.: 17/09/2012 - Manterola Maria Magdalena (CAS-15535 - Mundo Maipu - Inhabilitacion de Relojes)- Se considera la inhabilitación de los relojes. Es decir, si un reloj esta inhabilitado no se pueden cargar registraciones asociadas al mismo.
' Descripcion:
'Formato: los campos estan separados por 1 tab
'   Legajo
'   Fecha
'   Hora
'   número de reloj
'   Tipo de Marca
'
'Ejemplo
'    00001 28/03/2011 18:17 00051 00
'    00001 28/03/2011 18:24 00051 00
'    03001 29/03/2011 10:16 00051 00
'    03001 29/03/2011 10:16 00051 00
'-------------------------
            'Tipo de Marca
                '00 es "entrada" --> E
                '01 es "salida" --> S
                '02 es "Salida Almuerzo" --> S
                '03 es "Regreso de Almuerzo" --> E
' ---------------------------------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Long
Dim tipotarj As Integer
Dim NroTarj As String
Dim NroTarj_aux As String
Dim descarte  As String
Dim Validar As Boolean
Dim tipo_reloj As Integer
Dim Origen As String
Dim comedor As Boolean
Dim Legajo As String
Dim Tipo_Marca As String
Dim TipoReg As Long
Dim campos
Dim Reg_Valida As Boolean
'Dim Auxiliar As String


    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    Validar = False
    comedor = False
    
    Flog.writeline "   - Registración --> " & strreg
    
    'Separador = Chr(9)  'TAB
    separador = Chr(32)  'Espacio
    campos = Split(strreg, separador)
    
    'Formato
    '------------------------
    'número de legajo
    'Fecha
    'Hora
    'número de reloj
    'tipo de marca
    '------------------------
    'Empleado
    NroTarj = Trim(campos(0))
    'Fecha
    fecha_aux = Trim(campos(1))
    'Hora
    Hora = Trim(campos(2))
    'Nro reloj
    nroreloj = Trim(campos(3))
    'Tipo de Marca
    Tipo_Marca = Trim(campos(4))
    
    
    StrSql = "SELECT * FROM gti_tiporeg WHERE tiporegcod = '" & Tipo_Marca & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Tipo de marca desconocido " & Tipo_Marca & ". Registracion descartada."
        Validar = False
        TipoReg = 0
    Else
        entradasalida = UCase(objRs!tiporeges)
        Validar = True
        TipoReg = objRs!tiporegnro
    End If
    'Select Case Tipo_Marca
    'Case "00":
    '    EntradaSalida = "E"
    '    Tipo_Marca = "01"
    '    Flog.writeline "Tipo de marca desconocido " & Tipo_Marca & ". Se usará valor Dafault."
    'Case "0", "00", "4", "04":
    '    EntradaSalida = "E"
    'Case "1", "01", "3", "03"
    '    EntradaSalida = "S"
    'Case Else
    '    EntradaSalida = ""
    '    Flog.writeline "Tipo de marca desconocido " & Tipo_Marca
    'End Select

'====================================================================
' Validar los parametros Levantados
    
    If Validar Then
        Flog.writeline "Validaciones..."
        
        Flog.writeline "Busco el reloj"
        StrSql = "SELECT relnro, tptrnro,relhabil FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            GoTo Fin
        Else
        
            'MMM - 17/09/2012 *******************************************
            Flog.writeline "Busco si el reloj esta correctamente habilitado. Si no lo esta, se produce ERROR."
            If objRs!relhabil = 0 Then
                Flog.writeline "Error. Reloj Inhabilitado: " & nroreloj
                Flog.writeline "SQL: " & StrSql
                InsertaError 4, 32
                GoTo Fin
            Else
            'MMM - 17/09/2012 *******************************************
            
                codReloj = objRs!relnro
                tipotarj = objRs!tptrnro
            End If
        End If
        
        ''Que la fecha sea válida
        'If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
        '    Flog.writeline " Error Fecha: " & fecha_aux
        '    InsertaError 4, 4
        '    GoTo Fin
        'Else
        '    RegFecha = CDate(Dia & "/" & Mes & "/" & Anio)
        'End If
        regfecha = CDate(fecha_aux)
        
        'Que la hora sea válida
        If Not objFechasHoras.ValidarHora(Hora) Then
            Flog.writeline " Error Hora: " & Hora
            InsertaError 4, 37
            GoTo Fin
        End If
        
    
        ''Busco que el nro de tarjeta sea válido
        'StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & TipoTarj & " AND hstjnrotar = '" & NroTarj & "' AND (hstjfecdes <= " & ConvFecha(RegFecha) & ") AND ( (" & ConvFecha(RegFecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        'OpenRecordset StrSql, objRs
        'If Not objRs.EOF Then
        '    Ternro = objRs!Ternro
        'Else
        '    Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & TipoTarj & " , Reloj: " & codReloj
        '    Flog.writeline "SQL: " & StrSql
        '    InsertaError 1, 33
        '    GoTo Fin
        'End If
         
        'Valido el nro de legajo
        StrSql = "SELECT ternro FROM empleado WHERE empleg = " & NroTarj
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Ternro = objRs!Ternro
        Else
            Flog.writeline "Error. No se encuentra el Legajo: " & NroTarj
            Flog.writeline "SQL: " & StrSql
            InsertaError 1, 33
            GoTo Fin
        End If
        
        'Carmen Quintero - 15/05/2015
        Reg_Valida = True
        StrSql = "SELECT relnro FROM gti_rel_estr "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
            'Valido que el reloj sea de control de acceso para el empleado
            StrSql = "SELECT ternro FROM his_estructura H "
            StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
            StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
            StrSql = StrSql & " AND ( h.ternro = " & Ternro
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Reg_Valida = False
                Flog.writeline "    El reloj No está habilitado para el empleado "
            End If
        End If
        'Fin Carmen Quintero - 15/05/2015
         
        StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
            If Reg_Valida Then
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I','" & Tipo_Marca & "')"
            Else
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X','" & Tipo_Marca & "')"
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "SQL: -->" & StrSql
            
            Call InsertarWF_Lecturas(Ternro, regfecha)
            Flog.writeline "Inserto en temporal WF_Lecturas"
        Else
            Flog.writeline " Registracion ya Existente"
            Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
            Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
            InsertaError 1, 92
        End If
        Flog.writeline "Linea Procesada"
      End If

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub

Private Sub InsertaFormato24(ByVal strreg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato de Lectura para Cliente CAJA DE ODONTOLOGOS
' Autor      : Gonzalez Nicolás
' Fecha      : 30/09/2011
' Ultima Mod.:
' Descripcion:
'Formato: los campos estan separados por 1 ,
'Fecha,hora,legajo,nro reloj
'   Fecha
'   Hora
'   Legajo
'   número de reloj

'
'Ejemplo
'    2011/07/15,16:00,220152,1175
' ---------------------------------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Long
Dim tipotarj As Integer
Dim NroTarj As String
Dim NroTarj_aux As String
Dim descarte  As String
Dim Validar As Boolean
Dim tipo_reloj As Integer
Dim Origen As String
'Dim comedor As Boolean
Dim Legajo As String
Dim Tipo_Marca As String
Dim TipoReg As Long
Dim campos
Dim Reg_Valida As Boolean
'Dim Auxiliar As String


    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    Validar = True
'    comedor = False
    
    Flog.writeline "   - Registración --> " & strreg
    
    
    'BUSCO SEPARADOR | SI NO ESTA CONFIGURADO , POR DEFAULT
    StrSql = "SELECT modseparador FROM modelo WHERE modnro = 178"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        separador = objRs!modseparador
    Else
        separador = ","
    End If
    objRs.Close
    
    'Separador = Chr(9)  'TAB
    'Separador = Chr(32)  'Espacio
    campos = Split(strreg, separador)
    
    'Formato
    '------------------------
    '   Fecha
    '   Hora
    '   Legajo
    '   número de reloj
    '------------------------
    
    '---Fecha
    fecha_aux = Trim(campos(0))
    '---Hora
    Hora = Trim(campos(1))
    '---Empleado
    NroTarj = Trim(campos(2))
    '---Nro reloj
    nroreloj = Trim(campos(3))
    
    TipoReg = 0
            
'    StrSql = "SELECT * FROM gti_tiporeg WHERE tiporegcod = '" & Tipo_Marca & "'"
'    OpenRecordset StrSql, objRs
'    If objRs.EOF Then
'        Flog.writeline "Tipo de marca desconocido " & Tipo_Marca & ". Registracion descartada."
'        Validar = False
'        TipoReg = 0
'    Else
'        EntradaSalida = UCase(objRs!tiporeges)
'        Validar = True
'        TipoReg = objRs!tiporegnro
'    End If
    'Select Case Tipo_Marca
    'Case "00":
    '    EntradaSalida = "E"
    '    Tipo_Marca = "01"
    '    Flog.writeline "Tipo de marca desconocido " & Tipo_Marca & ". Se usará valor Dafault."
    'Case "0", "00", "4", "04":
    '    EntradaSalida = "E"
    'Case "1", "01", "3", "03"
    '    EntradaSalida = "S"
    'Case Else
    '    EntradaSalida = ""
    '    Flog.writeline "Tipo de marca desconocido " & Tipo_Marca
    'End Select

'====================================================================
' Validar los parametros Levantados
    
    If Validar Then
        Flog.writeline "Validaciones..."
        
        Flog.writeline "Busco el reloj"
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            GoTo Fin
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
        
        ''Que la fecha sea válida
        'If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
        '    Flog.writeline " Error Fecha: " & fecha_aux
        '    InsertaError 4, 4
        '    GoTo Fin
        'Else
        '    RegFecha = CDate(Dia & "/" & Mes & "/" & Anio)
        'End If
        regfecha = CDate(fecha_aux)
        
        'Que la hora sea válida
        If Not objFechasHoras.ValidarHora(Hora) Then
            Flog.writeline " Error Hora: " & Hora
            InsertaError 4, 37
            GoTo Fin
        End If
        
    
        ''Busco que el nro de tarjeta sea válido
        'StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & TipoTarj & " AND hstjnrotar = '" & NroTarj & "' AND (hstjfecdes <= " & ConvFecha(RegFecha) & ") AND ( (" & ConvFecha(RegFecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        'OpenRecordset StrSql, objRs
        'If Not objRs.EOF Then
        '    Ternro = objRs!Ternro
        'Else
        '    Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & TipoTarj & " , Reloj: " & codReloj
        '    Flog.writeline "SQL: " & StrSql
        '    InsertaError 1, 33
        '    GoTo Fin
        'End If
         
        'Valido el nro de legajo
        StrSql = "SELECT ternro FROM empleado WHERE empleg = " & NroTarj
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Ternro = objRs!Ternro
        Else
            Flog.writeline "Error. No se encuentra el Legajo: " & NroTarj
            Flog.writeline "SQL: " & StrSql
            InsertaError 1, 33
            GoTo Fin
        End If
        
         
        'Carmen Quintero - 15/05/2015
        Reg_Valida = True
        StrSql = "SELECT relnro FROM gti_rel_estr "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
            'Valido que el reloj sea de control de acceso para el empleado
            StrSql = "SELECT ternro FROM his_estructura H "
            StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
            StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
            StrSql = StrSql & " AND ( h.ternro = " & Ternro
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Reg_Valida = False
                Flog.writeline "    El reloj No está habilitado para el empleado "
            End If
        End If
        'Fin Carmen Quintero - 15/05/2015
 
        StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
        
            If Reg_Valida Then
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I','" & Tipo_Marca & "')"
            Else
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X','" & Tipo_Marca & "')"
            End If
            
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "SQL: -->" & StrSql
            
            Call InsertarWF_Lecturas(Ternro, regfecha)
            Flog.writeline "Inserto en temporal WF_Lecturas"
        Else
            Flog.writeline " Registracion ya Existente"
            Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
            Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
            InsertaError 1, 92
        End If
        Flog.writeline "Linea Procesada"
      End If

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub
Private Sub InsertaFormato25(ByVal strreg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato de Lectura para Cliente  Laboratorios SL- QA
' Autor      : Gonzalez Nicolás
' Fecha      : 17/10/2011
' Ultima Mod.:
' Descripcion:
'Formato: los campos estan separados por 1 Espacio
'I (9 caracteres): Id de persona (justificado a izquierda con espacios)
'aaaa-mm-dd hh:mm:ss (19 caracteres): Fecha y hora
'E: Tipo de evento (0: Entrada | 1:Salida | 255: No definido), de 1 a 3 caracteres.
'X: No se utilizan


'------------------
'
'Ejemplo
'321 2011-07-11 18:15:25 X 0 X X
' ---------------------------------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Long
Dim tipotarj As Integer
Dim NroTarj As String
Dim NroTarj_aux As String
Dim descarte  As String
Dim Validar As Boolean
Dim tipo_reloj As Integer
Dim Origen As String
'Dim comedor As Boolean
Dim Legajo As String
Dim Tipo_Marca As String
Dim TipoReg As Long
Dim campos
Dim Reg_Valida As Boolean
'Dim Auxiliar As String


    On Error GoTo MError
    
    RegLeidos = RegLeidos + 1
    Validar = True
'    comedor = False
    
    Flog.writeline "   - Registración --> " & strreg
    'Separador = Chr(9)  'TAB
    separador = Chr(32)  'Espacio
    campos = Split(strreg, separador)
    
    'Formato
    '------------------------
    '   Legajo
    '   Fecha
    '   Hora
    '   Tipo de evento (0: Entrada | 1:Salida | 255: No definido)
    '------------------------
    'strCmdLine = 7927
    'Err.Number = 1
    
    '---Empleado
    NroTarj = Trim(campos(0))
    
    '---Fecha
    fecha_aux = Trim(campos(1))
    '---Hora
    Hora = Trim(campos(2))
    
    'Tipo de Evento - (0: Entrada | 1:Salida | 255: No definido)
    entradasalida = Trim(campos(4))
    If entradasalida = "0" Then
        entradasalida = "E"
    ElseIf entradasalida = "1" Then
        entradasalida = "S"
    ElseIf entradasalida = "255" Then
        entradasalida = ""
    Else
        entradasalida = ""
    End If
    
   
    TipoReg = 0
            


'====================================================================
' Validar los parametros Levantados
    
    If Validar Then
        Flog.writeline "Validaciones..."
        
        Flog.writeline "Busco el reloj"
        'StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & NroReloj & "'"
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj "
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            GoTo Fin
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
        
        regfecha = CDate(fecha_aux)
        
        'Que la hora sea válida
        If Not objFechasHoras.ValidarHoraLarga(Hora) Then
            Flog.writeline " Error Hora: " & Hora
            InsertaError 4, 37
            GoTo Fin
        End If
        
        Hora = Format(Hour(Hora), "00") & Format(Minute(Hora), "00")
        
        
        'Valido el nro de legajo
        StrSql = "SELECT ternro FROM empleado WHERE empleg = " & CInt(NroTarj)
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Ternro = objRs!Ternro
        Else
            Flog.writeline "Error. No se encuentra el Legajo: " & NroTarj
            Flog.writeline "SQL: " & StrSql
            InsertaError 1, 33
            GoTo Fin
        End If
        
        'Carmen Quintero - 15/05/2015
        Reg_Valida = True
        StrSql = "SELECT relnro FROM gti_rel_estr "
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
            'Valido que el reloj sea de control de acceso para el empleado
            StrSql = "SELECT ternro FROM his_estructura H "
            StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
            StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
            StrSql = StrSql & " AND ( h.ternro = " & Ternro
            StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
            OpenRecordset StrSql, objRs
            If objRs.EOF Then
                Reg_Valida = False
                Flog.writeline "    El reloj No está habilitado para el empleado "
            End If
        End If
        'Fin Carmen Quintero - 15/05/2015
         
        StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
        
            If Reg_Valida Then
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I','" & Tipo_Marca & "')"
            Else
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X','" & Tipo_Marca & "')"
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "SQL: -->" & StrSql
            
            Call InsertarWF_Lecturas(Ternro, regfecha)
            Flog.writeline "Inserto en temporal WF_Lecturas"
        Else
            Flog.writeline " Registracion ya Existente"
            Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
            Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
            InsertaError 1, 92
        End If
        Flog.writeline "Linea Procesada"
      End If

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub


















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
Case 22: 'Lectura de Registraciones
    If Version >= "1.39" Then
        'Revisar los campos
        'gti_registracion.tiporeg
        Texto = "Revisar los campos: gti_registracion.tiporeg"
        
        StrSql = "Select tiporeg from gti_registracion WHERE regnro = 1"
        OpenRecordset StrSql, rs
        
        V = True
    End If

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




Private Sub InsertaFormato26(strreg As String)
'-------------------------------------------------------------------
'Autor: Gonzalez Nicolás - 06/10/2011
'Formato:
'   Legajo, fecha, hora, nro de reloj, E/S
'Ejemplo:
'   000001,22/11/2005,1000,0010,E
'   000001,22/11/2005,1800,0010,S
'Ult. Modif: 15/11/2011 - Gonzalez Nicolás - Se cambio la forma de guardar el registro EntradaSalida
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean

    RegLeidos = RegLeidos + 1
    pos1 = 1
    pos2 = InStr(pos1, strreg, ",")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Legajo:  " & NroLegajo
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, ",")
    regfecha = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Fecha:  " & regfecha
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, ",")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, ",")
    nroreloj = Mid(strreg, pos1, pos2 - pos1)
    nrorelojtxt = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. No se encontro el Reloj: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    

    pos1 = pos2 + 1
   ' pos2 = InStr(pos1 + 1, strReg, ",")
   ' EntradaSalida = IIf(UCase(Trim(Mid(strReg, pos1))) = "E", "E", "S")
    entradasalida = Mid(strreg, pos1, Len(strreg))
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    StrSql = StrSql
    If Not objRs.EOF Then
        Ternro = objRs!Ternro
    Else
        Flog.writeline "Error. No se encontro la terjeta para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj: " & codReloj
        InsertaError 1, 33
        Exit Sub
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
        
End Sub
Private Sub InsertaFormato27(strreg As String)

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean

    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Legajo:  " & NroLegajo
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    regfecha = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Fecha:  " & regfecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    nroreloj = Trim(Mid(strreg, pos1, pos2 - pos1))
    nrorelojtxt = Trim(Mid(strreg, pos1, pos2 - pos1))
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    

    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = Trim(Mid(strreg, pos1))
    
    Select Case entradasalida
        Case "20":
            entradasalida = "E"
        Case "21":
            entradasalida = "S"
        Case "99":
            entradasalida = "R"
    End Select
       
    'Que exista el legajo
    Flog.writeline "Busco el nro de legajo "
    StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & NroLegajo
        InsertaError 6, 8
        Exit Sub
    Else
        Ternro = objRs!Ternro
    End If
       
    'Flog.writeline "Busco el nro de tarjeta "
    'StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & TipoTarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(RegFecha) & ") AND ( (" & ConvFecha(RegFecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    'OpenRecordset StrSql, objRs
    'If Not objRs.EOF Then
    '   Ternro = objRs!Ternro
    'Else
      'StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & nroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      'OpenRecordset StrSql, objRs
      'If Not objRs.EOF Then
      '   Ternro = objRs!Ternro
      'Else
     '    Flog.writeline "Error. Trajeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & TipoTarj & " , Reloj: " & codReloj
     '    Flog.writeline "SQL: " & StrSql
     '    InsertaError 1, 33
     '    Exit Sub
      'End If
    'End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & IIf(entradasalida <> "R", entradasalida, Null) & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & IIf(entradasalida <> "R", entradasalida, Null) & "'," & codReloj & ",'X')"
        End If
        
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
End Sub
Private Sub InsertaFormato28(strreg As String)

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean

'FORMATO
'N°tarjeta DD/MM/AAAA HH:MM RELOJ E/S
'11410 02/07/2012 11:57 01 0E
    RegLeidos = RegLeidos + 1
    
    
    pos1 = 2 'Toma a partir de la posición 2, el primer registro corresponde al dedo que utilizo para fichar
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Legajo:  " & NroLegajo
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    regfecha = Mid(strreg, pos1, pos2 - pos1)
    Flog.writeline "Fecha:  " & regfecha
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
     pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    'NroReloj = Trim(Mid(strReg, pos1, pos2 - pos1))
    nrorelojtxt = Trim(Mid(strreg, pos1, pos2 - pos1))
    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    

    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = Trim(Mid(strreg, pos1))
    
'    If UCase(EntradaSalida) = "0E" Then
'        EntradaSalida = "E"
'    ElseIf UCase(EntradaSalida) = "0S" Then
'        EntradaSalida = "S"
'    Else
'        EntradaSalida = ""
'    End If
    
    entradasalida = "E"
    
    
      
    'Que exista el legajo
    Flog.writeline "Busco el nro de legajo "
    StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & NroLegajo
        InsertaError 6, 8
        Exit Sub
    Else
        Ternro = objRs!Ternro
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & IIf(entradasalida <> "R", entradasalida, Null) & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & IIf(entradasalida <> "R", entradasalida, Null) & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
End Sub


Private Sub InsertaFormato29(strreg As String)
'-------------------------------------------------------------------
'FGZ - 02/03/2010 Formato para SPEC
'Formato: Legajo FechaHora Lector Terminal
'-----------
'Formato de longitud fija de 60 caracteres
'Separador de campos: espacio
'Detalle
'-----------
'campo 1: Nro de Legajo
'campo 2: Fecha y Hora de la registracion formato dd/mm/yyyy hh:mm a.m./p.m.
'campo 3: Entrada/salida, define si el registro es de Entrada = E, Salida = S o No distingue Entrada de Salida = SN.
'campo 4: reloj


'Ejemplo:
'EmpleadoFechaLectorTerminal
'12344567  28/08/2012 02:53 p.m. SN      SNT2
'12344567  28/08/2012 02:53 p.m. SN      SNT2
'12344567  28/08/2012 02:53 p.m. SN      SNT2
'12344567  28/08/2012 02:25 p.m. S   Oficinas
'12344567  28/08/2012 02:25 p.m. E   Oficinas
'12344567  28/08/2012 02:25 p.m. S   Oficinas
'12344567  28/08/2012 02:25 p.m. E   Oficinas
'12344567  28/08/2012 02:25 p.m. S   Oficinas
'12344567  28/08/2012 02:25 p.m. E   Oficinas
'12344567  28/08/2012 02:25 p.m. S   Oficinas
'12344567  28/08/2012 02:25 p.m. E   Oficinas
'12344567  28/08/2012 01:37 p.m. S   Oficinas
'12344567  28/08/2012 01:37 p.m. S   Oficinas
'-------------------------------------------------------------------
Const FormatoInternoFechaCorto = "dd/mm/yyyy"
Const FormatoInternoHora = "HH:mm"

Dim Legajo As String
Dim Ternro As Long
Dim regfecha As String
Dim Fecha_Hora As String
Dim Hora As String

Dim entradasalida As String
Dim reloj As String

Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Sep As String
Dim Reg_Valida As Boolean

If strreg = "EmpleadoFechaLectorTerminal" Then
    'salteo la linea porque es de encabezado
    
Else
    RegLeidos = RegLeidos + 1
    Sep = " "
    
    pos1 = 1
    pos2 = 10
    Legajo = Trim(Mid(strreg, pos1, pos2))
    Flog.writeline "Legajo:  " & Legajo
    
    pos1 = 11
    pos2 = 22
    Fecha_Hora = Trim(Mid(strreg, pos1, pos2))
    Flog.writeline "Fecha y hora (DD/MM/YYYY HH:MM A.M./P.M.):  " & Fecha_Hora
        
    pos1 = 33
    pos2 = 3
    entradasalida = UCase(Trim(Mid(strreg, pos1, pos2)))
    Flog.writeline "Entrada/Salida (Entrada = E, Salida = S, SN = No distingue E de S):  " & entradasalida
    If entradasalida = "SN" Then
        entradasalida = " "
        'Queda por determinar si SN se descarta o simplemente No se marca
    End If
        
    pos1 = 36
    pos2 = 25
    reloj = Trim(Mid(strreg, pos1, pos2))
    Flog.writeline "Nro Reloj:  " & reloj
        
    '----------------------------------------------------------------------
    regfecha = Format(Fecha_Hora, FormatoInternoFechaCorto)
    Hora = Format(Fecha_Hora, FormatoInternoHora)
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & reloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE reldabr = '" & reloj & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Reloj no encontrado: " & reloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
            Flog.writeline "Reloj Encontrado"
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
        Flog.writeline "Reloj Encontrado"
    End If
    
    
    Flog.writeline "Busco el legajo"
    StrSql = "SELECT empleado.ternro FROM empleado "
    StrSql = StrSql & " WHERE empleg = '" & Legajo & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
         Flog.writeline "Error. El legajo no se encuentra."
         Flog.writeline "SQL: " & StrSql
         InsertaError 1, 33
         Exit Sub
    Else
        Ternro = objRs!Ternro
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT regnro FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & Legajo & "  ;  '" & regfecha & "'    ;    " & Hora
        
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & Legajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
End If

End Sub


Private Sub InsertaFormato30(strreg As String)
'-------------------------------------------------------------------
'Autor: FGZ - 10/01/2013
'Formato:
'    -Los Campos son de longitud fija y están separados por un espacio:
'    N° Tarjeta -5 dígitos
'    Fecha 10 dígitos
'    Hora 5 dígitos
'    Reloj 2 dígitos
'    Tarjeta Habilitada 2 dígitos
'    Apellido Nombre 30 dígitos
'
'OBS
'   Tarjeta Habilitada: los valores son 01 Habilitada y 03 No Habilitada
'---------------
'Ejemplo:
'   N° Tarjeta Fecha Hora Reloj Tarjeta Habilitada Apellido Nombre
'   51925 16/09/2012 04:28 33 01 MARTORELL SEBASTIAN
'
'Ult. Modif:
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Habilitada As String
Dim APYN As String
Dim regestado
Dim separador As String
Dim Reg_Valida As Boolean

separador = " "

If Left(strreg, 2) = "N°" Then
    'salteo la linea porque es de encabezado
    
Else
    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    'pos2 = InStr(pos1, strReg, Separador)
    'NroLegajo = Mid(strReg, pos1, pos2 - pos1)
    pos2 = 5
    NroLegajo = Mid(strreg, pos1, pos2)
    Flog.writeline "N° Tarjeta:  " & NroLegajo

    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, Separador)
    'RegFecha = Mid(strReg, pos1, pos2 - pos1)
    pos1 = 7
    pos2 = 10
    regfecha = Mid(strreg, pos1, pos2)
    Flog.writeline "Fecha:  " & regfecha

    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, Separador)
    'Hora = Trim(Mid(strReg, pos1, pos2 - pos1))
    pos1 = 18
    pos2 = 5
    Hora = Trim(Mid(strreg, pos1, pos2))
    Flog.writeline "Hora:  " & Hora
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If

    'pos1 = pos2 + 1
    'pos2 = InStr(pos1 + 1, strReg, ",")
    'NroReloj = Mid(strReg, pos1, pos2 - pos1)
    'NroRelojTxt = Mid(strReg, pos1, pos2 - pos1)
    pos1 = 24
    pos2 = 2
    nroreloj = Mid(strreg, pos1, pos2)
    nrorelojtxt = Mid(strreg, pos1, pos2)
    Flog.writeline "Reloj:  " & nrorelojtxt

    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. No se encontro el Reloj: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If

'El resto de la linea (Habilitada, Apellido y Nombre no tienen funcionalidad
    pos1 = 27
    pos2 = 2
    Habilitada = Mid(strreg, pos1, pos2)

    If Habilitada = "01" Then
        regestado = "I"
    Else
        regestado = "X"
    End If

    pos1 = 30
    pos2 = 30
    APYN = Mid(strreg, pos1, pos2)


'Marco todas las registraciones como entrada dado que el formato no lo distingue
entradasalida = "E"



'Validaciones

    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    StrSql = StrSql
    If Not objRs.EOF Then
        Ternro = objRs!Ternro
    Else
        Flog.writeline "Error. No se encontro la terjeta para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj: " & codReloj
        InsertaError 1, 33
        Exit Sub
    End If
    
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015

    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then

        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'" & regestado & "')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
End If
    
    
    
       
End Sub


Private Sub InsertaFormatoSpec(strreg As String)
'-------------------------------------------------------------------
'Autor: Sebastian Stremel - 18/02/2013
'Formato:
'    -Los Campos son de longitud fija:
'    Fijo GW00
'    Identifiación   3 Dig
'    Respuesta       1 char
'    Resultado       1 char
'    Incidencia      4 Dig
'    Tarjeta o cod. 16 Dig
'    AAAAMMDDhhmmss 14 lugares
'
'OBS
'   Tarjeta Habilitada: los valores son 01 Habilitada y 03 No Habilitada
'---------------
'Ejemplo:
'
'
'
'Ult. Modif:
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
'Dim fecha As Date
Dim regfecha As String
'Dim hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Habilitada As String
Dim APYN As String
Dim regestado
Dim separador As String

'variables seba 18/02/2013
Dim identificacion As String
Dim respuesta As String
Dim Resultado As String
Dim incidencia As String
Dim tarjeta As String
Dim Fecha As String
Dim Hora As String
Dim seg As String
Dim Anio As String
Dim Mes As String
Dim Dia As String
Dim Reg_Valida As Boolean

separador = " "


RegLeidos = RegLeidos + 1

pos1 = 5

identificacion = Mid(strreg, pos1, 3)
Flog.writeline "Identificacion:  " & identificacion

StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & identificacion & "'"
OpenRecordset StrSql, objRs
If objRs.EOF Then
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & identificacion & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Error. No se encontro el Reloj: " & identificacion
        Flog.writeline "SQL: " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
Else
    codReloj = objRs!relnro
    tipotarj = objRs!tptrnro
End If

'Reg estado se lo dejo fijo
regestado = "I"


'pos1 = 30
'pos2 = 30
'APYN = Mid(strReg, pos1, pos2)


pos1 = 10

respuesta = Mid(strreg, pos1, 1)
Flog.writeline "Respuesta:  " & respuesta

pos1 = 11

Resultado = Mid(strreg, pos1, 1)
Flog.writeline "Resultado:  " & Resultado
If ((Resultado <> "E") Or (Resultado <> "S") Or (Resultado <> "")) Then
    Resultado = ""
End If
entradasalida = Resultado
'If Not objFechasHoras.ValidarHora(Hora) Then
'    Flog.writeline " Error Hora: " & Hora
'    InsertaError 4, 38
'    Exit Sub
'End If

pos1 = 12
incidencia = Mid(strreg, pos1, 4)
Flog.writeline "Incidencia:  " & incidencia

pos1 = 16
tarjeta = Mid(strreg, pos1, 16)
Flog.writeline "Tarjeta:  " & tarjeta

pos1 = 32
Fecha = Mid(strreg, pos1, 8)

Anio = Mid(Fecha, 1, 4)
Mes = Mid(Fecha, 5, 2)
Dia = Mid(Fecha, 7, 2)
regfecha = Anio + "/" + Mes + "/" + Dia
Flog.writeline "Fecha:  " & Fecha

pos1 = 40
Hora = Mid(strreg, pos1, 4)
Flog.writeline "Hora:  " & Hora

pos1 = 44
seg = Mid(strreg, pos1, 2)
Flog.writeline "Seg:  " & seg

'Validaciones

StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & tarjeta & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
'StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & TipoTarj & " AND hstjnrotar = '" & tarjeta & "' AND (hstjfecdes <= " & RegFecha & ") AND ( (" & RegFecha & " <= hstjfechas) OR ( hstjfechas is null ))"

StrSql = StrSql
OpenRecordset StrSql, objRs
'StrSql = StrSql
If Not objRs.EOF Then
    Ternro = objRs!Ternro
Else
    Flog.writeline "Error. No se encontro la terjeta para el Legajo: " & tarjeta & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj: " & identificacion
    InsertaError 1, 33
    Exit Sub
End If

'Carmen Quintero - 15/05/2015
Reg_Valida = True
StrSql = "SELECT relnro FROM gti_rel_estr "
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
    'Valido que el reloj sea de control de acceso para el empleado
    StrSql = "SELECT ternro FROM his_estructura H "
    StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
    StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
    StrSql = StrSql & " AND ( h.ternro = " & Ternro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Reg_Valida = False
        Flog.writeline "    El reloj No está habilitado para el empleado "
    End If
End If
'Fin Carmen Quintero - 15/05/2015

StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
OpenRecordset StrSql, objRs
If objRs.EOF Then

    Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    If Reg_Valida Then
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'" & regestado & "')"
    Else
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    Call InsertarWF_Lecturas(Ternro, regfecha)
Else
    Flog.writeline " Registracion ya Existente"
    Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
    Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
    InsertaError 1, 92
End If



    
       
End Sub


Private Sub InsertaFormatoSpec2(strreg As String)
'-------------------------------------------------------------------
'Autor: FGZ - 03/06/2013
'Formato: Los Campos son separados por ;
'   tipo de registro Marcaje
'   número de registro
'   código de empleado
'   tipo de incidencia
'   identificado de terminal
'   lector 1
'** fecha formato YYYYMMDD
'** hora hhmmss
'** origen del marcaje
'   carácter fijo
'   resultado del marcaje
'** número de tarjeta

'OBS
'El origen es el código de reloj.
'Usar solo los campos con **, el resto se descarta.
'---------------
'Ejemplo:
'M3;46;-1;0;1;;20130520;135417;4097;1;2;0000000209000503;;;
'
'Ult. Modif:
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim regfecha As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Habilitada As String
Dim APYN As String
Dim regestado
Dim separador As String

Dim identificacion As String
Dim respuesta As String
Dim Resultado As String
Dim incidencia As String
Dim tarjeta As String
Dim Fecha As String
Dim Hora As String
Dim seg As String
Dim Anio As String
Dim Mes As String
Dim Dia As String
Dim Reg_Valida As Boolean

Dim Datos

RegLeidos = RegLeidos + 1

separador = ";"


Datos = Split(strreg, separador)

Fecha = Datos(6)
Hora = Left(Datos(7), 4)
nrorelojtxt = Datos(8)
tarjeta = Datos(11)

'Reloj
StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
OpenRecordset StrSql, objRs
If objRs.EOF Then
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Error. No se encontro el Reloj: " & nrorelojtxt
        Flog.writeline "SQL: " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
Else
    codReloj = objRs!relnro
    tipotarj = objRs!tptrnro
End If

'Reg estado se lo dejo fijo
regestado = "I"

'Aparentemente no distinguen entrada de salida
'Flog.writeline "Resultado:  " & Resultado
'If ((Resultado <> "E") Or (Resultado <> "S") Or (Resultado <> "")) Then
'    Resultado = ""
'End If
entradasalida = "" 'Resultado

'Acomodo la fecha
Flog.writeline "Fecha:  " & Fecha
Anio = Mid(Fecha, 1, 4)
Mes = Mid(Fecha, 5, 2)
Dia = Mid(Fecha, 7, 2)
regfecha = Anio + "/" + Mes + "/" + Dia

Flog.writeline "Hora:  " & Hora


'Validaciones
StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & tarjeta & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    Ternro = objRs!Ternro
Else
    Flog.writeline "Error. No se encontro la terjeta para el Legajo: " & tarjeta & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj: " & identificacion
    InsertaError 1, 33
    Exit Sub
End If

'Carmen Quintero - 15/05/2015
Reg_Valida = True
StrSql = "SELECT relnro FROM gti_rel_estr "
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
    'Valido que el reloj sea de control de acceso para el empleado
    StrSql = "SELECT ternro FROM his_estructura H "
    StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
    StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
    StrSql = StrSql & " AND ( h.ternro = " & Ternro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Reg_Valida = False
        Flog.writeline "    El reloj No está habilitado para el empleado "
    End If
End If
'Fin Carmen Quintero - 15/05/2015

StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
OpenRecordset StrSql, objRs
If objRs.EOF Then

    Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora

    If Reg_Valida Then
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'" & regestado & "')"
    Else
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    Call InsertarWF_Lecturas(Ternro, regfecha)
Else
    Flog.writeline " Registracion ya Existente"
    Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
    Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
    InsertaError 1, 92
End If

End Sub


Private Sub InsertaFormatoM193(strreg As String)
'-------------------------------------------------------------------
'Autor: FGZ - 04/06/2013
'Formato: Los Campos son separados por TABs
' <legajo>    <fecha>  <hora>  <reloj>  <entrada/salida>
'
'El primer campo es el legajo, por ejemplo: 20250
'La fecha no lleva el año completo: DD/MM/YY ejemplo 14/10/12
'El reloj, el modelo nuestro 2.
'La entrada se identifica con: 00, y la salida con: 01.

'OBS
'   Es similar al modelo 199

'---------------
'Ejemplo:
'00137 14/10/12 00:01 00001 00
'00284 14/10/12 00:01 00001 00
'00576 14/10/12 00:02 00001 00
'----------------
'Ult. Modif:
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim NroReloj_aux As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim NroTarj As Integer
Dim NroTarj_aux As String
Dim Aux  As String
Dim Reg_Valida As Boolean

    On Error GoTo MError
    
    separador = Chr(32)
    
    RegLeidos = RegLeidos + 1
    Flog.writeline "   - Registración --> " & strreg
    
    'Legajo
    pos1 = 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    
    'Fecha DD/MM/YY
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    fecha_aux = Mid(strreg, pos1, pos2 - pos1)
    
    'Hora HH:MM
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    Hora = Mid(strreg, pos1, pos2 - pos1)
    Hora = Replace(Hora, ":", "")
    
    'Nro reloj
    pos1 = pos2 + 1
    pos2 = InStr(pos1 + 1, strreg, separador)
    NroReloj_aux = Mid(strreg, pos1, pos2 - pos1)
        
    'Aux (Entrada/Salida)
    pos1 = pos2 + 1
    pos2 = Len(strreg)
    Aux = Mid(strreg, pos1, pos2)
    Select Case Aux
    Case "00"   'Entrada
        entradasalida = "E"
    Case "01"   'Salida
        entradasalida = "S"
    Case Else   'No reconocida
        entradasalida = ""
    End Select



'====================================================================
' Validar los parametros Levantados
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & NroLegajo
        InsertaError 1, 8
        GoTo Fin
    Else
        Ternro = objRs!Ternro
    End If
    
    'Que la fecha sea válida
    Dia = Mid(fecha_aux, 1, 2)
    Mes = Mid(fecha_aux, 4, 2)
    Anio = Mid(fecha_aux, 7, 4)
    If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
        Flog.writeline "       ****** Fecha no válida --> " & fecha_aux
        InsertaError 2, 4
        GoTo Fin
    Else
        regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline "       ****** Hora no válida --> " & Hora
        InsertaError 3, 38
        GoTo Fin
    End If
    
    'Busco el Reloj
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & NroReloj_aux & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el Reloj. SQL --> " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Reloj: " & codReloj
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        
    Else
        Flog.writeline "       ****** Registracion ya Existente"
        Flog.writeline "         Error Legajo: " & NroLegajo & " y Reloj: " & codReloj
        Flog.writeline "         Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    
End Sub


Private Sub InsertarFormatoRaffo(ByVal strreg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato de Lectura para Cliente CAJA DE ODONTOLOGOS
' Autor      : Sebastian Stremel
' Fecha      : 10/07/2013
'Formato: los campos estan separados por ;

'CAMPO FORMATO
'Fecha dd/MM/aaaa
'Hora HH:MM
'Tarjeta 8 dígitos - se debe rellenar con ceros a la izquierda
'Nro Reloj   2
'
'Ejemplo
'    2011/07/15,16:00,220152,1175
' ---------------------------------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim fecha_aux As String
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Long
Dim tipotarj As Integer
Dim NroTarj As Long
Dim NroTarjTexto As String
Dim NroTarj_aux As String
Dim descarte  As String
Dim Validar As Boolean
Dim tipo_reloj As Integer
Dim Origen As String
Dim Legajo As String
Dim Tipo_Marca As String
Dim TipoReg As Long
Dim campos
Dim Reg_Valida As Boolean

On Error GoTo MError

RegLeidos = RegLeidos + 1
Validar = True

Flog.writeline "   - Registración --> " & strreg


'BUSCO SEPARADOR | SI NO ESTA CONFIGURADO , POR DEFAULT
StrSql = "SELECT modseparador FROM modelo WHERE modnro = 172"
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    separador = objRs!modseparador
Else
    separador = ";"
End If
objRs.Close

campos = Split(strreg, separador)

'Formato
'------------------------
'   Fecha
'   Hora
'   Legajo
'   número de reloj
'------------------------

'---Fecha
fecha_aux = Trim(campos(0))

'---Hora
Hora = Trim(campos(1))

'---Empleado
NroTarj = Trim(campos(2))
NroTarjTexto = Trim(campos(2))

'---Nro reloj
nroreloj = Left(Trim(campos(3)), 2)

TipoReg = 0
        
If Validar Then
    Flog.writeline "Validaciones..."
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Error. Reloj no encontrado: " & nroreloj
        Flog.writeline "SQL: " & StrSql
        InsertaError 4, 32
        GoTo Fin
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    ''Que la fecha sea válida
    If IsDate(fecha_aux) Then
        regfecha = CDate(fecha_aux)
    Else
        Flog.writeline " Error Fecha: " & fecha_aux
        InsertaError 4, 4
        GoTo Fin
    
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 37
        GoTo Fin
    End If
    

    'Busco que el nro de tarjeta sea válido
    
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroTarj & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroTarjTexto & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. Tarjeta no encontrada para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " , Reloj: " & codReloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 1, 33
            GoTo Fin
        Else
            Ternro = objRs!Ternro
        End If
    Else
        Ternro = objRs!Ternro
    End If
    objRs.Close

     
    'Valido el nro de legajo
'    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & NroTarj
'    OpenRecordset StrSql, objRs
'    If Not objRs.EOF Then
'        Ternro = objRs!Ternro
'    Else
'        Flog.writeline "Error. No se encuentra el Legajo: " & NroTarj
'        Flog.writeline "SQL: " & StrSql
'        InsertaError 1, 33
'        GoTo Fin
'    End If
    
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
     
    'si la longitud es menor de 8 completo con ceros
    If Len(NroTarj) < 8 Then
        NroTarj = String(8 - Len(NroTarj), "0") & NroTarj
    End If
     
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & regfecha & "'  ; Hora: " & Hora & "  ; Nro. Tarjeta: " & NroTarj
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I','" & Tipo_Marca & "')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
                    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X','" & Tipo_Marca & "')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
  End If

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub
Private Sub InsertaFormatoClaxon(strlinea As String)
'---------------------------------
'Descripcion: Formato de lectura para cliente Claxon
'Autor      : Sebastian Stremel
'Fecha      : 16/03/2015
'CAMPO          FORMATO
'Fecha          date
'Legajo         Int
'Entrada        hh:mm
'Salida         hh:mm
'---------------------------------
Dim Datos
Dim indice As Long
Dim Ternro As String
Dim Fecha As String
Dim NroLegajo As String
Dim entrada As String
Dim Salida As String
Dim codReloj As Integer
Dim Reg_Valida As Boolean

'Dim Dia As String
'Dim Mes As String
'Dim Anio As String
'Dim nroreloj As String
'
'Dim indice As Long
'Dim codReloj As String
'Dim entradasalida As String
'Dim Tipo_Marca As String
'
'Dim campoHora As String
'Dim campoMinuto As String
'Dim Hora As String


On Error GoTo MError

RegLeidos = RegLeidos + 1
Flog.writeline "   - Registración --> " & strlinea
Datos = Split(strlinea, separador)

'Busco el Reloj
StrSql = "SELECT relnro FROM gti_reloj WHERE reldefault = -1"
OpenRecordset StrSql, objRs
If objRs.EOF Then
    Flog.writeline "       ****** No se encontro el Reloj Por Defecto "
    InsertaError 5, 32
    Exit Sub
Else
    codReloj = objRs!relnro
End If

If RegLeidos > 1 Then ' la primer linea no se toma
    For indice = 0 To UBound(Datos)
        Select Case indice
            '------------------------------------------------------------------------------------------------
            Case 0  'Fecha de registracion
                Fecha = Datos(indice)
                If EsNulo(Fecha) Then
                    Flog.writeline "la fecha de la registracion no puede ser nula"
                    GoTo Fin
                End If
            '------------------------------------------------------------------------------------------------
            Case 1 'Legajo del empleado
                StrSql = "SELECT ternro FROM empleado WHERE empleg = '" & Datos(indice) & "'"
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    Ternro = objRs!Ternro
                    NroLegajo = Datos(indice)
                Else
                    Flog.writeline "Error. No se encuentra el Legajo: " & Datos(indice)
                    Flog.writeline "SQL: " & StrSql
                    InsertaError 1, 33
                    GoTo Fin
                End If
            '------------------------------------------------------------------------------------------------
            Case 2 'Entrada HH:MM
                If Trim(Datos(indice)) <> "" Then
                    entrada = Datos(indice)
        
                    'Que la entrada sea válida
                    If Not objFechasHoras.ValidarHora(entrada) Then
                        Flog.writeline " Error entrada: " & entrada
                        InsertaError 4, 37
                        GoTo Fin
                    End If
                Else
                    Flog.writeline "Error campo entrada no informado."
                    'GoTo Fin
                End If
                
            '------------------------------------------------------------------------------------------------
            Case 3 'Salida HH:MM
                If Trim(Datos(indice)) <> "" Then
                    Salida = Datos(indice)
                    
                    'Que la salida sea válida
                    If Not objFechasHoras.ValidarHora(Salida) Then
                        Flog.writeline " Error entrada: " & Salida
                        InsertaError 4, 37
                        GoTo Fin
                    End If
                Else
                    Flog.writeline "Error campo salida no informado."
                    GoTo Fin
                End If
        
            '------------------------------------------------------------------------------------------------
        End Select
    Next
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(Fecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015

    'Verifico que a la hora de la entrada y la hora de la salida no halla ninguna registracion
    StrSql = " SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha)
    StrSql = StrSql & " AND ((reghora = '" & entrada & "' ) OR (reghora='" & Salida & "')) "
    StrSql = StrSql & " AND ternro = " & Ternro & " AND relnro=" & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        'Registro la entrada
        If Not EsNulo(entrada) Then
            If Reg_Valida Then
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal, relnro) "
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "("
                StrSql = StrSql & Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & entrada & "','E'," & codReloj & ")"
            Else
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & entrada & "','E'," & codReloj & ",'X')"
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & Fecha & "'  ; Hora: " & entrada
            
            Call InsertarWF_Lecturas(Ternro, Fecha)
            Flog.writeline "Inserto en temporal WF_Lecturas"
        End If
        
        'Registro la salida
        If Not EsNulo(Salida) Then
            If Reg_Valida Then
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal, relnro) "
                StrSql = StrSql & " VALUES "
                StrSql = StrSql & "("
                StrSql = StrSql & Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Salida & "','S'," & codReloj & ")"
            Else
                StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Salida & "','S'," & codReloj & ",'X')"
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & Fecha & "'  ; Hora: " & entrada
            
            Call InsertarWF_Lecturas(Ternro, Fecha)
            Flog.writeline "Inserto en temporal WF_Lecturas"
        End If

    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo
        Flog.writeline " Entrada: " & entrada & " - Salida: " & Salida & " - Fecha: '" & Fecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
    Flog.writeline "---------------------------------------------------------------------------------------"
    Flog.writeline ""

End If
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing


End Sub
Private Sub InsertaFormatoSuizo(ByVal strlinea As String)
'---------------------------------
'Descripcion: Formato de lectura para cliente Suizo
'Autor      : Sebastian Stremel
'Fecha      : 27/04/2015
'CAMPO          FORMATO
'Nro de legajo ==> long
'Nro de reloj  ==> Integer
'Dato fijo     ==> 00
'Fecha         ==> AAAA-MM-DD
'Hora Formato  ==> HH:MM:SS
'---------------------------------
Dim Datos
Dim indice As Long
Dim Ternro As String
Dim NroLegajo As String
Dim codReloj As Integer
Dim codFijo As String
Dim Fecha As String

Dim entrada As String
Dim Salida As String

Dim campoHora As String
Dim campoMinuto As String
Dim Hora As String
Dim Reg_Valida As Boolean

'Dim Dia As String
'Dim Mes As String
'Dim Anio As String
'Dim nroreloj As String
'
'Dim indice As Long
'Dim codReloj As String
'Dim entradasalida As String
'Dim Tipo_Marca As String
'



On Error GoTo MError

RegLeidos = RegLeidos + 1
Flog.writeline "   - Registración --> " & strlinea
Datos = Split(strlinea, separador)

'If RegLeidos > 1 Then ' la primer linea no se toma
For indice = 0 To UBound(Datos)
    Select Case indice
        '------------------------------------------------------------------------------------------------
        Case 0  'Legajo del empleado
            If Trim(Datos(indice)) <> "" Then
                StrSql = "SELECT ternro FROM empleado WHERE empleg = '" & Datos(indice) & "'"
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    Ternro = objRs!Ternro
                    NroLegajo = Datos(indice)
                Else
                    Flog.writeline "Error. No se encuentra el Legajo: " & Datos(indice)
                    Flog.writeline "SQL: " & StrSql
                    InsertaError 1, 33
                    GoTo Fin
                End If
            Else
                Flog.writeline "Legajo no informado."
                GoTo Fin
            End If
            
        Case 1 'Reloj
            Flog.writeline "Busco el reloj"
            If Trim(Datos(indice)) <> "" Then
                StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & Datos(indice) & "'"
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    codReloj = objRs!relnro
                Else
                    Flog.writeline "Error. Reloj no encontrado: " & Datos(indice)
                    Flog.writeline "SQL: " & StrSql
                    InsertaError 4, 32
                    GoTo Fin
                End If
            Else
                Flog.writeline "Codigo reloj no informado."
                GoTo Fin
            End If
        
        Case 2 'Dato fijo numerico
            If Trim(Datos(indice)) <> "" Then
                codFijo = Datos(indice)
            Else
                Flog.writeline "Codigo fijo no informado."
                GoTo Fin
            End If
        
        Case 3 'Fecha Formato AAAA-MM-DD
            If Trim(Datos(indice)) <> "" Then
                If Len(Datos(indice)) = 19 Then
                    If IsDate(Mid(Datos(indice), 1, 10)) Then
                        Fecha = Mid(Datos(indice), 1, 10)
                    Else
                        Flog.writeline " Error en la fecha: " & Fecha
                        InsertaError 4, 37
                        GoTo Fin
                    End If
                    
                    Hora = Mid(Datos(indice), 12, 5)
                    If Not objFechasHoras.ValidarHora(Hora) Then
                        Flog.writeline " Error en la Hora: " & Hora
                        InsertaError 4, 37
                        GoTo Fin
                    Else
                        Hora = Replace(Hora, ":", "")
                    End If
                    
                    
                Else
                    Flog.writeline "campo fecha y hora mal informado no cumple la longitud."
                End If
            Else
                Flog.writeline "Campo Hora no informado."
                GoTo Fin
            End If
    End Select
Next

'Carmen Quintero - 15/05/2015
Reg_Valida = True
StrSql = "SELECT relnro FROM gti_rel_estr "
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
    'Valido que el reloj sea de control de acceso para el empleado
    StrSql = "SELECT ternro FROM his_estructura H "
    StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
    StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
    StrSql = StrSql & " AND ( h.ternro = " & Ternro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(Fecha) & "))"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Reg_Valida = False
        Flog.writeline "    El reloj No está habilitado para el empleado "
    End If
End If
'Fin Carmen Quintero - 15/05/2015

'Verifico que en la fecha y hora no halla registracion
StrSql = " SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha)
StrSql = StrSql & " AND reghora = '" & Hora & "'  "
StrSql = StrSql & " AND ternro = " & Ternro & " AND relnro=" & codReloj
OpenRecordset StrSql, objRs
If objRs.EOF Then
    If Reg_Valida Then
        'Inserto la registracion
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora, relnro) "
        StrSql = StrSql & " VALUES "
        StrSql = StrSql & "("
        StrSql = StrSql & Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & codReloj & ")"
    
    Else
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "'," & codReloj & ",'X')"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & Fecha & "'  ; Hora: " & entrada
    
    Call InsertarWF_Lecturas(Ternro, Fecha)
    Flog.writeline "Inserto en temporal WF_Lecturas"
Else
    Flog.writeline " Registracion ya Existente"
    Flog.writeline " Error Legajo: " & NroLegajo
    Flog.writeline " Fecha: " & Fecha & " - Hora: " & Hora & "'"
    InsertaError 1, 92
End If
Flog.writeline "Linea Procesada"
Flog.writeline "---------------------------------------------------------------------------------------"
Flog.writeline ""

'End If
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing


End Sub
Private Sub InsertaFormatoClaxonV2(strlinea As String)
'---------------------------------
'Descripcion: Formato de lectura para cliente Claxon
'Autor      : Sebastian Stremel
'Fecha      : 27/04/2015
'CAMPO          FORMATO
'Fecha          date
'Legajo         Int
'Hora           string HH:MM
'---------------------------------
Dim Datos
Dim indice As Long
Dim Ternro As String
Dim Fecha As String
Dim NroLegajo As String
Dim Hora As String
Dim codReloj As Integer
Dim Reg_Valida As Boolean


On Error GoTo MError

RegLeidos = RegLeidos + 1
Flog.writeline "   - Registración --> " & strlinea
Datos = Split(strlinea, separador)

'Busco el Reloj
StrSql = "SELECT relnro FROM gti_reloj WHERE reldefault = -1"
OpenRecordset StrSql, objRs
If objRs.EOF Then
    Flog.writeline "       ****** No se encontro el Reloj Por Defecto "
    InsertaError 5, 32
    Exit Sub
Else
    codReloj = objRs!relnro
End If

If RegLeidos > 1 Then ' la primer linea no se toma
    For indice = 0 To UBound(Datos)
        Select Case indice
            '------------------------------------------------------------------------------------------------
            Case 0  'Fecha de registracion
                If Trim(Datos(indice)) <> "" Then
                    Fecha = Datos(indice)
                    If Not IsDate(Fecha) Then
                        Flog.writeline "la fecha de la registracion es incorrecta."
                        GoTo Fin
                    End If
                Else
                    Flog.writeline "Error campo fecha no informado."
                    GoTo Fin
                End If
                
            '------------------------------------------------------------------------------------------------
            Case 1 'Legajo del empleado
                If Trim(Datos(indice)) <> "" Then
                    StrSql = "SELECT ternro FROM empleado WHERE empleg = '" & Datos(indice) & "'"
                    OpenRecordset StrSql, objRs
                    If Not objRs.EOF Then
                        Ternro = objRs!Ternro
                        NroLegajo = Datos(indice)
                    Else
                        Flog.writeline "Error. No se encuentra el Legajo: " & Datos(indice)
                        Flog.writeline "SQL: " & StrSql
                        InsertaError 1, 33
                        GoTo Fin
                    End If
                Else
                    Flog.writeline "Error campo legajo no informado."
                    GoTo Fin
                End If
            '------------------------------------------------------------------------------------------------
            Case 2 'Hora HH:MM
                If Trim(Datos(indice)) <> "" Then
                    Hora = Datos(indice)
        
                    'Que la hora sea válida
                    If Not objFechasHoras.ValidarHora(Hora) Then
                        Flog.writeline " Error campo Hora: " & Hora
                        InsertaError 4, 37
                        GoTo Fin
                    Else
                        Hora = Replace(Hora, ":", "")
                    End If
                Else
                    Flog.writeline "Error campo Hora no informado."
                    GoTo Fin
                End If

        End Select
    Next
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(Fecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015

    'Verifico que a la hora de la entrada y la hora de la salida no halla ninguna registracion
    StrSql = " SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha)
    StrSql = StrSql & " AND reghora = '" & Hora & "'  "
    StrSql = StrSql & " AND ternro = " & Ternro & " AND relnro=" & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        If Reg_Valida Then
            'Registro la registracion
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora, regentsal, relnro) "
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "("
            StrSql = StrSql & Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','E'," & codReloj & ")"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','E'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & Fecha & "'  ; Hora: " & Hora
        
        Call InsertarWF_Lecturas(Ternro, Fecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo
        Flog.writeline " Fecha: " & Fecha & " - Hora: " & Hora
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
    Flog.writeline "---------------------------------------------------------------------------------------"
    Flog.writeline ""

End If
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing


End Sub

Private Sub InsertaFormatoSalto(strlinea As String)
'---------------------------------
'Descripcion: Formato de lectura para cliente Salto Grande - modelo 159
'Autor      : Miriam Ruiz
'Fecha      : 07/05/2015
'CAMPO          FORMATO
'Fecha          date
'Legajo         Int
'Hora           string HH:MM
'---------------------------------
Dim Datos
Dim indice As Long
Dim Ternro As String
Dim Fecha As String
Dim NroLegajo As String
Dim Hora As String
Dim codReloj As Integer
Dim nroreloj As String
Dim Anio As String
Dim Mes As String
Dim Dia As String
Dim Reg_Valida As Boolean

On Error GoTo MError

RegLeidos = RegLeidos + 1
Flog.writeline "   - Registración --> " & strlinea
Datos = Split(strlinea, separador)

'Busco el Reloj
'StrSql = "SELECT relnro FROM gti_reloj WHERE reldefault = -1"
'OpenRecordset StrSql, objRs
'If objRs.EOF Then
'    Flog.writeline "       ****** No se encontro el Reloj Por Defecto "
'    InsertaError 5, 32
'    Exit Sub
'Else
'    codReloj = objRs!relnro
'End If

If RegLeidos > 0 Then
    For indice = 0 To UBound(Datos)
        Select Case indice
            '------------------------------------------------------------------------------------------------
            Case 6 'Fecha de registracion
                If Trim(Datos(indice)) <> "" Then
                    Fecha = Datos(indice)
                    Flog.writeline "Fecha:  " & Fecha
                    Anio = Mid(Fecha, 1, 4)
                    Mes = Mid(Fecha, 5, 2)
                    Dia = Mid(Fecha, 7, 2)
                    Fecha = Anio + "/" + Mes + "/" + Dia

                    If Not IsDate(Fecha) Then
                        Flog.writeline "la fecha de la registracion es incorrecta."
                        GoTo Fin
                    End If
                Else
                    Flog.writeline "Error campo fecha no informado."
                    GoTo Fin
                End If
                
            '------------------------------------------------------------------------------------------------
            Case 2 'Legajo del empleado
                If Trim(Datos(indice)) <> "" Then
                    StrSql = "SELECT ternro FROM empleado WHERE empleg = '" & Datos(indice) & "'"
                    OpenRecordset StrSql, objRs
                    If Not objRs.EOF Then
                        Ternro = objRs!Ternro
                        NroLegajo = Datos(indice)
                    Else
                        Flog.writeline "Error. No se encuentra el Legajo: " & Datos(indice)
                        Flog.writeline "SQL: " & StrSql
                        InsertaError 1, 33
                        GoTo Fin
                    End If
                Else
                    Flog.writeline "Error campo legajo no informado."
                    GoTo Fin
                End If
            '------------------------------------------------------------------------------------------------
             Case 4 'R, reloj, codigo externo del reloj
                 If Trim(Datos(indice)) <> "" Then
                     nroreloj = Left(Trim(Datos(indice)), 2)
                      Flog.writeline "Busco el reloj"
                      StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
                       OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    codReloj = objRs!relnro
                Else
                    Flog.writeline "Error. Reloj no encontrado: " & nroreloj
                    Flog.writeline "SQL: " & StrSql
                    InsertaError 5, 32
                    Exit Sub
                End If
            Else
                Flog.writeline "Error. Reloj no informado."
                Exit Sub
            End If
            '------------------------------------------------------------------------------------------------
            Case 7 'Hora HHMMss
                If Trim(Datos(indice)) <> "" Then
                    Hora = Datos(indice)
                    If Len(Hora) > 5 Then
                        Hora = Left(Hora, 4)
                    Else
                        Hora = "0" & Left(Hora, 3)
                    End If
                    'Que la hora sea válida
                    If Not objFechasHoras.ValidarHora(Hora) Then
                        Flog.writeline " Error campo Hora: " & Hora
                        InsertaError 4, 37
                        GoTo Fin
                    Else
                        Hora = Replace(Hora, ":", "")
                    End If
                Else
                    Flog.writeline "Error campo Hora no informado."
                    GoTo Fin
                End If

        End Select
    Next
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(Fecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015

    'Verifico que a la hora de la entrada y la hora de la salida no halla ninguna registracion
    StrSql = " SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha)
    StrSql = StrSql & " AND reghora = '" & Hora & "'  "
    StrSql = StrSql & " AND ternro = " & Ternro & " AND relnro=" & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        If Reg_Valida Then
            'Registro la registracion
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora, regentsal, relnro) "
            StrSql = StrSql & " VALUES "
            StrSql = StrSql & "("
            StrSql = StrSql & Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','E'," & codReloj & ")"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','E'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & Fecha & "'  ; Hora: " & Hora
        
        Call InsertarWF_Lecturas(Ternro, Fecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo
        Flog.writeline " Fecha: " & Fecha & " - Hora: " & Hora
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
    Flog.writeline "---------------------------------------------------------------------------------------"
    Flog.writeline ""

End If
Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing


End Sub

Private Sub InsertaFormatoPollPar(strreg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato de Lectura para Cliente POLLPAR
' Autor      : LED
' Fecha      : 03/11/2014
'Formato:

'CAMPO          FORMATO
'ST             locacion   Pollpar, Transchaco          se ignora eset campo
'R              nro de reloj                            codigo externo del reloj
'TM             evento                                  entrada , salida, almuerzo, regreso
'legajo         Legajo                                  Legajo, no lleva tarjeta.
'campo desconocido                                      se ignora
'HH             Hora                                    formato 24 hs
'MM             Minuto                                  minutos
'M              Mes                                     Mes
'DD             Dia                                     Dia
'AA             Año                                     Formato ultimos dos digitos

'
'Ejemplo
'  1, 1, 1,0000000194,0000000000,11,15, 9,25,13,  ,  ,  ,  ,  ,          ,          ,        ,
' ---------------------------------------------------------------------------------------------

Dim Ternro As String
Dim Fecha As Date
Dim Dia As String
Dim Mes As String
Dim Anio As String
Dim nroreloj As String
Dim Datos
Dim indice As Long
Dim codReloj As String
Dim entradasalida As String
Dim Tipo_Marca As String
Dim Legajo As String
Dim campoHora As String
Dim campoMinuto As String
Dim Hora As String
Dim NroLegajo As String
Dim Reg_Valida As Boolean

On Error GoTo MError

RegLeidos = RegLeidos + 1

Flog.writeline "   - Registración --> " & strreg


Datos = Split(strreg, separador)


For indice = 0 To UBound(Datos)
    Select Case indice
        '------------------------------------------------------------------------------------------------
        Case 0  'ST, se ignora este campo
        '------------------------------------------------------------------------------------------------
        Case 1 'R, reloj, codigo externo del reloj
            If Trim(Datos(indice)) <> "" Then
                nroreloj = Left(Trim(Datos(indice)), 2)
                Flog.writeline "Busco el reloj"
                StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    codReloj = objRs!relnro
                Else
                    Flog.writeline "Error. Reloj no encontrado: " & nroreloj
                    Flog.writeline "SQL: " & StrSql
                    InsertaError 4, 32
                    GoTo Fin
                End If
            Else
                Flog.writeline "Error. Reloj no informado."
                GoTo Fin
            End If
        '------------------------------------------------------------------------------------------------
        Case 2 'TM, evento: entrada , salida, almuerzo, regreso
            If Trim(Datos(indice)) <> "" Then
                Flog.writeline "Busco datos del tipo de registracion"
                StrSql = " SELECT tiporeges, tiporegcod FROM gti_tiporeg where upper(tiporegcod) = '" & UCase(Trim(Datos(indice))) & "'"
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    entradasalida = objRs!tiporeges
                    Tipo_Marca = objRs!tiporegcod
                Else
                    Flog.writeline "Error. Tipo de registracion no encontrada, codigo: " & Trim(Datos(indice))
                    Flog.writeline "SQL: " & StrSql
                    InsertaError 4, 32
                    GoTo Fin
                End If
            Else
                Flog.writeline "Error. Tipo de Registracion no informado."
                GoTo Fin
            End If
        
        '------------------------------------------------------------------------------------------------
        Case 3 'legajo, no lleva tarjeta.
            If Trim(Datos(indice)) <> "" Then
                StrSql = "SELECT ternro FROM empleado WHERE empleg = '" & Datos(indice) & "'"
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    Ternro = objRs!Ternro
                    NroLegajo = Datos(indice)
                Else
                    Flog.writeline "Error. No se encuentra el Legajo: " & Datos(indice)
                    Flog.writeline "SQL: " & StrSql
                    InsertaError 1, 33
                    GoTo Fin
                End If
            Else
                Flog.writeline "Error. Legajo no informado."
                GoTo Fin
            End If
        
        '------------------------------------------------------------------------------------------------
        Case 4 'campo desconocido, se ignora
        '------------------------------------------------------------------------------------------------
        Case 5 'HH, horas
            If Trim(Datos(indice)) <> "" Then
                campoHora = Right("00" & Trim(Datos(indice)), 2)
            Else
                Flog.writeline "Error. Hora no informado."
                GoTo Fin
            End If
        '------------------------------------------------------------------------------------------------
        Case 6 'MM, Minutos
            If Trim(Datos(indice)) <> "" Then
                campoMinuto = Right("00" & Trim(Datos(indice)), 2)
            Else
                Flog.writeline "Error. Minutos no informados."
                GoTo Fin
            End If
    
            'Que la hora sea válida
            Hora = campoHora & campoMinuto
            If Not objFechasHoras.ValidarHora(Hora) Then
                Flog.writeline " Error Hora: " & Hora
                InsertaError 4, 37
                GoTo Fin
            End If
            
        '------------------------------------------------------------------------------------------------
        Case 7 'M, Mes
            If Trim(Datos(indice)) <> "" Then
                Mes = Right("00" & Trim(Datos(indice)), 2)
            Else
                Flog.writeline "Error. Mes no informado."
                GoTo Fin
            End If
            
        '------------------------------------------------------------------------------------------------
        Case 8 'DD, Dia
            If Trim(Datos(indice)) <> "" Then
                Dia = Right("00" & Trim(Datos(indice)), 2)
            Else
                Flog.writeline "Error. Dia no informado."
                GoTo Fin
            End If

        '------------------------------------------------------------------------------------------------
        Case 9 'AA, Año
            If Trim(Datos(indice)) <> "" Then
                Anio = "20" & Right("00" & Trim(Datos(indice)), 2)
            Else
                Flog.writeline "Error. Año no informado."
                GoTo Fin
            End If
            
            Fecha = Dia & "/" & Mes & "/" & Anio
            If Not IsDate(Fecha) Then
                Flog.writeline " Error Fecha: " & Fecha
                InsertaError 4, 4
                GoTo Fin
            Else
                Flog.writeline "Fecha correctamente obtenida: " & Fecha & "."
            End If
    End Select
Next

'Carmen Quintero - 15/05/2015
Reg_Valida = True
StrSql = "SELECT relnro FROM gti_rel_estr "
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
    'Valido que el reloj sea de control de acceso para el empleado
    StrSql = "SELECT ternro FROM his_estructura H "
    StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
    StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
    StrSql = StrSql & " AND ( h.ternro = " & Ternro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(Fecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(Fecha) & "))"
    'Flog.writeline "sql: " & StrSql
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Reg_Valida = False
        Flog.writeline "    El reloj No está habilitado para el empleado "
    End If
End If
'Fin Carmen Quintero - 15/05/2015

StrSql = " SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(Fecha) & " AND reghora = '" & Hora & "'" & _
         " AND ternro = " & Ternro & " AND relnro = " & codReloj
OpenRecordset StrSql, objRs
If objRs.EOF Then
    If Reg_Valida Then
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I','" & Tipo_Marca & "')"
    Else
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(Fecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X','" & Tipo_Marca & "')"
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "               INSERTO REGISTRACION - Legajo: " & NroLegajo & "  ; Fecha: '" & Fecha & "'  ; Hora: " & Hora
    
    Call InsertarWF_Lecturas(Ternro, Fecha)
    Flog.writeline "Inserto en temporal WF_Lecturas"
Else
    Flog.writeline " Registracion ya Existente"
    Flog.writeline " Error Legajo: " & NroLegajo & " " & codReloj
    Flog.writeline " Hora: " & Hora & " - Fecha: '" & Fecha & "'"
    InsertaError 1, 92
End If
Flog.writeline "Linea Procesada"
Flog.writeline "---------------------------------------------------------------------------------------"
Flog.writeline ""

Fin:
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
    Exit Sub
    
MError:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * 3) & " Error " & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 3) & "**********************************************************"
    Flog.writeline
    
    If objRs.State = adStateOpen Then objRs.Close
    Set objRs = Nothing
End Sub

Private Sub InsertaFormato31(strreg As String)

' ---------------------------------------------------------------------------------------------
' Descripcion:  - CAS-21338 - Owens - Formato de Lectura para Zk Software x628, Zk SoftwareF8, Zk Software MA300
' Autor      : Mauricio Zwenger
' Fecha      : 16/09/2013
' Formato    : los campos estan separados por " " (espacio);

'CAMPO FORMATO
'Legajo             :9 digitos formato "000000000"
'Fecha              :dd/MM/aa
'Hora               :HH:MM
'Entrada/Salida     :I=entrada, O=salida
'Nro Reloj          :un digito
'
'Ejemplo
'    000009008 27/08/2013 13:46 O 6
'    000009590 27/08/2013 18:10 I 6
' ---------------------------------------------------------------------------------------------
'modificado : 02/12/2014
'MDZ - CAS-27487 - se valida variante de formato para reloj Surplast
'
'Ejemplo
'   Legajo   Fecha      Hora     Estado  Reloj
'   26722831 14/07/2014 15:07:32 Entrada 1
'   26722831 15/07/2014 00:30:49 Entrada 1

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Dia
Dim Mes
Dim Anio
Dim Reg_Valida As Boolean

    On Error GoTo ME_Local
    
    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Fecha = Mid(strreg, pos1, pos2 - pos1)
    
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    
    'MDZ - si viene con los segundos se los saco
    If Len(Hora) = 8 Then
        Hora = Left(Hora, 5)
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = UCase(Trim(Mid(strreg, pos1, pos2 - pos1)))
    
    'Valido "I" o "Entrada" como una Entrada y "O" o "Salida" como una Salida
    If entradasalida = "I" Or UCase(entradasalida) = "ENTRADA" Then
        entradasalida = "E"
    ElseIf entradasalida = "O" Or UCase(entradasalida) = "SALIDA" Then
        entradasalida = "S"
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    nroreloj = Trim(Mid(strreg, pos1))
    nrorelojtxt = Trim(Mid(strreg, pos1))
    
    '====================================================================
    ' Validar los parametros Levantados
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & CLng(NroLegajo)
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el legajo --> " & CLng(NroLegajo)
        InsertaError 1, 8
        GoTo Fin
    Else
        Ternro = objRs!Ternro
    End If
    
    'Que la fecha sea válida
    Dia = Mid(Fecha, 1, 2)
    Mes = Mid(Fecha, 4, 2)
    Anio = Mid(Fecha, 7, 4)
    If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
        Flog.writeline "       ****** Fecha no válida --> " & Fecha
        InsertaError 2, 4
        GoTo Fin
    Else
        If CDate(Dia & "/" & Mes & "/" & Anio) Then
            regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
        Else
            Flog.writeline "       ****** Fecha no válida --> " & Fecha
            InsertaError 2, 4
            GoTo Fin
        End If
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline "       ****** Hora no válida --> " & Hora
        InsertaError 3, 38
        GoTo Fin
    End If
    
    'Busco el Reloj
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el Reloj. SQL --> " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        
    Else
        Flog.writeline " Registracion ya Existente "
        Flog.writeline " Hora: " & Hora & " - Fecha: " & regfecha
        InsertaError 1, 92
    End If
        
Fin:
Exit Sub
ME_Local:
    HuboError = True
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    GoTo Fin
End Sub


Private Sub InsertaFormato32(strreg As String)

' ---------------------------------------------------------------------------------------------
' Descripcion:  - CAS-21337 - SAN CAMILO - GTI-QA-Formato de Reloj para reloj marca Lenox
' Autor      : Mauricio Zwenger
' Fecha      : 17/09/2013
' Formato    : los campos estan separados por " " (espacio);

'CAMPO FORMATO
'huella             :6 digitos formato "000000"
'Fecha              :ddMMaa
'Hora               :HH:MM
'Entrada/Salida     :ENT=entrada, SAL=salida
'
'Ejemplo
'    22321 060213 13:44 ENT
'    35267 060213 06:46 ENT
'    35267 060213 14:14 SAL
'    36098 060213 05:44 ENT
' ---------------------------------------------------------------------------------------------


Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim FechaStr As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Dia
Dim Mes
Dim Anio
Dim Huella
Dim Reg_Valida As Boolean

    On Error GoTo ME_Local
    
    RegLeidos = RegLeidos + 1
    
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    Huella = Mid(strreg, pos1, pos2 - pos1)
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    FechaStr = Mid(strreg, pos1 + 1, pos2 - pos1)
    
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    
    
    pos1 = pos2
    pos2 = Len(Trim(strreg)) + 1 'InStr(pos1 + 1, strreg, " ")
    entradasalida = Trim(UCase(Mid(strreg, pos1, pos2 - pos1)))
    
    Flog.writeline RegLeidos & ":    " & entradasalida
    
    If entradasalida = "ENT" Then
        entradasalida = "E"
    ElseIf entradasalida = "SAL" Then
        entradasalida = "S"
    Else
        Flog.writeline "       ****** Formato no válido de EntradaSalida --> " & entradasalida
        InsertaError 1, 4
        GoTo Fin
    End If
    
      
    
    '====================================================================
    ' Validar los parametros Levantados
       
    'Que la fecha sea válida
    Dia = Mid(FechaStr, 1, 2)
    Mes = Mid(FechaStr, 3, 2)
    Anio = Mid(FechaStr, 5, 2)
    If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
        Flog.writeline "       ****** Fecha no válida --> " & FechaStr
        InsertaError 3, 4
        GoTo Fin
    Else
        If IsDate(Dia & "/" & Mes & "/" & Anio) Then
            regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
        Else
            Flog.writeline "       ****** Fecha no válida --> " & Fecha
            InsertaError 3, 4
            GoTo Fin
        End If
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline "       ****** Hora no válida --> " & Hora
        InsertaError 4, 38
        GoTo Fin
    End If
    
    'Busco el Reloj
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE reldefault = -1"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el Reloj Por Defecto "
        InsertaError 5, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    'busco la tarjeta/huella para obtener el ternro
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & Huella & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & Huella & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      OpenRecordset StrSql, objRs
      If Not objRs.EOF Then
         Ternro = objRs!Ternro
      Else
        StrSql = "SELECT ternro FROM gti_histarjeta WHERE convert(bigint,hstjnrotar) = " & CLng(Huella) & " AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
           Ternro = objRs!Ternro
        Else
           Flog.writeline "Error. Nro de Huella no encontrado : " & Huella & ", Tipo de tarjeta: " & tipotarj & " y codigo de reloj:  " & codReloj
           Flog.writeline "SQL: " & StrSql
           InsertaError 1, 33
           Exit Sub
        End If
      End If
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        
    Else
        Flog.writeline " Registracion ya Existente "
        Flog.writeline " Hora: " & Hora & " - Fecha: " & regfecha
        InsertaError 1, 92
    End If
        
Fin:
Exit Sub
ME_Local:
    HuboError = True
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    GoTo Fin
End Sub

Private Sub InsertaFormato33(strreg As String)

' ---------------------------------------------------------------------------------------------
' Descripcion:  - CAS-21170 - SGS - Interfase de levantamiento de registraciones con parte de movilidad
' Autor      : Mauricio Zwenger
' Fecha      : 11/10/2013
' Formato    : los campos estan separados por ","

'CAMPO FORMATO
'Legajo             :
'Fecha              :dd/MM/aaaa
'Hora Inicio        :HH:MM
'Hora Fin           :HH:MM
'estructura1        :
'estructura2        :
'estructura3        :
'
'Ejemplo
'    50276,01/08/2013,08:00,16:00,7000,01,780795
'    50276,02/08/2013,08:00,12:00,7000,01,780795
'    52013,27/07/2013,06:00,18:00,3000,01,738629
'    52013,28/07/2013,06:00,18:00,3000,01,738629
'    52013,29/07/2013,06:00,18:00,3000,01,738629
'    52013,30/07/2013,06:00,18:00,3000,01,738629
'    52013,02/08/2013,06:00,18:00,3000,01,738629
' ---------------------------------------------------------------------------------------------


Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim FechaStr As String
Dim HoraIni As String
Dim HoraSal As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Estr1 As String
Dim Estr2 As String
Dim Estr3 As String

Dim Testr1 As String
Dim Testr2 As String
Dim Testr3 As String

Dim Estrnro1 As Long
Dim Estrnro2 As Long
Dim Estrnro3 As Long

Dim regnro1
Dim regnro2
Dim ol
Dim Reg_Valida As Boolean

'tipos de estructuras
Testr1 = 5
Testr2 = 102
Testr3 = 0

Dim Registro

    On Error GoTo ME_Local
    
    RegLeidos = RegLeidos + 1
    
    
    If UBound(Split(strreg, ",")) = 6 Then
        
        Registro = Split(strreg, ",")

        NroLegajo = Registro(0)
        FechaStr = Registro(1)
        HoraIni = Registro(2)
        HoraSal = Registro(3)
        Estr1 = Registro(4)
        Estr2 = Registro(5)
        ol = Registro(6)
        
    ElseIf UBound(Split(strreg, ",")) = 4 Then
        
        NroLegajo = Registro(0)
        FechaStr = Registro(1)
        HoraIni = Registro(2)
        HoraSal = Registro(3)
        
    Else
    
        Flog.writeline "       ****** Formato de Registro no válido ****** "
        InsertaError 1, 4
        GoTo Fin
    End If
    
       
    
    
    
    '====================================================================
    ' Validar los parametros Levantados
       
    'Que la fecha sea válida
    If Not IsDate(FechaStr) Then
        Flog.writeline "       ****** Fecha no válida --> " & FechaStr
        InsertaError 3, 4
        GoTo Fin
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(HoraIni) Then
        Flog.writeline "       ****** Hora de Inicio no válida --> " & HoraIni
        InsertaError 4, 38
        GoTo Fin
    End If
    
    If Not objFechasHoras.ValidarHora(HoraSal) Then
        Flog.writeline "       ****** Hora de Salida no válida --> " & HoraSal
        InsertaError 4, 38
        GoTo Fin
    End If
    
    'Busco el Reloj
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE reldefault = -1"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el Reloj Por Defecto "
        InsertaError 5, 32
        GoTo Fin
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    'busco el legajo para obtener el ternro
    StrSql = "select ternro from empleado WHERE empleg=" & NroLegajo
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
          Ternro = objRs!Ternro
    Else
        Flog.writeline "       ****** No se encontro Legajo --> " & NroLegajo
        InsertaError 5, 32
        GoTo Fin
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(CDate(FechaStr)) & " AND (htethasta is null or htethasta >= " & ConvFecha(CDate(FechaStr)) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    'genero la registracion de Ingreso
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(CDate(FechaStr)) & " AND reghora = '" & HoraIni & "' AND ternro = " & Ternro & " AND regentsal = 'E' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(CDate(FechaStr)) & ",'" & HoraIni & "','E'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(CDate(FechaStr)) & ",'" & HoraIni & "','E'," & codReloj & ",'X')"
        End If
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, CDate(FechaStr))
        
    Else
        Flog.writeline " Registracion INGRESO ya Existente "
        Flog.writeline " Hora: " & HoraIni & " - Fecha: " & FechaStr
        InsertaError 1, 92
         GoTo Fin
    End If
    'obtengo el id de la registracion de ingreso
    StrSql = "SELECT regnro FROM gti_registracion WHERE regfecha = " & ConvFecha(CDate(FechaStr)) & " AND reghora = '" & HoraIni & "' AND ternro = " & Ternro & " AND regentsal = 'E' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        regnro1 = objRs!Regnro
    Else
        Flog.writeline " Error al generar la registracion de Ingreso  "
        Flog.writeline " Hora: " & HoraIni & " - Fecha: " & FechaStr
        InsertaError 1, 92
         GoTo Fin
    End If
    
    
    'genero la registracion de salida
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(CDate(FechaStr)) & " AND reghora = '" & HoraSal & "' AND ternro = " & Ternro & " AND regentsal = 'S' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(CDate(FechaStr)) & ",'" & HoraSal & "','S'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(CDate(FechaStr)) & ",'" & HoraSal & "','S'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, CDate(FechaStr))
        
    Else
        Flog.writeline " Registracion SALIDA ya Existente "
        Flog.writeline " Hora: " & HoraSal & " - Fecha: " & FechaStr
        InsertaError 1, 92
         GoTo Fin
    End If
    'obtengo el id de la registracion de salida
    StrSql = "SELECT regnro FROM gti_registracion WHERE regfecha = " & ConvFecha(CDate(FechaStr)) & " AND reghora = '" & HoraIni & "' AND ternro = " & Ternro & " AND regentsal = 'E' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        regnro2 = objRs!Regnro
    Else
        Flog.writeline " Error al generar la registracion de Ingreso  "
        Flog.writeline " Hora: " & HoraIni & " - Fecha: " & FechaStr
        InsertaError 1, 92
         GoTo Fin
    End If
    
    
    'busco las estructuras segun los codigos externos y los tipos de estructuras
    StrSql = "select estrnro from estructura where tenro=" & Testr1 & " and estrcodext='" & Estr1 & "'"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Estrnro1 = objRs!estrnro
    Else
        Flog.writeline " No se encontro Estructura "
        Flog.writeline " Tipo: " & Testr1 & " - Codigo Externo: " & Estr1
        InsertaError 1, 92
         GoTo Fin
    End If
    
    StrSql = "select estrnro from estructura where tenro=" & Testr2 & " and estrcodext='" & Estr2 & "'"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Estrnro2 = objRs!estrnro
    Else
        Flog.writeline " No se encontro Estructura "
        Flog.writeline " Tipo: " & Testr1 & " - Codigo Externo: " & Estr1
        InsertaError 1, 92
         GoTo Fin
    End If
    
    
    'creo el parte de movilidad
    StrSql = "SELECT * FROM partemovilidad WHERE ternro=" & Ternro & " AND pmregfin=" & ConvFecha(CDate(FechaStr)) & " AND pmreghin = '" & HoraIni & "' AND pmregfin=" & ConvFecha(CDate(FechaStr)) & " AND pmreghin = '" & HoraSal & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
            
        StrSql = " INSERT INTO partemovilidad (" & _
                "ternro," & _
                "pmregfin," & _
                "pmreghin," & _
                "pmregfsa," & _
                "pmreghsa," & _
                "pmtenro1," & _
                "pmesnro1," & _
                "pmtenro2," & _
                "pmesnro2," & _
                "pmregnro1," & _
                "pmregnro2," & _
                "pmol" & _
                ") VALUES (" & _
                Ternro & "," & _
                ConvFecha(CDate(FechaStr)) & "," & _
                "'" & HoraIni & "'," & _
                ConvFecha(CDate(FechaStr)) & "," & _
                "'" & HoraSal & "'," & _
                Testr1 & "," & Estrnro1 & "," & _
                Testr2 & "," & Estrnro2 & "," & _
                regnro1 & "," & regnro2 & "," & _
                 ol & ")"
                
        objConn.Execute StrSql, , adExecuteNoRecords
        
        
    Else
        Flog.writeline " Ya existe el parte de mobilidad asosiado a la resgitracion "
        Flog.writeline " Fecha: " & FechaStr & " - Hora Inicio: " & HoraIni & " - Hora Salida: " & HoraSal
        InsertaError 1, 92
         GoTo Fin
    End If
      
        
Fin:
Exit Sub
ME_Local:

    HuboError = True
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    GoTo Fin
End Sub

Private Sub InsertaFormato34(strreg As String)
'-------------------------------------------------------------------
'Dimatz Rafael - 03/12/2013
'Formato:
'   Legajo, fecha, hora, nro de reloj, E/S
'Ejemplo:
'   000435  2013-12-03 05:38:46 1   I
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim separador As String
Dim Datos
Dim FechaHora
Dim HH, MM, SS
Dim Reg_Valida As Boolean

'----------------------------------------------------------
separador = Chr(9)
Datos = Split(strreg, separador)


    NroLegajo = Datos(0)
    Flog.writeline "Legajo:  " & NroLegajo
    
    FechaHora = Datos(1)
    regfecha = Mid(FechaHora, 1, 10)
    Flog.writeline "Fecha:  " & regfecha
    HH = Mid(FechaHora, 12, 2)
    MM = Mid(FechaHora, 15, 2)
        
    Hora = HH & MM
    Flog.writeline "Hora:  " & Hora

    nroreloj = Datos(2)
    
    nrorelojtxt = Datos(2)
    
    entradasalida = Datos(3)
    
    If entradasalida = "I" Then
        entradasalida = "E"
    Else
        entradasalida = "S"
    End If
        
    RegLeidos = RegLeidos + 1

    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
'    pos1 = pos2
'    pos2 = InStr(pos1 + 1, strreg, " ")
'    nroreloj = Mid(strreg, pos1, pos2 - pos1)
'    nrorelojtxt = Mid(strreg, pos1, pos2 - pos1)
'    Flog.writeline "Nro Reloj:  " & nrorelojtxt
    
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. No se encontro el Reloj: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    

'    pos1 = pos2
'    pos2 = InStr(pos1 + 1, strreg, " ")
'    entradasalida = IIf(UCase(Trim(Mid(strreg, pos1))) = "I", "E", "S")

       
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
      'StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & nroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
      'OpenRecordset StrSql, objRs
      'If Not objRs.EOF Then
      '   Ternro = objRs!Ternro
      'Else
         Flog.writeline "Error. No se encontro la terjeta para el Legajo: " & NroLegajo & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj: " & codReloj
         InsertaError 1, 33
         Exit Sub
      'End If
    End If
    
     'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
         
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
        
End Sub

Private Sub InsertaFormatoSpec3(strreg As String)
'-------------------------------------------------------------------
'Autor: Dimatz Rafael - 09/12/2013
'Formato: Los Campos son separados por ;
'   tipo de registro Marcaje
'   número de registro
'   código de empleado
'   tipo de incidencia
'   identificado de terminal
'   lector 1
'** fecha formato YYYYMMDD
'** hora hhmmss
'** origen del marcaje
'   carácter fijo
'   resultado del marcaje
'** número de tarjeta

'OBS
'El origen es el código de reloj.
'Usar solo los campos con **, el resto se descarta.
'---------------
'Ejemplo:
'M3;70;10517;0;2;;20131103;000304;8193;1;0;NULL;;;
'
'Ult. Modif:
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim regfecha As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Habilitada As String
Dim APYN As String
Dim regestado
Dim separador As String

Dim identificacion As String
Dim respuesta As String
Dim Resultado As String
Dim incidencia As String
Dim tarjeta As String
Dim Fecha As String
Dim Hora As String
Dim seg As String
Dim Anio As String
Dim Mes As String
Dim Dia As String
Dim Reg_Valida As Boolean

Dim Datos

RegLeidos = RegLeidos + 1

separador = ";"


Datos = Split(strreg, separador)

Fecha = Datos(6)
Hora = Left(Datos(7), 4)
nrorelojtxt = Datos(8)
tarjeta = Datos(11)
If tarjeta = "NULL" Then
    'tarjeta = datos(3)
    'FGZ - 10/12/2013 ----------
    tarjeta = Datos(2)
    'FGZ - 10/12/2013 ----------
End If

'Reloj
StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
OpenRecordset StrSql, objRs
If objRs.EOF Then
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Error. No se encontro el Reloj: " & nrorelojtxt
        Flog.writeline "SQL: " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
Else
    codReloj = objRs!relnro
    tipotarj = objRs!tptrnro
End If

'Reg estado se lo dejo fijo
regestado = "I"

'Aparentemente no distinguen entrada de salida
'Flog.writeline "Resultado:  " & Resultado
'If ((Resultado <> "E") Or (Resultado <> "S") Or (Resultado <> "")) Then
'    Resultado = ""
'End If
entradasalida = "" 'Resultado

'Acomodo la fecha
Flog.writeline "Fecha:  " & Fecha
Anio = Mid(Fecha, 1, 4)
Mes = Mid(Fecha, 5, 2)
Dia = Mid(Fecha, 7, 2)
regfecha = Anio + "/" + Mes + "/" + Dia

Flog.writeline "Hora:  " & Hora

'Validaciones
StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & tarjeta & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    Ternro = objRs!Ternro
Else
    Flog.writeline "    Warning. No se encontro la terjeta para el Legajo: " & tarjeta & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj: " & identificacion
    Flog.writeline "    Se validará con el nro de legajo. "
    'InsertaError 1, 33
    'Exit Sub
    
    'FGZ - 10/12/2013 --------------------------------------------------------
    StrSql = "SELECT ternro FROM empleado WHERE empleg = " & tarjeta
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Ternro = objRs!Ternro
    Else
        Flog.writeline "        Error. No se encontro el legajo asociado: " & tarjeta & ""
        InsertaError 1, 33
        Exit Sub
    End If
    'FGZ - 10/12/2013 --------------------------------------------------------
End If

'Carmen Quintero - 15/05/2015
Reg_Valida = True
StrSql = "SELECT relnro FROM gti_rel_estr "
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
    'Valido que el reloj sea de control de acceso para el empleado
    StrSql = "SELECT ternro FROM his_estructura H "
    StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
    StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
    StrSql = StrSql & " AND ( h.ternro = " & Ternro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Reg_Valida = False
        Flog.writeline "    El reloj No está habilitado para el empleado "
    End If
End If
'Fin Carmen Quintero - 15/05/2015

StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
OpenRecordset StrSql, objRs
If objRs.EOF Then

    Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    
    If Reg_Valida Then
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'" & regestado & "')"
    Else
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Call InsertarWF_Lecturas(Ternro, regfecha)
Else
    Flog.writeline " Registracion ya Existente"
    Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
    Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
    InsertaError 1, 92
End If

End Sub



Private Sub InsertaFormatoMarkovations(strreg As String)
'-------------------------------------------------------------------
'Autor: FGZ - 28/04/2014
'-------------------------------------------------------------------
'Formato:
'   Legajo <TAB> fechaHora <TAB> hora <TAB> fecharegistros <TAB> nro de reloj
'Ejemplo:
'1   24/04/2014 08:00:05 08:00   24/04/2014 09:00:01 1
'1   24/04/2014 08:01:20 08:01   24/04/2014 09:00:02 1
'12  24/04/2014 08:03:01 08:03   24/04/2014 09:00:03 1
'13  24/04/2014 08:05:10 08:05   24/04/2014 09:00:04 1
'14  24/04/2014 08:10:13 08:10   24/04/2014 09:00:05 1
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim separador As String
Dim Datos
Dim FechaHora
Dim FechaHoraRegistro
Dim HH, MM, SS
Dim Reg_Valida As Boolean

separador = Chr(9)
Datos = Split(strreg, separador)


    NroLegajo = Datos(0)
    Flog.writeline "Legajo:  " & NroLegajo
    
    FechaHora = Datos(1)
    regfecha = Mid(FechaHora, 1, 10)
    Flog.writeline "Fecha:  " & regfecha
    HH = Mid(FechaHora, 12, 2)
    MM = Mid(FechaHora, 15, 2)
    'Hora = HH & MM
    
    Hora = Datos(2)
    Flog.writeline "Hora:  " & Hora
    HH = Mid(Hora, 1, 2)
    MM = Mid(Hora, 4, 2)

    'No se usa
    FechaHoraRegistro = Datos(3)
    
    'Nro de reloj
    nroreloj = Datos(4)
    nrorelojtxt = Datos(4)
    
    'No viene informado por lo cual son todas entradas
    entradasalida = "E"
    'entradasalida = datos(3)
    'If entradasalida = "I" Then
    '    entradasalida = "E"
    'Else
    '    entradasalida = "S"
    'End If
        
    RegLeidos = RegLeidos + 1

    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline " Error Hora: " & Hora
        InsertaError 4, 38
        Exit Sub
    End If
    
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "Error. No se encontro el Reloj: " & nroreloj
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
            tipotarj = objRs!tptrnro
        End If
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
       
        
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
        'Que exista el legajo
        StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Flog.writeline "       ****** No se encontro el legajo --> " & NroLegajo
            InsertaError 1, 8
            Exit Sub
        Else
            Ternro = objRs!Ternro
        End If
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
        If Reg_Valida Then
    
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        
        objConn.Execute StrSql, , adExecuteNoRecords
        Call InsertarWF_Lecturas(Ternro, regfecha)
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
        
End Sub



Private Sub InsertaFormato35(strreg As String)

' ---------------------------------------------------------------------------------------------
' Descripcion:  - CAS 25421 - Lucaioli
' Autor      : Margiotta, Emanuel
' Fecha      : 17/06/2014
' Formato    : los campos estan separados por " " (espacio);

'CAMPO FORMATO
'Legajo             :9 digitos formato "000000000" / tarjeta
'Fecha              :dd/MM/aa
'Hora               :HH:MM
'Entrada/Salida     :I=entrada, O=salida
'Nro Reloj          :un digito
'
'Ejemplo
'    000009008 27/08/2013 13:46 O 6
'    000009590 27/08/2013 18:10 I 6
' ---------------------------------------------------------------------------------------------


Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Dia
Dim Mes
Dim Anio
Dim Reg_Valida As Boolean

    On Error GoTo ME_Local
    
    RegLeidos = RegLeidos + 1
    
    pos1 = 1
    pos2 = InStr(pos1, strreg, " ")
    NroLegajo = Mid(strreg, pos1, pos2 - pos1)
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Fecha = Mid(strreg, pos1, pos2 - pos1)
    
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    Hora = Trim(Mid(strreg, pos1, pos2 - pos1))
    
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    entradasalida = UCase(Trim(Mid(strreg, pos1, pos2 - pos1)))
    
    If entradasalida = "I" Then
        entradasalida = "E"
    ElseIf entradasalida = "O" Then
        entradasalida = "S"
    End If
    
    pos1 = pos2
    pos2 = InStr(pos1 + 1, strreg, " ")
    nroreloj = Trim(Mid(strreg, pos1))
    nrorelojtxt = Trim(Mid(strreg, pos1))
    
    '====================================================================
    ' Validar los parametros Levantados
    
    'Que exista el legajo
    StrSql = "SELECT * FROM empleado where empleg = " & NroLegajo
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & NroLegajo & "' AND (hstjfecdes <= " & ConvFecha(Fecha) & ") AND ( (" & ConvFecha(Fecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Ternro = objRs!Ternro
        Else
            Flog.writeline " Error Legajo: " & NroLegajo & " " & codReloj
            InsertaError 1, 33
            Exit Sub
        End If
        'Flog.writeline "       ****** No se encontro el legajo --> " & NroLegajo
        'InsertaError 1, 8
        'GoTo Fin
    Else
        Ternro = objRs!Ternro
    End If
    
    'Que la fecha sea válida
    Dia = Mid(Fecha, 1, 2)
    Mes = Mid(Fecha, 4, 2)
    Anio = Mid(Fecha, 7, 4)
    If Not IsNumeric(Dia) Or Not IsNumeric(Mes) Or Not IsNumeric(Anio) Then
        Flog.writeline "       ****** Fecha no válida --> " & Fecha
        InsertaError 2, 4
        GoTo Fin
    Else
        If CDate(Dia & "/" & Mes & "/" & Anio) Then
            regfecha = CDate(Dia & "/" & Mes & "/" & Anio)
        Else
            Flog.writeline "       ****** Fecha no válida --> " & Fecha
            InsertaError 2, 4
            GoTo Fin
        End If
    End If
    
    'Que la hora sea válida
    If Not objFechasHoras.ValidarHora(Hora) Then
        Flog.writeline "       ****** Hora no válida --> " & Hora
        InsertaError 3, 38
        GoTo Fin
    End If
    
    'Busco el Reloj
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nroreloj & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "       ****** No se encontro el Reloj. SQL --> " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
    
    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015
    
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        
    Else
        Flog.writeline " Registracion ya Existente "
        Flog.writeline " Hora: " & Hora & " - Fecha: " & regfecha
        InsertaError 1, 92
    End If
        
Fin:
Exit Sub
ME_Local:
    HuboError = True
    Flog.writeline "Error: " & Err.Number
    Flog.writeline "Decripcion: " & Err.Description
    GoTo Fin
End Sub

Private Sub InsertaFormato158(strreg As String)
'-------------------------------------------------------------------
'Autor: LED - 05/06/2015 - CAS-30945 - ASM - Nuevo formato de reloj
'Formato: Los Campos son separados por ;
'   tipo de registro Marcaje
'   número de registro
'** Legajo de empleado
'   sin uso
'   sin uso
'   sin uso
'** fecha formato YYYYMMDD
'** hora hhmmss
'** Reloj (origen del marcaje)
'   carácter fijo
'   Sin uso
'   Null
'   Sin uso
'   Sin uso

'OBS
'El origen es el código de reloj.
'Usar solo los campos con **, el resto se descarta.
'---------------
'Ejemplo:
'M3;46;-1;0;1;;20130520;135417;4097;1;2;0000000209000503;;;
'
'Ult. Modif:
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim regfecha As String
Dim entradasalida As String
Dim nrorelojtxt As String

Dim codReloj As Integer
Dim tipotarj As Integer

Dim regestado
Dim separador As String
Dim Fecha As String
Dim Hora As String
Dim Anio As String
Dim Mes As String
Dim Dia As String
Dim Reg_Valida As Boolean
Dim Datos

'Dim identificacion As String

'Dim tarjeta As String


RegLeidos = RegLeidos + 1

separador = ";"
Datos = Split(strreg, separador)

Fecha = Datos(6)
Hora = Left(Datos(7), 4)
nrorelojtxt = Datos(8)
'tarjeta = datos(11)
NroLegajo = Datos(2)

'Reloj
StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
OpenRecordset StrSql, objRs
If objRs.EOF Then
    StrSql = "SELECT relnro, tptrnro FROM gti_reloj WHERE relcodext = '" & nrorelojtxt & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Error. No se encontro el Reloj: " & nrorelojtxt
        Flog.writeline "SQL: " & StrSql
        InsertaError 4, 32
        Exit Sub
    Else
        codReloj = objRs!relnro
        tipotarj = objRs!tptrnro
    End If
Else
    codReloj = objRs!relnro
    tipotarj = objRs!tptrnro
End If

'Reg estado se lo dejo fijo
regestado = "I"

'Aparentemente no distinguen entrada de salida
'Flog.writeline "Resultado:  " & Resultado
'If ((Resultado <> "E") Or (Resultado <> "S") Or (Resultado <> "")) Then
'    Resultado = ""
'End If
entradasalida = "" 'Resultado

'Acomodo la fecha
Flog.writeline "Fecha:  " & Fecha
Anio = Mid(Fecha, 1, 4)
Mes = Mid(Fecha, 5, 2)
Dia = Mid(Fecha, 7, 2)
regfecha = Anio + "/" + Mes + "/" + Dia

Flog.writeline "Hora:  " & Hora


'Validaciones
'StrSql = "SELECT ternro FROM gti_histarjeta WHERE tptrnro = " & tipotarj & " AND hstjnrotar = '" & tarjeta & "' AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
'OpenRecordset StrSql, objRs
'If Not objRs.EOF Then
'    Ternro = objRs!Ternro
'Else
'    Flog.writeline "Error. No se encontro la terjeta para el Legajo: " & tarjeta & ", tipo de tarjeta: " & tipotarj & " y codigo de reloj: " & identificacion
'    InsertaError 1, 33
'    Exit Sub
'End If

StrSql = "SELECT ternro FROM empleado WHERE empleg = " & NroLegajo
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    Ternro = objRs!Ternro
Else
    Flog.writeline "Error. No se encontro empleado para el Legajo: " & NroLegajo
    InsertaError 3, 8
    Exit Sub
End If

Reg_Valida = True
StrSql = "SELECT relnro FROM gti_rel_estr "
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
    'Valido que el reloj sea de control de acceso para el empleado
    StrSql = "SELECT ternro FROM his_estructura H "
    StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
    StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
    StrSql = StrSql & " AND ( h.ternro = " & Ternro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Reg_Valida = False
        Flog.writeline "    El reloj No está habilitado para el empleado "
    End If
End If


StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
OpenRecordset StrSql, objRs
If objRs.EOF Then

    Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora

    If Reg_Valida Then
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'" & regestado & "')"
    Else
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
            Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
    End If
    
    objConn.Execute StrSql, , adExecuteNoRecords
    Call InsertarWF_Lecturas(Ternro, regfecha)
Else
    Flog.writeline " Registracion ya Existente"
    Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
    Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
    InsertaError 1, 92
End If

End Sub
Private Sub InsertaFormato157(strreg As String)
'---------------------------------------------------------------------------
'Autor      : Gonzalez Nicolás - CAS-33995 - SOLAR - Nuevo formato de reloj
'Fecha      : 12/01/2016
'Modificado :
'---------------------------------------------------------------------------

'Detalle
'-----------
'Formato del archivo: txt
'Separador de campos: 1 espacio o más
'Linea 1 -> Encabezado
'Linea 2 -> Linea en blanco
'Linea 3 -> Comienzan los datos a inserir


'Columna 1: Fecha
'Formato: DD/MM/AAAA

'Columna 2: Cédula de Identidad (Documento del empleado)

'Columna 3: Hora de primera registración del día
'Formato: HH: MM
'E ->

'Columna 4: Hora de segunda registración del día
'Formato: HH: MM
'(si no hay dato, lo trae en blanco)
'S ->


'Ejemplo:
'       Dia CI Nro. Marc-Ent Marc-Sal

'01/01/2016 1234567    08:10    18:00
'02/01/2016 1234567    08:05
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim regfecha As String
Dim entradasalida As String
Dim nrorelojtxt As String
Dim codReloj As Integer
Dim regestado
Dim separador As String
Dim Fecha As String
Dim Hora As String
Dim Anio As String
Dim Mes As String
Dim Dia As String
Dim Nrodoc As String
Dim Datos
Dim datos2

Dim HEntrada As String
Dim HSalida As String
Dim HayEntrada As Boolean
Dim HaySalida As Boolean
Dim ax As Long
HayEntrada = True
HaySalida = True

'---------------------------------------------------------------------------
'CONTROLO QUE EXISTE UN RELJO COMO DEFAULT
StrSql = "SELECT relnro,reldabr FROM GTI_RELOJ WHERE reldefault = -1"
OpenRecordset StrSql, objRs
If objRs.EOF Then
    Flog.writeline "Error. No se encontro el Reloj configurado como DEFAULT "
    InsertaError 4, 32
    Exit Sub
Else
    codReloj = objRs!relnro
    Flog.writeline "Reloj N°:" & codReloj & " (" & objRs!reldabr & ")"
End If
'---------------------------------------------------------------------------

RegLeidos = RegLeidos + 1
separador = " "
datos2 = ""
Datos = Split(strreg, separador)
For ax = 0 To UBound(Datos)
    If Datos(ax) <> "" Then
        datos2 = datos2 & ";" & Datos(ax)
    End If
Next
datos2 = Split(datos2, ";")


Fecha = datos2(1)
Nrodoc = datos2(2)
HEntrada = ""
If UBound(datos2) > 2 Then
    HEntrada = datos2(3)
End If
HSalida = ""
If UBound(datos2) > 3 Then
    HSalida = datos2(4)
End If

'---------------------------------------------------------------------------
'CONTROL DE FECHA
If EsNulo(Fecha) = True Then
    Flog.writeline "Error. La fecha informada debe ser DD/MM/AAAA."
    InsertaError 4, 32
    Exit Sub
End If

'---------------------------------------------------------------------------
'CONTROL DE DOCUMENTO
'---------------------------------------------------------------------------
If EsNulo(Nrodoc) = False Then
    'BUSCO EMPLEADO POR N° DE DOC, DE NO EXISTIR NO CONTINUA
    StrSql = "SELECT empleado.ternro ,empleado.empleg  "
    StrSql = StrSql & " FROM empleado"
    StrSql = StrSql & " INNER JOIN ter_doc ON ter_doc.ternro = empleado.ternro"
    StrSql = StrSql & " WHERE ter_doc.Nrodoc = '" & Nrodoc & "'"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "Error. No se encontro el Empleado para el N° de documento: " & Nrodoc
        InsertaError 4, 32
        Exit Sub
    Else
        Ternro = objRs!Ternro
        NroLegajo = objRs!EmpLeg
    End If
Else
    Flog.writeline "Error. No se encontro el Empleado para el N° de documento: " & Nrodoc
    InsertaError 4, 32
    Exit Sub
End If
'---------------------------------------------------------------------------


'---------------------------------------------------------------------------
'CONTROLO HS de E y de S
'---------------------------------------------------------------------------
If EsNulo(HEntrada) = True And EsNulo(HSalida) = True Then
    HayEntrada = False
    HaySalida = False
    Flog.writeline "Error. Al registro al menos debe contener hora de Entrada o de Salida."
    InsertaError 4, 32
    Exit Sub
Else
    'Controlo> E
    If (EsNulo(HEntrada) = False) And (Len(HEntrada) < 5 Or Len(HEntrada) > 5) Then
        HayEntrada = False
        Flog.writeline "Error. Hora de entrada incorrecta :" & HEntrada
        Flog.writeline "       Formato: HH:MM"
        InsertaError 4, 32
        Exit Sub
    Else
        If (EsNulo(HEntrada) = True) Then
            HayEntrada = False
        End If
    End If

    'Controlo> S
    If (EsNulo(HSalida) = False) And (Len(HSalida) < 5 Or Len(HSalida) > 5) Then
        HaySalida = False
        Flog.writeline "Error. Hora de Salida incorrecta :" & HSalida
        Flog.writeline "       Formato: HH:MM"
        InsertaError 4, 32
        Exit Sub
    Else
        If (EsNulo(HSalida) = True) Then
            HaySalida = False
        End If
    End If
End If


'Reg estado se lo dejo fijo en I
regestado = "I"
regfecha = Fecha
'---------------------------------------------------------------------------
'CONTROLA SI HAY ASOCIADAS ESTRUTURAS A UN RELOJ
'---------------------------------------------------------------------------
StrSql = "SELECT relnro FROM gti_rel_estr "
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
    'Valido que el reloj sea de control de acceso para el empleado
    StrSql = "SELECT ternro FROM his_estructura H "
    StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
    StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
    StrSql = StrSql & " AND ( h.ternro = " & Ternro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        regestado = "X" 'ACTUALIZO regestado
        Flog.writeline "    El reloj No está habilitado para el empleado "
    End If
End If

If HayEntrada = True Then
    entradasalida = "E"
    Hora = Replace(HEntrada, ":", "")
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "                       INSERTO REGISTRACION DE ENTRADA - Legajo: " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'" & regestado & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        Call InsertarWF_Lecturas(Ternro, regfecha)
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
End If

If HaySalida = True Then
    entradasalida = "S"
    Hora = Replace(HSalida, ":", "")
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        Flog.writeline "                       INSERTO REGISTRACION DE SALIDA - Legajo: " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
        StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
        Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'" & regestado & "')"
        objConn.Execute StrSql, , adExecuteNoRecords
        Call InsertarWF_Lecturas(Ternro, regfecha)
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If

End If
End Sub

Private Sub InsertaFormato155(strreg As String)
'---------------------------------------------------------------------------
'Autor      : Gonzalez Nicolás - CAS-36794 - TNPlatex - Nuevo formato de reloj
'Fecha      : 20/04/2016
'Modificado : 25/04/2016 - Gonzalez Nicolás - CAS-36794 - TNPlatex - Nuevo formato de reloj - Se busca el reloj solamente por el cód. externo (nodo)
'---------------------------------------------------------------------------

'Detalle
'-----------
'Formato del archivo: txt/rlb
'Separador de campos: Coma

'Nodo    =4             -> Este es el ID en la base de datos del nodo(Equipo) del que provienen los registros. (Código externo de reloj)
'Nombre  =RelojDos      -> Este es el nombre que se le da al Equipo desde la base de datos del Nodo. (nombre del reloj asociado al código externo)
'Lectora =0             -> Esto corresponde al número de la lectora que posee el equipo, donde se produjo la marcación. Dos posibles Opciones 0 y 1.
'                          NO SE UTILIZA EN RHPRO
'Sentido = Entrada      -> Entrada Solo entrada y salida
'Tarjeta =1340677       -> Es el Número de Tarjeta de cada empleado
'Fecha   =23/03/2016    -> DD/MM/AAAA
'Hora    =05:58:16      -> HH:MM:SS
'Opción  =0             -> Corresponde al código ingresado por un teclado numérico si el equipo lo tuviera. En el caso de nuestros equipos siempre es cero
'                          NO SE UTILIZA EN RHPRO
'-------------------------------------------------------------------
Dim NroLegajo As String
Dim Ternro As Long
Dim regfecha As String
Dim entradasalida As String
'Dim nrorelojtxt As String
Dim codReloj As Integer
Dim regestado
Dim separador As String
Dim Fecha As String
Dim Hora As String
Dim Datos


'NEW
Dim relcodext
Dim reldabr
Dim hstjnrotar
Dim TipoReg
Dim Aux As String
'------------
Dim ax As Long



RegLeidos = RegLeidos + 1
separador = ","
Datos = Split(strreg, separador)

'********************************
'Campo => Nodo = 4
'********************************
If InStr(Datos(0), "=") > 0 Then
    ax = InStr(Datos(0), "=")
    relcodext = Trim(Mid(Datos(0), ax + 1, Len(Datos(0))))
End If

'********************************
'Campo => Nombre = RelojDos
'********************************
If InStr(Datos(1), "=") > 0 Then
    ax = InStr(Datos(1), "=")
    reldabr = Trim(Mid(Datos(1), ax + 1, Len(Datos(1))))
End If

'**************************************************
'Campo => Lectora = 1 | NO SE UTILIZA EN RHPRO
'**************************************************

'********************************
'Campo => Sentido = Salida
'********************************
If InStr(Datos(3), "=") > 0 Then
    'Entrada | Salida
    ax = InStr(Datos(3), "=")
    entradasalida = Trim(Mid(Datos(3), ax + 1, Len(Datos(3))))
    If UCase(entradasalida) <> "ENTRADA" And UCase(entradasalida) <> "SALIDA" Then
        Flog.writeline "Error. Solo se admiten registros de Entrada o Salida"
        Exit Sub
    Else
        entradasalida = Left(entradasalida, 1)
        If entradasalida = "E" Then
            TipoReg = "00"
        Else
            TipoReg = "01"
        End If
    End If
End If

'********************************
'Campo => tarjeta = 15639947
'********************************
If InStr(Datos(4), "=") > 0 Then
    ax = InStr(Datos(4), "=")
    hstjnrotar = Trim(Mid(Datos(4), ax + 1, Len(Datos(4))))
    If Len(hstjnrotar) > 20 Then
        Flog.writeline "Error. El N° de tarjeta " & hstjnrotar & " excede los 20 caracteres."
        InsertaError 4, 32
        Exit Sub
    End If
End If


'********************************
'Campo => Fecha = 23 / 03 / 2016
'********************************
If InStr(Datos(5), "=") > 0 Then
    ax = InStr(Datos(5), "=")
    Fecha = Trim(Mid(Datos(5), ax + 1, Len(Datos(5))))
    If EsNulo(Fecha) = True Then
        Flog.writeline "Error. La fecha informada debe ser DD/MM/AAAA."
        InsertaError 4, 32
        Exit Sub
    Else
        If Not IsDate(Fecha) Then
            Flog.writeline "Error. La fecha informada debe ser DD/MM/AAAA."
            InsertaError 4, 32
            Exit Sub
        End If
    End If
End If

'********************************
'Campo => Hora=06:04:20
'********************************
If InStr(Datos(6), "=") > 0 Then
    ax = InStr(Datos(6), "=")
    Aux = Trim(Mid(Datos(6), ax + 1, Len(Datos(6))))
    If Len(Aux) < 5 Or Len(Aux) > 8 Then
        Flog.writeline "Error. El formato de Hora es incorrecto. Debe ser 00:00:00 o 00:00"
        InsertaError 4, 32
        Exit Sub
    ElseIf Len(Aux) = 8 Then
        Hora = Left(Aux, 5)
    End If
    'Elimino :
    Hora = Replace(Hora, ":", "")
End If

'****************************************************
'Campo => Opcion = 0 | NO SE UTILIZA EN RHPRO
'****************************************************






'------------------------------------------------------------------------------------------------
'CONTROLO QUE EXISTE UN RELOJ CON EL CÓD. EXTERNO IDENTIFICADO
'------------------------------------------------------------------------------------------------
'StrSql = "SELECT relnro,reldabr,relcodext FROM GTI_RELOJ WHERE relcodext = '" & relcodext & "' AND reldabr ='" & reldabr & "'"
StrSql = "SELECT relnro,reldabr,relcodext FROM GTI_RELOJ WHERE relcodext = '" & relcodext & "'"
OpenRecordset StrSql, objRs
If objRs.EOF Then
    Flog.writeline "Error. No se encontro el Reloj " & reldabr & " con el código externo: " & relcodext
    InsertaError 4, 32
    Exit Sub
Else
    codReloj = objRs!relnro
    Flog.writeline "Reloj N°:" & codReloj & " (" & objRs!reldabr & ")"
End If
'------------------------------------------------------------------------------------------------

'Reg estado se lo dejo fijo en I
regestado = "I"
regfecha = Fecha

'------------------------------------------------------------------------------------------------
'CONTROLO QUE EXISTE EL NÚMERO DE TARJETA Y PERTENEZCA A UN EMPLEADO
'------------------------------------------------------------------------------------------------
StrSql = "SELECT gti_histarjeta.ternro,empleado.empleg "
StrSql = StrSql & " FROM gti_histarjeta "
StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = gti_histarjeta.ternro"
StrSql = StrSql & " WHERE hstjnrotar = '" & hstjnrotar & "'"
StrSql = StrSql & " AND (hstjfecdes <= " & ConvFecha(regfecha) & ") "
StrSql = StrSql & " AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
   Ternro = objRs!Ternro
   NroLegajo = objRs!EmpLeg
Else
    Flog.writeline "Error. La Tarjeta : " & hstjnrotar & " no se ha encontrado."
    InsertaError 4, 32
    Exit Sub
End If



'---------------------------------------------------------------------------
'CONTROLA SI HAY ASOCIADAS ESTRUTURAS A UN RELOJ
'---------------------------------------------------------------------------
StrSql = "SELECT relnro FROM gti_rel_estr "
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
    'Valido que el reloj sea de control de acceso para el empleado
    StrSql = "SELECT ternro FROM his_estructura H "
    StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
    StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
    StrSql = StrSql & " AND ( h.ternro = " & Ternro
    StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
        regestado = "X" 'ACTUALIZO regestado
        Flog.writeline "    El reloj No está habilitado para el empleado "
    End If
End If

'------------------------------------------------------------------------------------------------
'-- INSERTO REGISTROS
'------------------------------------------------------------------------------------------------
StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
OpenRecordset StrSql, objRs
If objRs.EOF Then
    If entradasalida = "E" Then
        Flog.writeline " INSERTO REGISTRACION DE ENTRADA - Legajo: " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    Else
        Flog.writeline " INSERTO REGISTRACION DE SALIDA - Legajo: " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    End If
    
    StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado,tiporeg) VALUES (" & _
    Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'" & regestado & "','" & TipoReg & "')"
    objConn.Execute StrSql, , adExecuteNoRecords
    Call InsertarWF_Lecturas(Ternro, regfecha)
Else
    Flog.writeline " Registracion ya Existente"
    Flog.writeline " Error Legajo: " & NroLegajo
    Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
    InsertaError 1, 92
End If


End Sub


Private Sub InsertaFormatoIQFarma(strreg As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Formato para IQFARMA
' Autor      : Dimatz Rafael
' Fecha      : 12/02/2016
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

'Formato:
'   Legajo fecha hora
'Ejemplo:
'   109727711 09/02/2915 14:08
'Separador: espacio en blanco
'-------------------------------------------------------------------

Dim NroLegajo As String
Dim Ternro As Long
Dim Fecha As Date
Dim regfecha As String
Dim Hora As String
Dim entradasalida As String
Dim nroreloj As Long
Dim nrorelojtxt As String
Dim pos1 As Byte
Dim pos2 As Byte
Dim codReloj As Integer
Dim tipotarj As Integer
Dim Reg_Valida As Boolean
Dim NroTarjeta As String
Dim Datos


' No Distingue E/S por eso se deja fijo siempre E. Luego usar la politica correspondiente que no toma en cuenta las E/S

    entradasalida = "E"
    
    RegLeidos = RegLeidos + 1

    separador = " "
    Datos = Split(strreg, separador)

    NroTarjeta = Datos(0)
    regfecha = Datos(1)
    Hora = Datos(2)
    Hora = Replace(Hora, ":", "")
    
    Flog.writeline "Busco el reloj"
    StrSql = "SELECT relnro FROM gti_reloj WHERE relhabil= -1"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
            Flog.writeline "Error. No hay Relojes Habilitados "
            Flog.writeline "SQL: " & StrSql
            InsertaError 4, 32
            Exit Sub
        Else
            codReloj = objRs!relnro
        End If
       
    Flog.writeline "Busco el nro de tarjeta "
    'FGZ - 10/03/2016
    StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & NroTarjeta & "' "
    'StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & Format(NroTarjeta, "0000000000") & "' "
    StrSql = StrSql & " AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       Ternro = objRs!Ternro
    Else
        'StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & NroTarjeta & "' "
        StrSql = "SELECT ternro FROM gti_histarjeta WHERE hstjnrotar = '" & Format(NroTarjeta, "0000000000") & "' "
        StrSql = StrSql & " AND (hstjfecdes <= " & ConvFecha(regfecha) & ") AND ( (" & ConvFecha(regfecha) & " <= hstjfechas) OR ( hstjfechas is null ))"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
           Ternro = objRs!Ternro
        Else
             Flog.writeline "Error. Trajeta: " & NroTarjeta & " no encontrada"
             Flog.writeline "SQL: " & StrSql
             InsertaError 1, 33
             Exit Sub
        End If
    End If

    'Carmen Quintero - 15/05/2015
    Reg_Valida = True
    StrSql = "SELECT relnro FROM gti_rel_estr "
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        'significa que los relojes tienen alcance por estructura ==> valido que el empleado tenga alcance para el reloj
        'Valido que el reloj sea de control de acceso para el empleado
        StrSql = "SELECT ternro FROM his_estructura H "
        StrSql = StrSql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = h.estrnro "
        StrSql = StrSql & " WHERE gti_rel_estr.relnro = " & codReloj
        StrSql = StrSql & " AND ( h.ternro = " & Ternro
        StrSql = StrSql & " AND htetdesde <= " & ConvFecha(regfecha) & " AND (htethasta is null or htethasta >= " & ConvFecha(regfecha) & "))"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            Reg_Valida = False
            Flog.writeline "    El reloj No está habilitado para el empleado "
        End If
    End If
    'Fin Carmen Quintero - 15/05/2015

'Siempre Igual
    StrSql = "SELECT * FROM gti_registracion WHERE regfecha = " & ConvFecha(regfecha) & " AND reghora = '" & Hora & "' AND ternro = " & Ternro & " AND regentsal = '" & entradasalida & "' AND relnro = " & codReloj
    OpenRecordset StrSql, objRs
    If objRs.EOF Then
    
        Flog.writeline "                       INSERTO REGISTRACION - " & NroLegajo & "  ;  '" & regfecha & "'    ;    " & Hora
    
        If Reg_Valida Then
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'I')"
        Else
            StrSql = " INSERT INTO gti_registracion(ternro,crpnnro,regfecha,reghora,regentsal,relnro,regestado) VALUES (" & _
                Ternro & "," & crpNro & "," & ConvFecha(regfecha) & ",'" & Hora & "','" & entradasalida & "'," & codReloj & ",'X')"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.writeline "SQL: -->" & StrSql
        
        Call InsertarWF_Lecturas(Ternro, regfecha)
        Flog.writeline "Inserto en temporal WF_Lecturas"
    Else
        Flog.writeline " Registracion ya Existente"
        Flog.writeline " Error Legajo: " & NroLegajo & " " & tipotarj & " " & codReloj
        Flog.writeline " Hora: " & Hora & " - Fecha: '" & regfecha & "'"
        InsertaError 1, 92
    End If
    Flog.writeline "Linea Procesada"
End Sub

