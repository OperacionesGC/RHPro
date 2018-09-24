Attribute VB_Name = "MdlExportacion"
Option Explicit

'Global Const Version = "1.01"
'Global Const FechaModificacion = "07/10/2005"
'Global Const UltimaModificacion = " " 'Nro de version

'Global Const Version = "1.02"
'Global Const FechaModificacion = "23/03/2006"
'Global Const UltimaModificacion = " " 'Se agrego CODMODASI, FECHAPROCVOL, FECHALIQH

'Global Const Version = "1.03"
'Global Const FechaModificacion = "30/03/2006"
'Global Const UltimaModificacion = " " 'Se corrigio Importe_Format. Ej: 12.03 lo convertía en 12.30

'Global Const Version = "1.04"
'Global Const FechaModificacion = "03/06/2006"
'Global Const UltimaModificacion = " " 'Se agrego el modelo de salida de exportacion de asientos.
                                      ' Viene como parámetro del asp

'Global Const Version = "1.05"
'Global Const FechaModificacion = "25/07/2006"
'Global Const UltimaModificacion = " " 'En el item SAPLINE hay que invertir el signo. Si en la salida esta poniendo
                                    'el signo positivo debe ir negativo y si pone negativo debe ir positivo

'Global Const Version = "1.06"
'Global Const FechaModificacion = "15/08/2006"
'Global Const UltimaModificacion = " " 'Se agrego el item PROCESO en el encabezado
                                      'Se agrego el item DOCINCTA, IMPORTED, IMPORTEH
                                      'Se modifico de manera que si no se informa el Modelo de Salida de Asiento, no de error

'Global Const Version = "1.07"
'Global Const FechaModificacion = "25/10/2006" ' Martin Ferraro
'Global Const UltimaModificacion = " " 'Se agrego el item IMPORTEDUSD, IMPORTEHUSD, COTIZA, INTERFACE, SECFECHA

'Global Const Version = "1.08"
'Global Const FechaModificacion = "07/12/2006" ' Martin Ferraro
'Global Const UltimaModificacion = " " 'Se Corrigieron errores en IMPORTED, IMPORTEH,IMPORTEDUSD, IMPORTEHUSD,

'Global Const Version = "1.09"
'Global Const FechaModificacion = "17/01/2007" ' Leticia Amadio
'Global Const UltimaModificacion = " " 'Se agrego un item, con programa Profit en Detalle - Custom: Roche

'Global Const Version = "1.10"
'Global Const FechaModificacion = "06/02/2007" ' FGZ
'Global Const UltimaModificacion = " "   'Se eliminó el item Profit en Detalle - Custom: Roche
'                                       'Se agregó un item item CC_Profit en Detalle que agrupa CCosto y Profit- Custom: Roche
'                                       'El item profit debe estar configurado en el mismo lugar que el centro de costo en el asiento. es decir
'                                       Antes
'                                           el modelo de asiento tenia apertura por centro de costo.
'                                           Habia lineas del modelo con apertura  y cuentas que no tenia apertura
'                                       Ahora
'                                           el modelo de asiento tiene apertura por centro de costo y Profit Center
'                                           las lineas del modelo con apertura por Centro de costo quedan igual y
'                                           las lineas que no tenia apertura se les pone apertura por Profit
'                                       Ejemplo:
'                                           Antes
'                                               XXXXXXXX-E1E1E1E1E1E1E1E1E1E1   ---> CON APERTURA
'                                               XXXXXXXX                        ---> SIN APERTURA
'                                           Ahora
'                                               XXXXXXXX-E1E1E1E1E1E1E1E1E1E1   ---> CON APERTURA POR CC
'                                               XXXXXXXX-E2E2E2E2E2E2E2E2E2E2   ---> SIN APERTURA POR PROFIT


'Global Const Version = "1.11"
'Global Const FechaModificacion = "21/02/2007" ' Martin Ferraro
'Global Const UltimaModificacion = " "   'Se agrego funcion SICUENTA

'Global Const Version = "1.12"
'Global Const FechaModificacion = "24/04/2007" ' Raul Chinestra
'Global Const UltimaModificacion = " "   'En la función BusCotiza la Parte decimal del nro estaba saliendo mal, por ejemplo 3.08 salia 3.80.

'Global Const Version = "1.13"
'Global Const FechaModificacion = "06/06/2007"   ' Martin Ferraro
'Global Const UltimaModificacion = " "           ' Se agrego Item CTAreplace para Praxiar Argentina

'Global Const Version = "1.14"
'Global Const FechaModificacion = "15/06/2007"   ' Martin Ferraro
'Global Const UltimaModificacion = " "           'se modifico la definicion del Item CTAreplace

'Global Const Version = "1.15"
'Global Const FechaModificacion = "11/10/2007"   ' Fernando Favre
'Global Const UltimaModificacion = " "           ' se agrego el item FECHA al encabezado y el item SICTA para CCU (Cia Cervecera Industrial)

'Global Const Version = "1.16"
'Global Const FechaModificacion = "29/10/2007"   ' Martin Ferraro
'Global Const UltimaModificacion = " "           ' se agrego un nuevo formato de nombre de archivo de exp.

'Global Const Version = "1.17"
'Global Const FechaModificacion = "04/03/2008"   ' Gustavo Ring
'Global Const UltimaModificacion = " "           ' Se corrigio para Arlei, imprimia un separador al comienzo de cada linea

'Global Const Version = "1.18"
'Global Const FechaModificacion = "18/04/2008"   ' FGZ
'Global Const UltimaModificacion = " "           ' se agrego un nuevo item CUENTAN

'Global Const Version = "1.19"
'Global Const FechaModificacion = "10/06/2008"    ' Gustavo Ring
'Global Const UltimaModificacion = " "            ' se agrego un nuevo item VOLCOD

'Global Const Version = "1.20"
'Global Const FechaModificacion = "26/06/2008"    ' Fernando Favre
'Global Const UltimaModificacion = " "            ' Se modificaron los ITEM FECHA tanto del encabezado como del detalle para permitir 10 caracteres. Por ej: DD/MM/YYYY

'Global Const Version = "1.21"
'Global Const FechaModificacion = "27/07/2008"    ' Martin Ferraro
'Global Const UltimaModificacion = " "            ' Se creo el item AsientoZ

'Global Const Version = "1.22"
'Global Const FechaModificacion = "07/01/2009"    ' Diego Nuñez
'Global Const UltimaModificacion = " "            ' Se creo el item que imprime el número de modelo del asiento MODELO_NRO

'Global Const Version = "1.23"
'Global Const FechaModificacion = "08/01/2009"    ' Diego Nuñez
'Global Const UltimaModificacion = " "            ' Se modificó la subrutina Importe para que tuviese en cuenta de forma correcta el último valor del asiento.

'Global Const Version = "1.24"
'Global Const FechaModificacion = "12/01/2009"    ' Diego Nuñez
'Global Const UltimaModificacion = " "            ' Se generó un nuevo nombre de salida del archivo de exportación, el cual se selecciona desde el asp de exportación.
                                                 ' El nombre del archivo de salida agregado tiene como formato: "nro_proceso" + "nro_modelo" + "Fecha actual(ddmmaaaa)"

'Global Const Version = "1.25"
'Global Const FechaModificacion = "23/01/2009"    ' Diego Nuñez
'Global Const UltimaModificacion = " "            ' Se creo un item que imprime los montos con decimales y signo + o - dependiendo si es Debe o haber respectivamente.
                                                 ' Además completa con espacios a la izquierda.

'Global Const Version = "1.26"
'Global Const FechaModificacion = "13/02/2009"    ' Diego Nuñez
'Global Const UltimaModificacion = " "            ' Se generaron nuevos nombres de salida del archivo de exportación, el cual se selecciona desde el asp de exportación.
                                                 ' El nombre del archivo de salida puede tener como formato:
                                                 '"DescModelo" + "NumProc" + "Fecha actual(ddmmaaaa)"
                                                 '"DescModelo" + "DescProc" + "Fecha actual(ddmmaaaa)"
                                                 '"DescModelo" + "NumProc" + "DescProc" + "Fecha actual(ddmmaaaa)"
                                                 '"DescModelo" + "DescProc" + "NumProc"  + "Fecha actual(ddmmaaaa)"
                                                 '"Modelo" + "Proceso" + "Fecha actual(ddmmaaaa)"
                                                 
'Global Const Version = 1.27
'Global Const FechaModificacion = "19/08/2009"   'Encriptacion de string connection
'Global Const UltimaModificacion = "Manuel Lopez"
'Global Const UltimaModificacion1 = "Encriptacion de string connection"

'Global Const Version = "1.28"
'Global Const FechaModificacion = "26/08/2009"    ' Diego Nuñez
'Global Const UltimaModificacion = " "            ' Se generó un nuevo nombre de salida del archivo de exportación, el cual se selecciona desde el asp de exportación.
                                                 ' El nombre del archivo de salida agregado tiene como formato: "nnnnmmaa" (nnnn=codigo de asiento, mm=mes de liquidación, aa=Año de liq)
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.29"
'Global Const FechaModificacion = "08/10/2009"    ' Manuel Lopez
'Global Const UltimaModificacion = " "            ' Se agrego el tipo de archivo 13 con el formato de la modificacion 1.28
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.30"
'Global Const FechaModificacion = "29/09/2010"    ' Martin
'Global Const UltimaModificacion = " "            ' Cuando levanta el separador decimales y de campos y se config en vacio guardaba
'Global Const UltimaModificacion1 = " "           ' nulo y fallaba en oracle

'Global Const Version = "1.31"
'Global Const FechaModificacion = "05/07/2011"       ' Zamarbide Juan Alberto - Caso CAS-13437 - Error en exportación de Asiento - Resolución del Bug realizada
'Global Const UltimaModificacion = " "               ' Se comentó la comprobación previa a la concatenación de la cadena a escribir en el archivo de Exportación. Dicha comprobación se encontraba tanto en el encabezado, como en el cuerpo y pie,
                                                    ' en cada caso fué comentada. El motivo por el cual se realizaba dicha comprobación, no se pudo establecer, pero se identificó que gracias a dicha comprobación no concatenaba el
                                                    ' separadorCampos (";") cuando existía una parte de la cadena que comenzara con "RR" (comprobación = If Mid(cadena, 1, 2) <> "RR" Or primero Then ).
                                                    ' Se Agregó en el Else de la comprobación anterior la línea "Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)" y se comentó la anterior
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.32"
'Global Const FechaModificacion = "06/07/2011"       ' Zamarbide Juan Alberto - Caso CAS-13081 - Vision - Amarilla Gas - Exportación Asiento Contable
'Global Const UltimaModificacion = " "               ' Se Agregó el Case "ASIENTO" en el armado del encabezado para que muestre el Nro de Asiento en dicho encabezado.
'Global Const UltimaModificacion1 = " "              ' Se agregó el comentario "Agregado en vers. 1.32" a las lineas agregadas en la versión.

'Global Const Version = "1.33"
'Global Const FechaModificacion = "22/07/2011"       ' Zamarbide Juan Alberto - Caso CAS-12397 - Adecuaciones GTI - Sykes
'Global Const UltimaModificacion = " "               ' Se Agregó los Items IMPORTEPAR en el cuerpo y IMPORTETOTALPAR en el pie. Los cuales agregan paréntesis a los importes del haber.
'Global Const UltimaModificacion1 = " "              ' En el caso del Importe del cuerpo si es un importe del haber se los agrega, el caso del total verifica si es negativo, remplaza el signo por dichos paréntesis.
                                                    ' Aclaración: No se toman los paréntesis como parte de la longitud.
                                                    
'Global Const Version = "1.34"
'Global Const FechaModificacion = "04/08/2011"       ' Zamarbide Juan Alberto - Caso CAS-12398 - Sykes CR - Nuevo Item Exportación Asiento
'Global Const UltimaModificacion = " "               ' Se Agregó el Item SICUENTAS. El mismo devuelve una cadena, Si una subcadena esta dentro de una cadena de la cuenta contable
'Global Const UltimaModificacion1 = " "              ' devolverá vacío, si no devolverá una subcadena de la cuenta contable. Los parámetros del Item son: SICUENTAS(X,Y,C,Q,W), donde,
                                                    ' X = posición inicial de la cadena de la cuenta, la cual se comparará
                                                    ' Y = tamaño de la cadena anterior
                                                    ' C = cadena con la que comparar la cadena dada por X e Y
                                                    ' Q = posición inicial de la cadena de la cuenta, la cual se devolverá si la comparación anterior es Falsa
                                                    ' W = tamaño de la cadena a devolver
                                                    
'Global Const Version = "1.35"
'Global Const FechaModificacion = "22/09/2011"       ' Zamarbide Juan Alberto - Caso CAS-13081 - Vision - Amarilla Gas - Exportación Asiento Contable
'Global Const UltimaModificacion = " "               ' Se volvío atrás el caso. Se quitó el el Item ASIENTO del encabezado, agregado en la ver. 1.32 dado que existe la función VOLCOD que cumple con los requerimientos del caso.
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.36"
'Global Const FechaModificacion = "09/01/2012"       ' Zamarbide Juan Alberto - CAS-14674 - CARDIF - GESTION COMPARTIDA - Adecuacion de Exportacion de Asiento
'Global Const UltimaModificacion = " "               ' Se Crearon los siguientes Items -> Cuerpo: IMPESPS - Devuelve el importe sin signo con espacios a la izq
'Global Const UltimaModificacion1 = " "              '                                 -> Pie: PORCC - Devuelve el porcentaje de un centro de costo con respecto al total de la cuenta,
                                                    ' se duplicó el pie, para que cicle o no con respecto a la cantidad de lineas del asiento, dependiendo o no si se utliza el PORCC
                                                    
'Global Const Version = "1.37"
'Global Const FechaModificacion = "18/01/2012"       ' Zamarbide Juan Alberto - CAS-14674 - CARDIF - GESTION COMPARTIDA - Adecuacion de Exportacion de Asiento
'Global Const UltimaModificacion = " "               ' Se Agregó el Item CUENTAZ al Pie, para poder ser utilizado por Cardif en la exportación de Asiento.
'Global Const UltimaModificacion1 = " "              ' Se Modificó el Item PORCC para que devuelva con espacios y no con ceros a la izquierda.
                                                    
'Global Const Version = "1.38"
'Global Const FechaModificacion = "01/06/2012"       ' Zamarbide Juan Alberto - CAS-15886 - BCO SANTANDER RHPROCHILE NUEVO FORMATO PARA EL ASIENTO
'Global Const UltimaModificacion = " "               ' Se Agregó el Item IMPORTECH al Cuerpo
'Global Const UltimaModificacion1 = " "              ' El mismo devuelve La parte entera del Monto y tiene la posibilidad de diferenciar por D o H, en cuyo caso, si el monto pertence al Debe y está configurado el Item con D
                                                    ' devuelve la parte entera con tantos 0 a la izquierda como la longitud configurada en el item, y si el monto pertenece al Haber, devuelve vacío. Lo mismo Ocurre si está configurado el Item con H.
                                                    
'Global Const Version = "1.39"
'Global Const FechaModificacion = "26/06/2012"       ' Sebastian Stremel - CAS-15886 - BCO SANTANDER RHPROCHILE NUEVO FORMATO PARA EL ASIENTO
'Global Const UltimaModificacion = " "               ' Se Agregó nuevos items para que complete con ceros o sin ceros los montos de los items
'Global Const UltimaModificacion1 = " "
                                                    
'Global Const Version = "1.40"
'Global Const FechaModificacion = "04/07/2012"       ' Deluchi Ezequiel - CAS-16153 - Se agrego un nuevo item
'Global Const UltimaModificacion = "Se Agregó un nuevo item Bajar de linea "
'Global Const UltimaModificacion1 = " "
                                                    
'Global Const Version = "1.41"
'Global Const FechaModificacion = "25/07/2012"       ' Deluchi Ezequiel - CAS-15902 - Si la tabla sistema tiene cargada la columna sis_expseguridad
'Global Const UltimaModificacion = ""                ' se mueven los archivos a la carpeta que esta cargada.
'Global Const UltimaModificacion1 = " "
                                                    
'Global Const Version = "1.42"
'Global Const FechaModificacion = "27/07/2012"       ' Deluchi Ezequiel - CAS-16171 - Se genera el archivo en 3 partes, detalle, encabezado y pie.
'Global Const UltimaModificacion = ""                ' Correcion en el corte del bucle de balanceo.
'Global Const UltimaModificacion1 = " "              ' Se agrego item CUENTASC, que para trae las cuentas sin completar con espacios.
'                                                    ' Se agrego item IMPORTETOTAL para el encabezado
                                                    
                                                    
                                                    
'Global Const Version = "1.43"
'Global Const FechaModificacion = "21/08/2012"       ' Deluchi Ezequiel - CAS-16171 - MIMO -
'Global Const UltimaModificacion = ""                ' Se agrego al encabezado los items IMPORTETOTALDH D,[C] y IMPORTETOTALDH H,[C],
'Global Const UltimaModificacion1 = " "              ' imprime total de Haber o Debe, con opcion de completar con 0's [C]="S".
'------------------------------------
'Global Const Version = "1.44"
'Global Const FechaModificacion = "19/09/2012"       ' CAS-16625 - SOS - ITEMS  CONTABLE - FGZ & NG
'Global Const UltimaModificacion = ""                ' Se cambió el nombre de los items IMPORTETOTALDH D,[C] y IMPORTETOTALDH H,[C],
'Global Const UltimaModificacion1 = " "              ' Ahora se llaman TOTALDEBEHABER D,[C] y TOTALDEBEHABER H,[C],
                                                    ' NG - Se modifico la fnc ImporteTotalDebeHaber() para que calcule el total del D y H
                                                    
'Global Const Version = "1.45"
'Global Const FechaModificacion = "21/09/2012"       ' Gonzalez Nicolás - Se nivelan versiones extraviadas.
'Global Const UltimaModificacion = ""                ' Se niveló: 1.38 - 07/02/2012
'Global Const UltimaModificacion1 = " "              ' FGZ - CAS-14674 - CARDIF - GESTION COMPARTIDA - Adecuacion de Exportacion de Asiento
'                                                    ' Se Modificó el Item PORCC en el detalle. Si la cuenta no tiene apertura por CC ==> la linea NO debe salir.
'                                                    ' ------
'                                                    ' Se niveló: 1.39 - 10/02/2012 - Gonzalez Nicolás - CAS-14674 - CARDIF - GESTION COMPARTIDA -
'                                                    ' Se blanquea la variable Aux_Linea al finalizar el loop general
'                                                    ' ------
'                                                    ' Se niveló: 1.40 - 23/02/2012 - Deluchi Ezequiel - CAS-13764 - H&A - Visualizacion de Archivos Externos -
'                                                    ' Se valida que exista la carpeta \PorUsr y \usuario, si no existen las crea.
                                                    

'Global Const Version = "1.46"
'Global Const FechaModificacion = "12/10/2012"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' Se desarrollo caso CAS - 16908 - CARDIF (GC)
'Global Const UltimaModificacion1 = " "              ' se creo el item ImporteCtroCostos
                                                    
'Global Const Version = "1.47"
'Global Const FechaModificacion = "23/10/2012"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' Se corrigieron errores que se producian cuando no ponian items en el encabezado y pie - caso CAS - 16908 - CARDIF (GC)
'Global Const UltimaModificacion1 = " "              '


'Global Const Version = "1.48"
'Global Const FechaModificacion = "17/01/2013"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' Se corrigieron cuando armaba el nombre del modelo 14 - caso CAS - 16908 - CARDIF (GC)
'Global Const UltimaModificacion1 = " "              '

'Global Const Version = "1.49"
'Global Const FechaModificacion = "29/01/2013"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' Se creo un nuevo item TotalGrupo(1,10) para cardif, lo que hace es totalizar por nro de cuentas iguales
'Global Const UltimaModificacion1 = " "              ' CAS-16908 - GESTION COMPARTIDA - Custom en Exportación de Asiento

'Global Const Version = "1.50"
'Global Const FechaModificacion = "30/01/2013"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' se completa el item TotalGrupo(1,10) con espacios si el valor ocupa menos lugares que la long del item
'Global Const UltimaModificacion1 = " "              ' CAS-16908 - GESTION COMPARTIDA - Custom en Exportación de Asiento

'Global Const Version = "1.51"
'Global Const FechaModificacion = "13/02/2013"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' Se modifica El item IMPORTECTROCOSTOS completa a izquierda con
                                                    ' espacios cuando tiene monto y cuando no tiene monto debe completa
                                                    ' con espacios en ambos casos hasta la longitud del ítem.
                                                    ' Se Elimino el separador de miles.
'Global Const UltimaModificacion1 = " "              ' CAS-16908 - GESTION COMPARTIDA - Custom en Exportación de Asiento - 2

'Global Const Version = "1.52"
'Global Const FechaModificacion = "15/02/2013"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' Se corrigio error que se producia cuando se seleccionaba todos los procesos de volcado
                                                    ' tenia en cuenta todos los procesos del periodo elegido pero no contemplaba la empresa seleccionada
'Global Const UltimaModificacion1 = " "              ' CAS-18413- CCU - ERROR EN EXPORTACION DE ASIENTO CONTABLE

'Global Const Version = "1.53"
'Global Const FechaModificacion = "21/02/2013"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' En el tipo de archivo 14, el nombre del archivo el año paso de 4 digitos a 2, ej(2012-->12)
'Global Const UltimaModificacion1 = " "              ' CAS-16908 - GESTION COMPARTIDA - Custom en Exportación de Asiento - 3

'Global Const Version = "1.54"
'Global Const FechaModificacion = "28/02/2013"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' Se realizaron cambios para que si es la ultima linea del archivo no haga salto de linea.
'Global Const UltimaModificacion1 = " "              ' CAS-16908 - GESTION COMPARTIDA - Custom en Exportación de Asiento - 4

'Global Const Version = "1.55"
'Global Const FechaModificacion = "01/03/2013"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' Se creo el item primerCampo,texto1,texto el cual completa con una C si es la primer linea o una D si no lo es.
'Global Const UltimaModificacion1 = " "              ' CAS-18139 - RHPro Consulting - Nuevo Item para Exportacion Contable

'Global Const Version = "1.56"
'Global Const FechaModificacion = "10/04/2013"       ' Sebastian Stremel -
'Global Const UltimaModificacion = ""                ' Se optimizo el tiempo del proceso modificando la forma en que cicla por cada uno de los procesos.
'Global Const UltimaModificacion1 = " "              ' CAS-18413- CCU - ERROR EN EXPORTACION DE ASIENTO CONTABLE

'Global Const Version = "1.57"
'Global Const FechaModificacion = "23/04/2013"       ' Margiotta, Emanuel (04-19 - CAS-17795- AMARILLA GAS- ESTIMAR CUSTOM ASIENTO CONTABLE) -
'Global Const UltimaModificacion = ""                ' Se agrego en la funciona parametros una opcion de generar archivos separados o no.
                                                    ' Se agrego un nuevo item (CUENTA_DH_VAR)
'Global Const UltimaModificacion1 = " "              ' CAS-17795- AMARILLA GAS- ESTIMAR CUSTOM ASIENTO CONTABLE

'Global Const Version = "1.58"
'Global Const FechaModificacion = "09/05/2013"       ' Carmen Quintero (CAS-19356 - VSO - Praxair Py - Error Item Contable) -
'Global Const UltimaModificacion = ""                ' Se modificó la función Importe_Format, para que el parámetro Monto sea de tipo Double
'                                                    ' en vez de Single.
'Global Const UltimaModificacion1 = " "              ' CAS-19356 - VSO - Praxair Py - Error Item Contable

'Global Const Version = "1.59"
'Global Const FechaModificacion = "20/05/2013"       ' Carmen Quintero (CAS-19356 - VSO - Praxair Py - Error Item Contable) -
'Global Const UltimaModificacion = ""                ' Se modificó para que se reemplace el caracter / por vacio, cuando se encuentre contenido en la descripcion
                                                    ' del proceso volcado
'Global Const UltimaModificacion1 = " "              ' CAS-19356 - VSO - Praxair Py - Error Item Contable


'Global Const Version = "1.60"
'Global Const FechaModificacion = "03/06/2013"       ' Sebastian Stremel CAS-17795 - AMARILLA GAS - MODIFICACION DE LOS NOMBRES DE LA EXPORTACION DE LOS ARCHIVOS
'Global Const UltimaModificacion = ""                ' Se creo el modelo de archivo 15 que cual si se selecciona que se generen los archivos por separados le pone a la cabecera(head) y al detalle(item)
                                                    '
'Global Const UltimaModificacion1 = " "              '

'Global Const Version = "1.61"
'Global Const FechaModificacion = "19/06/2013"       ' LED - CAS-19041 - PRAXAIR PY (PARTNER VISION) -  NUEVO ITEM CONTABLE
'Global Const UltimaModificacion = ""                ' Se creo nuevo programa CUENTAF [N,M] funciona igual que cuenta pero no completa con espacios
'                                                    '
'Global Const UltimaModificacion1 = " "              '


'Global Const Version = "1.62"
'Global Const FechaModificacion = "07/08/2013"       ' FGZ - CAS-20753 - NGA- NVS-AR - LIQ - Custom exportación asiento contable
'Global Const UltimaModificacion = ""                ' Se agregó la posibilidad de utilizar el item TAB tanto al encabezado como al PIE. Hasta hoy solo es posible utilizar en detalle.
'Global Const UltimaModificacion1 = " "              '

'Global Const Version = "1.63"
'Global Const FechaModificacion = "26/08/2013"       ' MDZ - CAS-19483 - BANCO INDUSTRIAL - CUSTOM NOMBRES DE ARCHIVOS SFTP
'Global Const UltimaModificacion = ""                ' se agrego case 16 al generar el nombre del archivo  (CCBSCH000ddmmaaaahhmmss.txt)
'Global Const UltimaModificacion1 = " "              '


'Global Const Version = "1.64"
'Global Const FechaModificacion = "01/10/2013"       ' Carmen Quintero - CAS-21568 - VSO - Praxair Py -Modificacion Item Contable
'Global Const UltimaModificacion = ""                ' Se cambio la logica del Item CUENTAF por la logica del item SICTA, con la diferencia que no respeta la longitud del campo si el valor devuelto por el ítem es menor.
'Global Const UltimaModificacion1 = " "              '

'Global Const Version = "1.65"
'Global Const FechaModificacion = "05/11/2013"       ' Deluchi Ezequiel - CAS-21994 - SGS - Custom Programa a Configurar Items de Asiento
'Global Const UltimaModificacion = ""                ' Se agrego nuevo item ImporteABSSD [S], trae el importe en el detalle sin separador de decimales y sin signo, completa con espacios a la izquierda (parametro S)
'Global Const UltimaModificacion1 = " "              '

'Global Const Version = "1.66"
'Global Const FechaModificacion = "06/11/2013"       ' Sebastian Stremel - CAS-21568 - VSO - Praxair Py - Nuevo Item Contable
'Global Const UltimaModificacion = ""                ' Se agrego nuevo item DIMPORTEF[C], trae el importe con decimales, separado por el separador del modelo 234.
'Global Const UltimaModificacion1 = " "              '

'Global Const Version = "1.67"
'Global Const FechaModificacion = "06/12/2013"       ' Borrelli Facundo - CAS-21998 - SGS - Custom Programa a Configurar Items de Asiento
'Global Const UltimaModificacion = ""                ' Se agrego el parametro TipoArchivo en la funcion generarArchivo y en la llamada a la funcion.
'Global Const UltimaModificacion1 = " "              ' Para que se genere de forma correcta la exportacion a .txt

'Global Const Version = "1.68"
'Global Const FechaModificacion = "07/01/2014"       ' Carmen Quintero - CAS-23112 - GC - Cardif - Error Exportacion Contable
'Global Const UltimaModificacion = ""                ' Se agrego la funcion truncar en los items TotalGrupo y ImporteCtroCostos
'Global Const UltimaModificacion1 = " "              ' Se cambio la manera de generar el nro de version del archivo

'Global Const Version = "1.69"
'Global Const FechaModificacion = "16/01/2014"       ' LED - CAS-23341 - ZOETIS COLOMBIA - Item exportación asiento contable
'Global Const UltimaModificacion = ""                ' Se agrego programa al detalle FINCUENTA, que dependiendo de como termine la cuenta hace devuelve un valor determinado
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.70"
'Global Const FechaModificacion = "17/01/2014"       ' LED - CAS-23341 - ZOETIS COLOMBIA - Item exportación asiento contable
'Global Const UltimaModificacion = ""                ' Correccion en la consulta que busca el empleado para obtener luego la estructura
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.71"
'Global Const FechaModificacion = "06/03/2014"       ' Carmen Quintero - CAS-23112 - GC - Cardif - Error Exportacion Contable
'Global Const UltimaModificacion = ""                ' Se agrego el parametro empresa en la función generarArchivo
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.72"
'Global Const FechaModificacion = "03/04/2014"       ' Carmen Quintero - CAS-23112 - GC - Cardif - Error Exportacion Contable
'Global Const UltimaModificacion = ""                ' Se modificó la funcion truncar en los items TotalGrupo y ImporteCtroCostos
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.73"
'Global Const FechaModificacion = "19/05/2014"       ' Carmen Quintero - CAS-24958 - VISION - BUG EN ITEMS DE ASIENTO CONTABLE
'Global Const UltimaModificacion = ""                ' Se modificó el item ImporteF
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.74"
'Global Const FechaModificacion = "21/05/2014"       ' Carmen Quintero - CAS-24958 - VISION - BUG EN ITEMS DE ASIENTO CONTABLE
'Global Const UltimaModificacion = ""                ' Se modifico la función Importe_Format, para que no coloque un espacio en blanco para el caso
'Global Const UltimaModificacion1 = " "              ' cuando el monto va en el debe


'Global Const Version = "1.75"
'Global Const FechaModificacion = "09/06/2014"       ' Carmen Quintero - CAS-24958 - VISION PRAXAIR PY - BUG EN ITEM CUENTAF
'Global Const UltimaModificacion = ""                ' Se modifico el item CUENTAF a la version original
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.76"
'Global Const FechaModificacion = "12/06/2014"       ' Carmen Quintero - CAS-24958 - VISION PRAXAIR PY - CUSTOM NUEVO PROGRAMA DE ASIENTO CONTABLE
'Global Const UltimaModificacion = ""                ' Se cambio el nombre del programa CUENTAF a CUENTA2F
'Global Const UltimaModificacion1 = " "              ' Se cambio la logica del Item CUENTAF por la logica del item SICTA, con la diferencia que no respeta la longitud del campo si el valor devuelto por el ítem es menor.
                                                    '(CAS-21568 - VSO - Praxair Py -Modificacion Item Contable)
'Global Const Version = "1.77"
'Global Const FechaModificacion = "01/07/2014"       ' Carmen Quintero - CAS-22808 - SGS - Bug Reporte Asiento Distribución Contable
'Global Const UltimaModificacion = ""                ' Se modifico la funcion del programa ImporteABSSD, por aplicarse redondeo a 2 digitos al monto
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.78"
'Global Const FechaModificacion = "02/07/2014"       ' Carmen Quintero - CAS-26116 - AMARILLA GAS - BUG EN ITEM IMPORTE AL EXPORTAR ASIENTO CONTABLE
'Global Const UltimaModificacion = ""                ' Se creo un nuevo item cuyo programa ImporteN.
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.79"
'Global Const FechaModificacion = "20/08/2014"       ' Sebastian Stremel - CAS-26810 - PRUDENTIAL - Nuevo Programa Item Asiento
'Global Const UltimaModificacion = ""                ' Se crearon 2 items nuevos IMPORTECTA [C] LINEAS [C]
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.80"
'Global Const FechaModificacion = "25/08/2014"       ' Sebastian Stremel - CAS-26810 - PRUDENTIAL - Nuevo Programa Item Asiento [Entrega 2]
'Global Const UltimaModificacion = ""                ' Se modifica el item LINEAS [C] para que muestre que numero de renglon se esta imprimiendo
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.81"
'Global Const FechaModificacion = "02/09/2014"       ' Carmen Quintero - CAS-19483 - BANCO INDUSTRIAL - BUG EN EXPORTAR ASIENTO
'Global Const UltimaModificacion = ""                ' Se agregó el item comprobante para el cliente banco industrial
'Global Const UltimaModificacion1 = " "              ' Margiotta Emanuel - Se agrego un nuevo item para SGS (ImporteABSSR). Funciona excatamente igual que ImporteABSSD con la diferencia que no redondea el monto maneja los digitos decimales.

'Global Const Version = "1.82"
'Global Const FechaModificacion = "23/09/2014"        ' Carmen Quintero - CAS-19483 - BANCO INDUSTRIAL - BUG EN ITEM  COMPROBANTE
'Global Const UltimaModificacion = ""                 ' Se agregó parámetro orden en la consulta principal del proceso
'Global Const UltimaModificacion1 = " "               ' Margiotta Emanuel - Se agrego un nuevo item para SGS (ImporteABSCR3). Funciona excatamente igual que ImporteABSSD con la diferencia que no redondea el monto maneja los digitos a 3 decimales.

'Global Const Version = "1.83"
'Global Const FechaModificacion = "04/11/2014"        ' LED - CAS-27758 - CLARIN - Nuevo Item Asiento DH variable
'Global Const UltimaModificacion = ""                 ' Nuevo item, progama Cuenta_DH_FIJ_VAR N1-N2,N3-N4
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.84"
'Global Const FechaModificacion = "13/11/2014"        ' CAS-27913 - SGS - BUG Exportacion Asiento Contable
'Global Const UltimaModificacion = ""                 ' Correcion en balanceo del item IMPORTEABSCR3, muestra 2 decimales ahora
'Global Const UltimaModificacion1 = " "

'---------------------------------------------------------------------------------------------------------------------------
'LED - VERSION NO LIBERADA - CAS-27913 - SGS - BUG Exportacion Asiento Contable
' Correcion en balanceo del item IMPORTEABSCR3 cambio de lineas al debe o haber - el item pasa a ser custom
'---------------------------------------------------------------------------------------------------------------------------

'Global Const Version = "1.85"
'Global Const FechaModificacion = "16/12/2014"        ' CAS-19483 - BANCO INDUSTRIAL - BUG EN ITEM  COMPROBANTE [Entrega 2]
'Global Const UltimaModificacion = "BUG EN ITEM COMPROBANTE [Entrega 2]" ' Se cuentan las estructuras ordenadas por cuenta.
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.86"
'Global Const FechaModificacion = "14/01/2015"        ' Miriam Ruiz-CAS-28253 - GE INTERNACIONAL INC - Nombre de Exportación Asiento Contable
'Global Const UltimaModificacion = "Nombre de Exportación Asiento Contable" ' Se agrega el nombre nro 17(CCLGL.PRARGN.CCLJE.txt).
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.87"
'Global Const FechaModificacion = "02/02/2015"        ' Dimatz Rafael CAS-28985 - AMARILLA - CUSTOM MODIFICACION ITEM DEBEHABER
'Global Const UltimaModificacion = "Nombre de Exportación Asiento Contable" ' Se agrega un Item Nuevo CUENTA_DH_VAR2 00,00-00,00
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.88"
'Global Const FechaModificacion = "24/02/2015"        ' Carmen Quintero - CAS-29399 - VISION - Error en exportacion de asiento items Importe
'Global Const UltimaModificacion = ""                 ' Se agregó validacion para el caso cuando es el ultimo registro del asiento en la funcion importe
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.89"
'Global Const FechaModificacion = "09/03/2015"        ' Dimatz Rafael - CAS 28736 - Se agrego Nuevo Item CUENTATEXT para UTDT
'Global Const UltimaModificacion = ""
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.90"
'Global Const FechaModificacion = "27/04/2015"        ' Carmen Quintero - CAS-29611 - Sykes El Salvador - Error en asiento contable [Entrega 2]
'Global Const UltimaModificacion = ""                 ' Se modificaron las funciones IMPORTEABS y IMPORTECTA, para que balanceen cuando es el ultimo item del asiento
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.91"
'Global Const FechaModificacion = "15/05/2015"        ' LED - CAS-30304 - BANCO INDUSTRIAL - CUSTOM ITEM SUMATORIA DEBE Y HABER
'Global Const UltimaModificacion = ""                 ' Se creo un nuevo item ImporteTotalDH2 D,C Donde D es un carácter, si D = “S” indica se muestra el separador decimal
'                                                     ' y Donde C es un carácter, si C = “S” indica que se completa con ceros a la izquierda hasta la longitud definida en el ítem.
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.92"
'Global Const FechaModificacion = "09/06/2015"        ' LED - CAS-30304 - BANCO INDUSTRIAL - CUSTOM ITEM SUMATORIA DEBE Y HABER (CAS-15298) [Entrega 2]
'Global Const UltimaModificacion = ""                 ' Correcion si esta configurado el item porcc en el pie, no encontraba el item ImporteTotalDH2
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.93"
'Global Const FechaModificacion = "16/06/2015"        ' Carmen Quintero - CAS-31478 - BANCO INDUSTRIAL - Error en ítem total comprobante
'Global Const UltimaModificacion = ""                 ' Se suma la variable NroSucursalInterno para obtener el valor de item TOTALCOMPROB
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.94"
'Global Const FechaModificacion = "10/07/2015"        ' Mauricio Zwenger - CAS-30801 - VISION PY - Custom item de exportacion de asiento
'Global Const UltimaModificacion = ""                 ' Se creo un nuevo item CUENTAFF similar a CUENTAF pero con un nuevo parametro
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.95"
'Global Const FechaModificacion = "13/08/2015"        ' Carmen Quintero - CAS-32038 - NGA - Custom ítems contables
'Global Const UltimaModificacion = ""                 ' - Se modifico el item TOTALREG, para que complete con ceros o no.
'Global Const UltimaModificacion1 = " "               ' - Se creo un nuevo item IMPORTETOTALHABERN similar a IMPORTETOTALDH2, pero solo para el haber y devuelve la suma en negativo

'Global Const Version = "1.96"
'Global Const FechaModificacion = "24/11/2015"        ' LM/Miriam Ruiz - CAS-33907 - PRAXAIR PY - Error en exportacion de asiento
'Global Const UltimaModificacion = ""                 ' - Se modificó el item ImporteCTA para que no duplique en la última línea
'Global Const UltimaModificacion1 = " "

'Global Const Version = "1.97"
'Global Const FechaModificacion = "14/12/2015"        ' Sebastian Stremel - CAS-33601 - RH Pro ( Producto ) - Peru – Item Contables
'Global Const UltimaModificacion = ""                 ' se corrigio error en query cuando se eligen todos los modelos.
'Global Const UltimaModificacion1 = " "               ' Se agrego formato YYMM al programa FECHAPROCVOL, se agrego item nrodoc, subcta_legajo, subcta_cc


'Global Const Version = "1.98"
'Global Const FechaModificacion = "14/12/2015"        ' Fernandez, Matias CAS-34489 - BCO INDUSTRIAL - Bug en ítem comprobante
'Global Const UltimaModificacion = ""                 ' El numero de comprobante cambia cuando cambia la surcursal.
'Global Const UltimaModificacion1 = " "               '

'Global Const Version = "1.99"
'Global Const FechaModificacion = "29/12/2015"        ' Stremel Sebastian, CAS-33601 - RH Pro (Producto) - Peru - Item Contables [Entrega 2]
'Global Const UltimaModificacion = ""                 ' Se agrega un nuevo item al detalle leg_cc, informa el numero de documento o el centro de costo segun la mascara
'Global Const UltimaModificacion1 = " "               '

'Global Const Version = "2.00"
'Global Const FechaModificacion = "06/01/2016"        ' Mauricio Zwenger - CAS-33093 - GE - Modificación Nombre y Salida Exportación Asiento Contable
'Global Const UltimaModificacion = ""                 ' se agrego fecha y hora (aaaammdd-HHmmss) a nombre de archivo CCLGL.PRARGN.CCLJE.txt
'Global Const UltimaModificacion1 = " "               '

'Global Const Version = "2.01"
'Global Const FechaModificacion = "15/01/2016"        ' Mauricio Zwenger - CAS-33093 - GE - Modificación Nombre y Salida Exportación Asiento Contable
'Global Const UltimaModificacion = ""                 ' se renombro archivo de CCLGL.PRARGN.CCLJE.txt a CCLGL.PRARGN.CCLJEES.txt y se agrego fecha y hora (aaaammdd-HHmmss)
'Global Const UltimaModificacion1 = " "               ' Se corrigio bug en generacion de archivo en UTF-8


'Global Const Version = "2.02"
'Global Const FechaModificacion = "27/01/2016"        ' Gonzalez Nicolás - CAS-35167 - ARGENOVA - Nuevo Ítem Asiento Condición
'Global Const UltimaModificacion = ""                 ' Nuevo ítem : SIPROCESO
'Global Const UltimaModificacion1 = " "               '

'Global Const Version = "2.03"
'Global Const FechaModificacion = "04/02/2016"        ' Sebastian Stremel - CAS-35167 - ARGENOVA - Nuevo Ítem Asiento Condición [Entrega 2]
'Global Const UltimaModificacion = ""                 ' Nuevo ítem : SIPROCESO
'Global Const UltimaModificacion1 = " "               ' Correccion del item siproceso - imprime el item en la linea que corresponde segun como se haya configurado


'Global Const Version = "2.04"
'Global Const FechaModificacion = "26/02/2016"        ' Fernandez, Matias - CAS-35428 - MONASTERIO BASE NCA - Bug en exportación de asientoGlobal Const UltimaModificacion = ""
'Global Const UltimaModificacion = ""                 ' se manda por parametro a la funcion generarArchivo el recordset de procesos
'Global Const UltimaModificacion1 = " "

'Global Const Version = "2.05"
'Global Const FechaModificacion = "01/03/2016"        ' Sebastian Stremel - CAS-33601 - RH Pro (Producto) - Peru - Item Contables [Entrega 3] (CAS-15298)
'Global Const UltimaModificacion = ""                 ' Se corrige programa leg_cc para que no produsca error cuando los datos son erroneos.
'Global Const UltimaModificacion1 = " "


Global Const Version = "2.06"
Global Const FechaModificacion = "09/03/2016"        ' Fernandez, Matias - CAS-35428 - MONASTERIO BASE NCA - Bug en exportación de asientoGlobal Const UltimaModificacion = ""
Global Const UltimaModificacion = ""                 ' se manda recordset de procesos y periodos al encabezado.
Global Const UltimaModificacion1 = " "

'-------------------------------------------------------------------------------------------------------------------

Public Type TR_Datos_Varios
    Convenio_Lecop As String        'String   long 4  -
    Filler As String                'String   long 1  -
    Cliente_Ya_Existente As String  'String   long 1  -
End Type

Public Type TR_Cuenta
    Cuenta  As String
    Monto As Double
End Type
Global IdUser               As String
Global Fecha                As Date
Global Hora                 As String

'Adrián - Declaración de dos nuevos registros.
Global rs_Empresa           As New ADODB.Recordset
Global rs_tipocod           As New ADODB.Recordset

Global Fecha_Inicio_periodo As Date
Global Fecha_Fin_Periodo    As Date
Global StrSql2              As String
Global SeparadorDecimales   As String
Global totalImporte         As Double
Global totalImporteD        As Double
Global totalImporteH        As Double

Global totalImporteDD       As Double   'LED - se definen nuevas variables para el item importeABSCR3
Global totalImporteHH       As Double   'LED - se definen nuevas variables para el item importeABSCR3

Global totalImporteDebe     As Double
Global totalImporteHaber    As Double
Global total                As Double
Global TotalABS             As Double
Global UltimaLeyenda        As String
Global EsUltimoItem         As Boolean
Global EsUltimoProceso      As Boolean
Global cuenta_ant           As String

Global totalImporteDH_USD   As Double
Global TotalDH_USD          As Double
Global lastpos              As Integer
Global Acuentas() As TR_Cuenta
Global Acuentas2() As TR_Cuenta
Global descMod As String
Global l_nro As Integer
Global ArchivoAux
Global ResultadoAcumuladoPorCC As Double
Global EsUltimoLineaCuenta As Boolean
Global ResultadoAcumuladoPorCC1 As Double
Global EsUltimoLineaCuenta1 As Boolean
Global ResultadoAcumuladoGrupo As Double
Global EsPrimeraLineaCuenta As Boolean
Global fExport
Global fExport2
Global fAuxiliar
Global fAuxiliarEncabezado
Global fAuxiliarDetalle
Global fAuxiliarPie

Global directorio As String
Global Archivo As String
Global intentos As Integer
Global carpeta



Public Sub Main()
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento inicial del Proceso.
' Autor      : FGZ
' Fecha      : 07/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim strCmdLine
Dim Nombre_Arch As String
Dim HuboError As Boolean
Dim rs_batch_proceso As New ADODB.Recordset
Dim PID As String
Dim bprcparam As String
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

    'Abro la conexion
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
    lastpos = 1
    Nombre_Arch = PathFLog & "Exp_Asiento_Contable" & "-" & NroProcesoBatch & ".log"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline "Version = " & Version
    Flog.writeline "Modificacion = " & UltimaModificacion
    Flog.writeline "Modificacion = " & UltimaModificacion1
    Flog.writeline "Fecha = " & FechaModificacion
    Flog.writeline "-----------------------------------------------------------------"
    Flog.writeline
    Flog.writeline "PID = " & PID
    
    'Cambio el estado del proceso a Procesando y el PID
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcprogreso = 0, bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProcesoBatch
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    
    'Obtengo los datos del proceso
    StrSql = "SELECT * FROM batch_proceso WHERE btprcnro = 27 AND bpronro =" & NroProcesoBatch
    OpenRecordset StrSql, rs_batch_proceso
    
    TiempoInicialProceso = GetTickCount
    
    If Not rs_batch_proceso.EOF Then
        IdUser = rs_batch_proceso!IdUser
        Fecha = rs_batch_proceso!bprcfecha
        Hora = rs_batch_proceso!bprchora
        bprcparam = rs_batch_proceso!bprcparam
        
        rs_batch_proceso.Close
        Set rs_batch_proceso = Nothing
        
        Call LevantarParamteros(NroProcesoBatch, bprcparam)
    Else
        Flog.writeline "No se encontro el proceso (tipo 27): " & NroProcesoBatch
    End If
    
    TiempoFinalProceso = GetTickCount
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    Flog.writeline Espacios(Tabulador * 0) & "Tiempo del proceso (milisegundos): " & (TiempoFinalProceso - TiempoInicialProceso)
    Flog.writeline Espacios(Tabulador * 0) & "=================================================="
    
    If Not HuboError Then
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProcesoBatch
    Else
        StrSql = "UPDATE batch_proceso SET bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Error' WHERE bpronro = " & NroProcesoBatch
    End If
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Flog.Close
    objconnProgreso.Close
    objConn.Close
    
End Sub
'Funcion que Valida si existe un archivo.
Function existe_archivo(Filename, l_ubioriginal)
                Dim sFilename
                Dim oFSO
                Dim pos
                Dim l_directorio
                Dim fso
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set l_directorio = fso.GetFolder(l_ubioriginal)
                sFilename = l_directorio & "\" & Filename
                Set oFSO = CreateObject("Scripting.FileSystemObject")
                

                If oFSO.FileExists(sFilename) Then
                               'El Archivo Existe
                               existe_archivo = True
                Else
                               'El Archivo NO Existe
                               existe_archivo = False
                End If
End Function 'existe_archivo(Filename,l_ubioriginal)

Function ValidarDesc(ByVal cadena As String) As String
    Dim ch As String
    Dim i As Long
    Dim cadenaAux As String
    
    cadenaAux = ""
    
    i = 1
    ch = Mid$(cadena, i, 1)
    i = i + 1
    
    Do Until i > Len(cadena) + 1
        
        Select Case Asc(ch)
            Case 47: '/'
                ch = Chr(32)
            Case 92: '\'
                ch = Chr(32)
            Case Else:
        End Select
        cadenaAux = cadenaAux & ch
        ch = Mid$(cadena, i, 1)
        i = i + 1
    Loop

ValidarDesc = cadenaAux

End Function

Public Sub Generacion(ByVal bpronro As Long, ByVal nroliq As Long, Asinro As String, ByVal Empresa As Long, ByVal TipoArchivo As Long, ByVal ProcVol As Long, ByVal ModSalidaAsiento As Long, ByVal separarArchivo As Integer, ByVal OrdenarPor As Long)

' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento de generacion del archivo de Asiento Contable
' Autor      : FGZ
' Fecha      : 25/10/2004
' Modificado : 09/08/2005 - Fapitalle N. - Se agrega la generacion de encabezado
'                                        - Se agregan casos en Programa: TAB y CUENTAZ
'                                        - Se agregan casos especiales en Programa para Shering
'                                        - Se agregan casos especiales en Programa para Halliburton
'              03/06/2006 - Fernando Favre - Se agrego el modelo de salida de asiento
'              10/06/2008 - Gustavo Ring   - Se agrego nuevo Item VOLCOD
'              06/07/2011 - Zamarbide Juan - Se agregó en el encabezado el Item "ASIENTO"
'              22/07/2011 - Zamarbide Juan - Se agregó nuevos Items IMPORTEPAR e IMPORTETOTALPAR para Sykes
'              04/08/2011 - Zamarbide Juan - Se agregó nuevo Item SICUENTAS para Sykes
'              22/09/2011 - Zamarbide Juan - Se quitó Item ASIENTO del encabezado, agregado ver. 1.32
'              05/01/2012 - Zamarbide Juan - Se agregó nuevo Item IMPESPS para Cardif ver. 1.36
'              27/01/2016 - Nicolás Gonzalez - Nuevo item SIPROCESO y se agrego conf. de confrep
' --------------------------------------------------------------------------------------------
Const ForReading = 1
Const TristateFalse = 0


'Dim directorio As String
'Dim Archivo As String
'Dim intentos As Integer
'Dim carpeta

Dim strLinea As String
Dim strLinea2 As String
Dim Aux_Linea As String
Dim cadena As String
Dim Aux_str As String
Dim tipo As String
Dim Cantidad As String
Dim posicion As String
Dim Formato  As String
Dim Nro As Long
Dim NroL As Long
Dim programa As String
Dim pos As Integer
Dim pos2 As Integer
Dim debeCod As String
Dim haberCod As String
Dim tmpStr As String
Dim separadorCampos
Dim completa As Boolean
Dim EsImporte As Boolean
Dim asi_cod_ant
Dim Enter As String
Dim Fecha_Proc As Date
Dim pliqhasta As Date
Dim TipNroDoc As Integer
Dim Fecha As String
Dim IgualA As String
Dim CambiaPor As String
Dim LongIgualAdesde As String
Dim LongIgualAhasta As String
Dim Posicion2 As String
Dim Posicion3 As String
Dim caracter As String

Dim Relleno As String
Dim ArrPar
Dim primero As Boolean
Dim PorcTotal
Dim ArrC() As String
Dim cta() As String
Dim ctadesde As Integer
Dim ctahasta As Integer
Dim asientoNro
Dim Descripcion As String

Dim Sucursal_Ant As String
Dim Sucursal As String
Dim NroSucursal As Long
Dim NroSucursalInterno As Long
Dim TotSuc As Long
Dim TotComprobante As Long
Dim dh As String                '18/11/2014 - LED - v1.85


'--- CONFREP V2.02
Dim ArrListaLineas
Dim ListaLineas As String
Dim ListaLineas2 As String
Dim ax As Long
Dim pos3 As Long

'Registros
Dim rs_Modelo As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Procesos As New ADODB.Recordset
Dim rs_Items As New ADODB.Recordset
Dim rs_Items2 As New ADODB.Recordset
Dim rs_ItemsPie As New ADODB.Recordset
Dim rs_Mod_Asiento As New ADODB.Recordset
Dim rs_Sistema As New ADODB.Recordset
Dim rs_conf As New ADODB.Recordset
Dim porcc As Boolean
Dim porcc1 As Boolean
Dim vinoEnter As Boolean

Dim rs_ItemsAux As New ADODB.Recordset
'sebastian stremel
Dim rs_desc As New ADODB.Recordset
Dim Dire As String
Dim pos1 As Integer
Dim nombreArchivoExp As String

Flog.writeline "Entro a generar "
Flog.writeline ""


'Archivo de exportacion
StrSql = "SELECT sis_dirsalidas FROM sistema WHERE sisnro = 1 "
OpenRecordset StrSql, rs
If Not rs.EOF Then
    directorio = Trim(rs!sis_dirsalidas)
End If

StrSql = "SELECT * FROM modelo WHERE modnro = 234"
OpenRecordset StrSql, rs_Modelo
If Not rs_Modelo.EOF Then
    If Not IsNull(rs_Modelo!modarchdefault) Then
        'VALIDO QUE EXISTA LA RUTA
        'directorio = directorio & Trim(rs_Modelo!modarchdefault)
        directorio = ValidarRuta(directorio, "\PorUsr", 1)
        directorio = ValidarRuta(directorio, "\" & IdUser, 1)
        directorio = ValidarRuta(directorio, "\" & Trim(rs_Modelo!modarchdefault), 1)
    Else
        Flog.writeline Espacios(Tabulador * 1) & "El modelo no tiene configurada la carpeta destino. El archivo será generado en el directorio default"
    End If
    'SeparadorDecimales = rs_Modelo!modsepdec
    'separadorCampos = rs_Modelo!modseparador
    
    '07/01/2016 - MDZ
    If TipoArchivo <> 17 Then
        SeparadorDecimales = IIf(EsNulo(rs_Modelo!modsepdec), "", rs_Modelo!modsepdec)
        separadorCampos = IIf(EsNulo(rs_Modelo!modseparador), "", rs_Modelo!modseparador)
    Else
        SeparadorDecimales = "."
        separadorCampos = ","
    End If
    
Else
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el modelo. El archivo será generado en el directorio default"
End If

'cargo el periodo
StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(nroliq)
OpenRecordset StrSql, rs_Periodo
If rs_Periodo.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "No se encontró el Periodo"
    Exit Sub
End If


ListaLineas = ""
StrSql = "SELECT confnrocol,confval, confval2 FROM confrepadv WHERE repnro = 505 "
OpenRecordset StrSql, rs_conf
If Not rs_conf.EOF Then
    Select Case CLng(rs_conf!confnrocol)
        Case 1: 'LISTA DE LINEAS A IMPRIMIR X MODELO
            If Asinro = CLng(rs_conf!confval) Then
                If EsNulo(rs_conf!confval2) = False Then
                    pos = InStr(rs_conf!confval2, "@")
                    If pos > 0 Then 'Lista con Rangos
                        ArrListaLineas = Split(rs_conf!confval2, "@")
                        pos = InStr(ArrListaLineas(0), "-")
                        If pos > 0 Then
                            pos2 = Left(ArrListaLineas(0), pos - 1)
                            pos3 = Mid(ArrListaLineas(0), pos + 1, Len(ArrListaLineas(0)))
                            ListaLineas = "0"
                            For ax = CLng(pos2) To CLng(pos3)
                                ListaLineas = ListaLineas & "," & ax
                            Next
                        End If
                        
                        If UBound(ArrListaLineas) > 0 Then
                            If Left(ArrListaLineas(1), 1) = "," Then
                                ListaLineas = ListaLineas & ArrListaLineas(1)
                            Else
                                ListaLineas = ListaLineas & "," & ArrListaLineas(1)
                            End If
                        End If
                    ElseIf InStr(rs_conf!confval2, "-") > 0 Then 'Solo x Rangos
                        pos = InStr(rs_conf!confval2, "-")
                        If pos > 0 Then
                            pos2 = Left(rs_conf!confval2, pos - 1)
                            pos3 = Mid(rs_conf!confval2, pos + 1, Len(rs_conf!confval2))
                            ListaLineas = "0"
                            For ax = CLng(pos2) To CLng(pos3)
                                ListaLineas = ListaLineas & "," & ax
                            Next
                        End If
                    Else 'Solo lista
                        ListaLineas = "0," & rs_conf!confval2
                    End If
                End If
                Flog.writeline "Se exportarán las siguientes lineas: " & Mid(ListaLineas, 3, Len(ListaLineas))
            End If
    End Select
End If
'--- CONFREP


'sebastian stremel 15/02/2013 desde aca
If ProcVol = 0 Then
    StrSql = " SELECT DISTINCT cuenta, monto, linea, desclinea, dh, linea_Asi.masinro, proc_vol.vol_cod, vol_desc, vol_fec_asiento, vol_fec_proc    "
    StrSql = StrSql & " FROM proc_vol "
    StrSql = StrSql & " INNER JOIN proc_vol_pl ON proc_vol.vol_cod = proc_vol_pl.vol_cod"
    StrSql = StrSql & " INNER JOIN proceso ON proc_vol_pl.pronro = proceso.pronro"
    StrSql = StrSql & " INNER JOIN proc_v_modasi ON proc_v_modasi.vol_cod = proc_vol.vol_cod"
    StrSql = StrSql & " INNER JOIN mod_asiento ON mod_asiento.masinro = proc_v_modasi.asi_cod"
    StrSql = StrSql & " INNER JOIN linea_asi ON proc_vol.vol_cod = linea_asi.vol_cod"
    StrSql = StrSql & " AND mod_asiento.cofcnro <> 2"
    StrSql = StrSql & " WHERE proc_vol.pliqnro = " & nroliq
    StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
    StrSql = StrSql & " AND   proceso.empnro   = " & Empresa
    StrSql = StrSql & " AND linea_asi.cuenta <> '999999.999'"
    'StrSql = StrSql & " ORDER BY vol_desc "
    'StrSql = StrSql & "ORDER BY linea"
    '23/09/2014
    If OrdenarPor = 1 Then
        StrSql = StrSql & "ORDER BY linea"
    End If
    If OrdenarPor = 2 Then
        StrSql = StrSql & " ORDER BY linea_asi.cuenta"
    End If
    'fin
    OpenRecordset StrSql, rs_Procesos
Else
    StrSql = "SELECT * FROM  proc_vol "
    StrSql = StrSql & " INNER JOIN linea_asi ON proc_vol.vol_cod = linea_asi.vol_cod "
    StrSql = StrSql & " WHERE proc_vol.pliqnro =" & nroliq
    StrSql = StrSql & " AND linea_asi.vol_cod IN (" & ProcVol & ")"
    StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
    StrSql = StrSql & " AND linea_asi.cuenta <> '999999.999'"
    
    'FILTRO LAS LINEAS A GENERAR
    'If ListaLineas <> Empty Then
    '    StrSql = StrSql & " and linea_asi.linea IN (" & ListaLineas & ")"
    'End If
    
    'StrSql = StrSql & "ORDER BY linea"
    '23/09/2014
    If OrdenarPor = 1 Then
        StrSql = StrSql & "ORDER BY linea"
    End If
    If OrdenarPor = 2 Then
        StrSql = StrSql & " ORDER BY linea_asi.cuenta"
    End If
    'fin
    OpenRecordset StrSql, rs_Procesos
End If

'hasta aca

'Busco los procesos a evaluar
'StrSql = "SELECT * FROM  proc_vol "
'StrSql = StrSql & " INNER JOIN linea_asi ON proc_vol.vol_cod = linea_asi.vol_cod "
'StrSql = StrSql & " WHERE proc_vol.pliqnro =" & nroliq
'If ProcVol <> 0 Then 'si no son todos
'    StrSql = StrSql & " AND linea_asi.vol_cod IN (" & ProcVol & ")"
'End If
'StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
'StrSql = StrSql & " AND linea_asi.cuenta <> '999999.999'"
'StrSql = StrSql & "ORDER BY linea"
'OpenRecordset StrSql, rs_Procesos


porcc = False
'------------------------------------------------------------------------
' Agregado ver. 1.36 - JAZ - CAS-14674-------------
' Genero los totales para realizar el PORCC
'------------------------------------------------------------------------
' Realizo las consultas para ver si existe el Item en el Detalle o Pie

' Por Pie en rs_ItemsPie
StrSql = "SELECT * FROM confitemicpie "
StrSql = StrSql & " INNER JOIN itemintcont ON confitemicpie.itemicnro = itemintcont.itemicnro "
If ModSalidaAsiento <> 0 Then
    StrSql = StrSql & " AND confitemicpie.moditenro = " & ModSalidaAsiento
End If
StrSql = StrSql & " WHERE itemicprog = 'PORCC'"
StrSql = StrSql & " ORDER BY confitemicpie.confitemicorden "
OpenRecordset StrSql, rs_ItemsPie

'Por Detalle en rs_Items
StrSql = "SELECT * FROM confitemic "
StrSql = StrSql & " INNER JOIN itemintcont ON confitemic.itemicnro = itemintcont.itemicnro "
If ModSalidaAsiento <> 0 Then
    StrSql = StrSql & " AND confitemic.moditenro = " & ModSalidaAsiento
End If
'StrSql = StrSql & " WHERE itemicprog = 'PORCC'"
'StrSql = StrSql & " WHERE itemicprog IN ('PORCC','IMPORTECTROCOSTOS','TOTALGRUPO(1,10)')"
StrSql = StrSql & " WHERE itemicprog IN ('PORCC','IMPORTECTROCOSTOS')"
StrSql = StrSql & " ORDER BY confitemic.confitemicorden "
OpenRecordset StrSql, rs_Items

' Verifico Previamente si se realiza por Porcc
If Not rs_Items.EOF Or Not rs_ItemsPie.EOF Then
    porcc = True
End If

' Si existe el Item, Genero los Totales
'If porcc Then
'    Do While Not rs_Procesos.EOF
'        Call AlmCuentaCC(rs_Procesos!cuenta, rs_Procesos!Monto, rs_Procesos!Linea, rs_Procesos!desclinea, rs_Procesos!dh)
'        rs_Procesos.MoveNext
'    Loop
'    rs_Procesos.MoveFirst
'End If

'Por Detalle en rs_Items
StrSql = "SELECT * FROM confitemic "
StrSql = StrSql & " INNER JOIN itemintcont ON confitemic.itemicnro = itemintcont.itemicnro "
If ModSalidaAsiento <> 0 Then
    StrSql = StrSql & " AND confitemic.moditenro = " & ModSalidaAsiento
End If
StrSql = StrSql & " WHERE itemicprog LIKE ('%TOTALGRUPO%')"
StrSql = StrSql & " ORDER BY confitemic.confitemicorden "
OpenRecordset StrSql, rs_Items2
' Verifico Previamente si se realiza por Porcc
If Not rs_Items2.EOF Then
    porcc1 = True
    'BUSCO EL PARAMETRO DE LA LONG DE LA CTA DESDE Y HASTA
        cta = Split(rs_Items2!itemicprog, ",")
        ctadesde = cta(1)
        ctahasta = cta(2)
    'HASTA ACA
    
End If

'Sebastian Stremel - 29/01/2013
'Si existe el Item, Genero los Totales
If porcc1 Or porcc Then
    If porcc And porcc1 Then
        Do While Not rs_Procesos.EOF
            Call AlmCuentaCC(rs_Procesos!Cuenta, rs_Procesos!Monto, rs_Procesos!Linea, rs_Procesos!descLinea, rs_Procesos!dh)
            Call AlmCuentaCCvariable(rs_Procesos!Cuenta, rs_Procesos!Monto, rs_Procesos!Linea, rs_Procesos!descLinea, rs_Procesos!dh, ctadesde, ctahasta)
            rs_Procesos.MoveNext
        Loop
    Else
        If porcc Then
            Do While Not rs_Procesos.EOF
                Call AlmCuentaCC(rs_Procesos!Cuenta, rs_Procesos!Monto, rs_Procesos!Linea, rs_Procesos!descLinea, rs_Procesos!dh)
            rs_Procesos.MoveNext
            Loop
        Else
            If porcc1 Then
                Do While Not rs_Procesos.EOF
                    Call AlmCuentaCCvariable(rs_Procesos!Cuenta, rs_Procesos!Monto, rs_Procesos!Linea, rs_Procesos!descLinea, rs_Procesos!dh, ctadesde, ctahasta)
                rs_Procesos.MoveNext
                Loop
            End If
        End If
    End If
    rs_Procesos.MoveFirst
End If

'Seteo el nombre del archivo generado
'FB Se agrego tipoArchivo

generarArchivo rs_Periodo, rs_Mod_Asiento, rs_desc, rs_Sistema, TipoArchivo, Asinro, nroliq, Empresa, separarArchivo, rs_Procesos

'desactivo el manejador de errores
'On Error GoTo 0
On Error GoTo CE

' Comienzo la transaccion
MyBeginTrans

''Busco los procesos a evaluar
'StrSql = "SELECT * FROM  proc_vol "
'StrSql = StrSql & " INNER JOIN linea_asi ON proc_vol.vol_cod = linea_asi.vol_cod "
'StrSql = StrSql & " WHERE proc_vol.pliqnro =" & Nroliq
'If ProcVol <> 0 Then 'si no son todos
'    StrSql = StrSql & " AND linea_asi.vol_cod IN (" & ProcVol & ")"
'End If
'StrSql = StrSql & " AND linea_asi.masinro IN (" & Asinro & ")"
'StrSql = StrSql & " AND linea_asi.cuenta <> '999999.999'"
'OpenRecordset StrSql, rs_Procesos

'seteo de las variables de progreso
PorcTotal = 100
Progreso = 0

StrSql = "SELECT * FROM confitemicpie "
StrSql = StrSql & " INNER JOIN itemintcont ON confitemicpie.itemicnro = itemintcont.itemicnro "
If ModSalidaAsiento <> 0 Then
    StrSql = StrSql & " AND confitemicpie.moditenro = " & ModSalidaAsiento
End If
StrSql = StrSql & " ORDER BY confitemicpie.confitemicorden "
OpenRecordset StrSql, rs_Items
If Not rs_Items.EOF Then PorcTotal = PorcTotal - 1
rs_Items.Close
StrSql = "SELECT * FROM confitemicenc "
StrSql = StrSql & " INNER JOIN itemintcont ON confitemicenc.itemicnro = itemintcont.itemicnro "
If ModSalidaAsiento <> 0 Then
    StrSql = StrSql & "AND confitemicenc.moditenro = " & ModSalidaAsiento
End If
StrSql = StrSql & " ORDER BY confitemicenc.confitemicorden "
OpenRecordset StrSql, rs_Items
If Not rs_Items.EOF Then PorcTotal = PorcTotal - 1
rs_Items.Close





CConceptosAProc = rs_Procesos.RecordCount
If CConceptosAProc = 0 Then
    CConceptosAProc = 1
    Flog.writeline Espacios(Tabulador * 1) & " No hay Proceso de Volcados para ese asiento en ese periodo"
Else
    Flog.writeline Espacios(Tabulador * 1) & " Lineas de Procesos de Volcados para ese asiento en ese periodo " & CConceptosAProc
End If
IncPorc = (PorcTotal / CConceptosAProc)





'Procesamiento
If rs_Procesos.EOF Then
    Flog.writeline Espacios(Tabulador * 2) & "No hay nada que procesar"
End If


'------------------------------------------------------------------------
' Genero el detalle de la exportacion
'------------------------------------------------------------------------
'rs_Procesos.MoveFirst

totalImporte = 0
totalImporteD = 0
totalImporteH = 0
totalImporteDD = 0
totalImporteHH = 0
totalImporteDebe = 0
totalImporteHaber = 0

total = 0

cuenta_ant = ""

totalImporteDH_USD = 0
TotalDH_USD = 0
ResultadoAcumuladoPorCC = 0
ResultadoAcumuladoPorCC1 = 0

'seba 29/01/2013
ResultadoAcumuladoGrupo = 0
'hasta aca

NroL = 1
UltimaLeyenda = ""
EsUltimoItem = False
EsUltimoProceso = False
EsImporte = False
vinoEnter = False
asi_cod_ant = -1
primero = True
EsUltimoLineaCuenta = False
EsUltimoLineaCuenta1 = False

Sucursal_Ant = ""
NroSucursal = 0
NroSucursalInterno = 0
TotSuc = 0
TotComprobante = 0

'If TipoArchivo = 17 Then
'          Set fExport2 = CreateObject("ADODB.Stream") 'Create the stream
'          fExport2.Type = adTypeText
'          fExport2.Charset = "UTF-8" 'Indicate the charactor encoding
'          fExport2.Open 'Initialize the stream
          
'End If
strLinea2 = ""
Do While Not rs_Procesos.EOF
        pos1 = pos1 + 1
        If EsUltimoRegistro(rs_Procesos) Then
            EsUltimoProceso = True
        End If
        If EsUltimoRegistroCuenta(rs_Procesos, rs_Procesos!Linea) Then
            EsUltimoLineaCuenta = True
            EsUltimoLineaCuenta1 = True
        End If
        
        If pos1 = 1 Then
            EsPrimeraLineaCuenta = True
        Else
            EsPrimeraLineaCuenta = False
        End If
        
        Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
        Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del proceso de volcado " & rs_Procesos!vol_cod & " Linea " & rs_Procesos!masinro & " cuenta: " & rs_Procesos!Cuenta
        Flog.writeline
        
        Cantidad_Warnings = 0
        Nro = Nro + 1 'Contador de Lineas
        
        If UCase(UltimaLeyenda) <> UCase(rs_Procesos!descLinea) Then
            NroL = NroL + 1
        End If
        UltimaLeyenda = rs_Procesos!descLinea
        
        StrSql = "SELECT * FROM confitemic "
        StrSql = StrSql & " INNER JOIN itemintcont ON confitemic.itemicnro = itemintcont.itemicnro "
        If ModSalidaAsiento <> 0 Then
            StrSql = StrSql & " AND confitemic.moditenro = " & ModSalidaAsiento
        End If
        StrSql = StrSql & " ORDER BY confitemic.confitemicorden "
        OpenRecordset StrSql, rs_Items
                    
        Aux_Linea = ""
        EsUltimoItem = False
        
        Do While Not rs_Items.EOF
            Flog.writeline Espacios(Tabulador * 2) & "Item: " & rs_Items!itemicdesabr
            'FGZ - 07/02/2012 -------- custom CARDIF
            'If porcc And Len(rs_Procesos!Cuenta) <= 11 Then 'Imprime sólo cuentas con CC -> Tamaño de Cuenta = 10 caracteres - Agregado ver 1.37
            'If porcc And Len(rs_Procesos!Cuenta) <= 11 Then 'Imprime sólo cuentas con CC -> Tamaño de Cuenta = 10 caracteres - Agregado ver 1.37
                'NO va la linea
            '    Flog.writeline Espacios(Tabulador * 3) & "Linea sin apertura. No se genera."
            'Else

            cadena = ""
            If rs_Items!itemicfijo Then
                If rs_Items!itemicvalorfijo = "" Then
                    cadena = String(256, " ")
                Else
                    cadena = rs_Items!itemicvalorfijo
                    If Len(cadena) < rs_Items!itemiclong Then
                        cadena = cadena & String(rs_Items!itemiclong - Len(cadena), " ")
                    End If
                End If
            Else
                programa = UCase(rs_Items!itemicprog)
                Flog.writeline Espacios(Tabulador * 3) & "Programa: " & programa
                Select Case programa
                '/////casos especiales de shering - start//////////
                Case "HEAD":
                    If rs_Procesos!masinro <> asi_cod_ant Then
                        Call Hacer_Header(rs_Procesos!dh, rs_Procesos!Cuenta, rs_Procesos!masinro, rs_Procesos!vol_fec_asiento, cadena)
                        asi_cod_ant = rs_Procesos!masinro
                        'fExport.writeline cadena
                        fAuxiliarDetalle.writeline cadena
                        cadena = ""
                    End If
                Case "DESCITEM":
                    Select Case Len(rs_Procesos!Cuenta)
                        Case 19:
                            cadena = "ITEMA"
                        Case 10:
                            cadena = "ITEMS"
                        Case 14:
                            cadena = "ITEMS"
                        Case Else: 'no deberia darse
                            cadena = "ITEMX"
                    End Select
                Case "IMPSHERING":
                    Call ImporteABS_3(rs_Procesos!Monto, rs_Procesos!dh, True, rs_Items!itemiclong, cadena)
                Case "FECHASHERING":
                    Select Case Len(rs_Procesos!Cuenta)
                        Case 19:
                            cadena = Format(rs_Procesos!vol_fec_asiento, "DDMMYYYY") + Mid(rs_Procesos!Cuenta, 14, 5)
                        Case Else:
                            cadena = "|"
                    End Select
                    If Len(cadena) < rs_Items!itemiclong Then
                        cadena = cadena + String(rs_Items!itemiclong - Len(cadena), " ")
                    Else
                        cadena = Left(cadena, rs_Items!itemiclong)
                    End If
                Case "TEXSHERING":
                    Select Case rs_Procesos!masinro
                        Case 1:
                            cadena = "HABERES Y RETENCIONES"
                        Case 2:
                            cadena = "APORTES PATRONALES"
                        Case 3:
                            cadena = "PREVISIONES"
                        Case 4:
                            cadena = "INTERES S/PRESTAMO " + Format(rs_Procesos!vol_fec_asiento, "MM/YYYY")
                        Case Else: 'no deberia darse
                            cadena = "<< masinro > 4 >>"
                    End Select
                    If Len(cadena) < rs_Items!itemiclong Then
                        cadena = cadena + String(rs_Items!itemiclong - Len(cadena), " ")
                    Else
                        cadena = Left(cadena, rs_Items!itemiclong)
                    End If
                Case "COSTOPRODUCTO":
                    Select Case Len(rs_Procesos!Cuenta)
                        Case 19:
                            cadena = String(12, " ")
                        Case 10:
                            cadena = "|" + String(3, " ") + "|" + String(7, " ")
                        Case 14:
                            cadena = Mid(rs_Procesos!Cuenta, 11, 4) + "|" + String(7, " ")
                        Case Else:  'no deberia darse
                            cadena = "| LW" + "|" + String(7, " ")
                    End Select
                Case "CTACONTABLE":
                    If rs_Procesos!dh Then
                        cadena = "40"
                    Else
                        cadena = "50"
                    End If
                    Select Case Len(rs_Procesos!Cuenta)
                        Case 19:
                            If rs_Procesos!masinro = 1 Then
                                cadena = "39"
                            End If
                            If rs_Procesos!masinro = 4 Then
                                cadena = "29"
                            End If
                            cadena = cadena + "00" + Mid(rs_Procesos!Cuenta, 11, 9) + String(6, " ")
                        Case 10:
                            cadena = cadena + rs_Procesos!Cuenta + " | | | "
                        Case 14:
                            cadena = cadena + Mid(rs_Procesos!Cuenta, 1, 10) + " | | | "
                        Case Else: 'no deberia darse
                            cadena = cadena + "<LENWRONG>" + " | | | "
                    End Select
                Case "PIESHERING":
                    If Hacer_Pie(rs_Procesos) Then
                        cadena = Mid(Aux_Linea, 1, 110) + String(13, " ") + Mid(Aux_Linea, 124, 6)
                        If rs_Procesos!masinro = 4 Then
                            cadena = Mid(cadena, 1, 49) + "INT.S/PREST.-ANTIC.AL PERSONAL                   " + Mid(cadena, 99, 30)
                        End If
                        'fExport.writeline cadena
                        fAuxiliarDetalle.writeline cadena
                        cadena = ""
                        Aux_Linea = "FINAL"
                    End If
                '/////casos especiales de shering - end //////////////
                
                '/////caso especial de halliburton - start //////////////
                Case "SAPLINE":
                    strLinea = ";"
                    Call NroCuenta(rs_Procesos!Cuenta, 1, 10, True, 10, cadena)
                    strLinea = strLinea & """" & cadena & """" & ";" 'SAPG/L account
                    If rs_Procesos!dh Then
                        strLinea = strLinea & "  "
                    Else
                        strLinea = strLinea & "- "
                    End If 'signo
                    strLinea = strLinea & Format(rs_Procesos!Monto, "00000000.00") & ";" 'amount
                    If (cadena = "0000640355") Or (cadena = "0000147180") Then
                        strLinea = strLinea & "V0"
                    End If
                    strLinea = strLinea & ";" 'taxcode
                    strLinea = strLinea & """" & Mid(rs_Procesos!Cuenta, 11, 10) & """" & ";" 'cost center
                    strLinea = strLinea & """" & Mid(rs_Procesos!Cuenta, 21, 12) & """" & ";" 'internal order
                    strLinea = strLinea & """" & Mid(rs_Procesos!Cuenta, 33, 10) & """" & ";" 'profit center
                    strLinea = strLinea & """" & Mid(rs_Procesos!Cuenta, 43, 8) & """" & ";" 'personnel number
                    strLinea = strLinea & ";" 'inter-company
                    strLinea = strLinea & ";" 'allocation
                    Call Leyenda(rs_Procesos!descLinea, 1, 50, True, 50, cadena)
                    strLinea = strLinea & cadena & ";" 'line item text
                    strLinea = strLinea & ";" 'quantity
                    strLinea = strLinea & ";" 'uom
                    strLinea = strLinea & ";" 'wbs element
                    strLinea = strLinea & ";" 'network
                    strLinea = strLinea & ";" 'activity
                    strLinea = strLinea & ";" 'tp profit center
                    strLinea = strLinea & ";" 'trading partner
                    strLinea = strLinea & ";" 'settlement period
                    strLinea = strLinea & ";" 'tax jur code
                    strLinea = strLinea & ";" 'asset trans type
                    strLinea = strLinea & ";" 'tax tran type
                    cadena = strLinea
                '/////caso especial de halliburton - end //////////////

                '/////caso especial de MARHS - start //////////////
                Case "DOCINCTA 0,0,0" To "DOCINCTA 99,99,99":
                    If Len(programa) > 9 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 10, pos - 10)
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Cantidad = Mid(programa, pos + 1, pos2 - pos - 1)
                        TipNroDoc = Mid(programa, pos2 + 1, Len(programa) - pos2)
                        Call DOCinCTA(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), CLng(TipNroDoc), True, 0, cadena)
                    Else
                        Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & ". Son 3 parametros, posicion inicil, longitud y Tipo de Documento."
                    End If
                Case "IMPORTED" To "IMPORTED Z":
                    If Len(programa) > 8 Then
                        posicion = Trim(Mid(programa, 10, 1))
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = False
                    End If
'                    If EsUltimoRegistroItem(rs_Items) Then
'                        EsUltimoItem = True
'                    End If
                    If rs_Procesos!dh Then
                        Call ImporteDH(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                    Else
                        Call ImporteDH(0, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                    End If
                Case "IMPORTEH" To "IMPORTEH Z":
                    If Len(programa) > 8 Then
                        posicion = Trim(Mid(programa, 10, 1))
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = False
                    End If
'                    If EsUltimoRegistroItem(rs_Items) Then
'                        EsUltimoItem = True
'                    End If
                    If rs_Procesos!dh Then
                        Call ImporteDH(0, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                    Else
                        Call ImporteDH(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                    End If
                
                'MAF - 27/10/2006
                Case "COTIZAUSD A,A" To "COTIZAUSD Z,Z":
                    pos = CLng(InStr(1, programa, ","))
                    If pos = 0 Then
                        cadena = " ERROR "
                        Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
                    Else
                        Fecha = Trim(Mid(programa, 11, pos - 11))
                        completa = (UCase(Mid(programa, pos + 1, Len(programa) - pos)) = "S")
                        Select Case Fecha
                            Case "A":
                                Call BusCotiza(rs_Procesos!vol_fec_asiento, completa, rs_Items!itemiclong, cadena)
                            Case "H":
                                Call BusCotiza(rs_Periodo!pliqhasta, completa, rs_Items!itemiclong, cadena)
                            Case Else
                                Call BusCotiza(rs_Procesos!vol_fec_asiento, completa, rs_Items!itemiclong, cadena)
                        End Select
                    End If
                    
                'MAF - 27/10/2006
                Case "IMPORTEDUSD A,A" To "IMPORTEDUSD Z,Z":
                    pos = CLng(InStr(1, programa, ","))
                    
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    
                    If pos = 0 Then
                        cadena = " ERROR "
                        Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
                    Else
                                            
                        Fecha = Trim(Mid(programa, 13, pos - 13))
                        completa = (UCase(Mid(programa, pos + 1, Len(programa) - pos)) = "S")
                    
                        Select Case Fecha
                            Case "A":
                                If rs_Procesos!dh Then
                                    Call ImporteUSD(rs_Procesos!vol_fec_asiento, rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                Else
                                    Call ImporteUSD(rs_Procesos!vol_fec_asiento, 0, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                End If
                            Case "H":
                                If rs_Procesos!dh Then
                                    Call ImporteUSD(rs_Periodo!pliqhasta, rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                Else
                                    Call ImporteUSD(rs_Periodo!pliqhasta, 0, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                End If
                            Case Else
                                If rs_Procesos!dh Then
                                    Call ImporteUSD(rs_Procesos!vol_fec_asiento, rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                Else
                                    Call ImporteUSD(rs_Procesos!vol_fec_asiento, 0, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                End If
                        End Select
                    End If
                
                'MAF - 27/10/2006
                Case "IMPORTEHUSD A,A" To "IMPORTEHUSD Z,Z":
                    pos = CLng(InStr(1, programa, ","))
                    
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    
                    If pos = 0 Then
                        cadena = " ERROR "
                        Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
                    Else
                                            
                        Fecha = Trim(Mid(programa, 13, pos - 13))
                        completa = (UCase(Mid(programa, pos + 1, Len(programa) - pos)) = "S")
                    
                        Select Case Fecha
                            Case "A":
                                If rs_Procesos!dh Then
                                    Call ImporteUSD(rs_Procesos!vol_fec_asiento, 0, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                Else
                                    Call ImporteUSD(rs_Procesos!vol_fec_asiento, rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                End If
                            Case "H":
                                If rs_Procesos!dh Then
                                    Call ImporteUSD(rs_Periodo!pliqhasta, 0, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                Else
                                    Call ImporteUSD(rs_Periodo!pliqhasta, rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                End If
                            Case Else
                                If rs_Procesos!dh Then
                                    Call ImporteUSD(rs_Procesos!vol_fec_asiento, 0, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                Else
                                    Call ImporteUSD(rs_Procesos!vol_fec_asiento, rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                                End If
                        End Select
                    End If
                    
                'MAF - 27/10/2006
                Case "INTERFACE" To "INTERFACE Z":
                    If Len(programa) > 9 Then
                        Fecha = Trim(Mid(programa, 11, 1))
                        Select Case Fecha
                            Case "A":
                                Call InterfaceFecha(rs_Procesos!vol_fec_asiento, rs_Items!itemiclong, cadena)
                            Case "H":
                                Call InterfaceFecha(rs_Periodo!pliqhasta, rs_Items!itemiclong, cadena)
                            Case Else
                                Call InterfaceFecha(rs_Procesos!vol_fec_asiento, rs_Items!itemiclong, cadena)
                        End Select
                    Else
                        Call InterfaceFecha(rs_Procesos!vol_fec_asiento, rs_Items!itemiclong, cadena)
                    End If

                'MAF - 27/10/2006
                Case "SECFECHA" To "SECFECHA Z":
                    If Len(programa) > 8 Then
                        Fecha = Trim(Mid(programa, 10, 1))
                        Select Case Fecha
                            Case "A":
                                Call SecuenciaFecha(rs_Procesos!vol_cod, rs_Procesos!vol_fec_asiento, rs_Items!itemiclong, cadena)
                            Case "H":
                                Call SecuenciaFecha(rs_Procesos!vol_cod, rs_Periodo!pliqhasta, rs_Items!itemiclong, cadena)
                            Case Else
                                Call SecuenciaFecha(rs_Procesos!vol_cod, rs_Procesos!vol_fec_asiento, rs_Items!itemiclong, cadena)
                        End Select
                    Else
                        Call SecuenciaFecha(rs_Procesos!vol_cod, rs_Procesos!vol_fec_asiento, rs_Items!itemiclong, cadena)
                    End If
                    
                
                '/////caso especial de MARHS - end //////////////
                
                Case "ESPACIOS":
                    cadena = String(rs_Items!itemiclong, " ")
                Case "TAB 1" To "TAB 9":
                    If Len(programa) > 4 Then
                        Cantidad = Mid(programa, 5, 1)
                    Else
                        Cantidad = "1"
                    End If
                    cadena = String(CLng(Cantidad), Chr(9))
                Case "CODMODASI" To "CODMODASI 99,99":
                    If Len(programa) > 10 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 10, pos - 10)
                        Cantidad = Mid(programa, pos + 1, Len(programa) - pos)
                        Call CodModAsiento(rs_Procesos!masinro, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    Else
                        posicion = "1"
                        Cantidad = rs_Items!itemiclong
                        Call CodModAsiento(rs_Procesos!masinro, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    End If
                Case "CUENTASC":
                        'Posicion 1 para que empieze del primer caracter
                        'Cantidad 100 porque es el tamaño maximo de caracteres para una cuenta
                        'Completar False para que no complete ni con espacion ni 0's
                        Call NroCuenta_n(rs_Procesos!Cuenta, 1, 100, False, rs_Items!itemiclong, cadena)
                
                Case "CUENTAN" To "CUENTAN 99,99":
                    If Len(programa) > 7 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 8, pos - 8)
                        Cantidad = Mid(programa, pos + 1, Len(programa) - pos)
                        Call NroCuenta_n(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    Else
                        posicion = "1"
                        Cantidad = rs_Items!itemiclong
                        Call NroCuenta_n(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    End If
                Case "CUENTZ" To "CUENTZ 99,99":
                    If Len(programa) > 7 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 8, pos - 8)
                        Cantidad = Mid(programa, pos + 1, Len(programa) - pos)
                        Call NroCuenta_1(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    Else
                        posicion = "1"
                        Cantidad = "10"
                        Call NroCuenta_1(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    End If
                                 
                'sebastian stremel 13/02
                Case "CUENTZ2" To "CUENTZ2 99,99":
                    If Len(programa) > 7 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 8, pos - 8)
                        Cantidad = Mid(programa, pos + 1, Len(programa) - pos)
                        Call NroCuenta_2(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    Else
                        posicion = "1"
                        Cantidad = "10"
                        Call NroCuenta_2(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    End If
                'hasta aca
                Case "CUENTA" To "CUENTA 99,99":
                    If Len(programa) > 7 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 8, pos - 8)
                        Cantidad = Mid(programa, pos + 1, Len(programa) - pos)
                        Call NroCuenta(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    Else
                        posicion = "1"
                        Cantidad = rs_Items!itemiclong
                        Call NroCuenta(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    End If
                 
                 Case "VOLCOD" To "VOLCOD 99,99":
                    If Len(programa) > 7 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 8, pos - 8)
                        Cantidad = Mid(programa, pos + 1, Len(programa) - pos)
                        Call nrovolcod(rs_Procesos!vol_cod, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    Else
                        posicion = "1"
                        Cantidad = rs_Items!itemiclong
                        Call nrovolcod(rs_Procesos!vol_cod, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    End If
                
                 Case "IMPORTECH" To "IMPORTECH Z":
                 Dim ceros As String
                    'Completa = False
                    If Len(programa) > 9 Then
                        posicion = Mid(programa, 11, 1)
                        ceros = Mid(programa, 12, 2)
                        Select Case UCase(posicion)
                            Case "D":
                                Call ImporteCH(rs_Procesos!Monto, rs_Procesos!dh, "D", rs_Items!itemiclong, cadena, ceros)
                            Case "H":
                                Call ImporteCH(rs_Procesos!Monto, rs_Procesos!dh, "H", rs_Items!itemiclong, cadena, ceros)
                            Case Else:
                                Call ImporteCH(rs_Procesos!Monto, rs_Procesos!dh, "", rs_Items!itemiclong, cadena, ceros)
                        End Select
                    Else
                        Call ImporteCH(rs_Procesos!Monto, rs_Procesos!dh, "", rs_Items!itemiclong, cadena, completa)
                    End If
                 Case "IMPORTE" To "IMPORTE Z":
                    If Len(programa) > 7 Then
                        posicion = Trim(Mid(programa, 9, 1))
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = True
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call Importe(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                    
                    
                 Case "IMPORTEN" To "IMPORTEN Z"::
                    If Len(programa) > 8 Then
                        posicion = Trim(Mid(programa, 10, 1))
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = False
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteN(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                    
                 Case "IMPESP":
                    completa = True
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteEsp(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                 Case "IMPESPS": 'CAS-14674 Agregado ver 1.36 - JAZ
                    completa = True
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteEspS(rs_Procesos!Monto, completa, rs_Items!itemiclong, rs_Procesos!dh, cadena)
                 Case "PORCC": 'CAS-14674 Agregado ver 1.36 - JAZ
                    Dim acount, cencosto As String
                    acount = Mid(CStr(rs_Procesos!Cuenta), 1, 10)
                    cencosto = Mid(CStr(rs_Procesos!Cuenta), 11, Len(rs_Procesos!Cuenta))
                    Call Porcentaje_CC(rs_Procesos!vol_cod, acount, cencosto, rs_Procesos!Monto, cadena)
                    'Call Porcentaje_CC(rs_Procesos!vol_cod, rs_Procesos!dh, completa, acount, cencosto, rs_Procesos!Monto, cadena)
                 
                 'sebastian stremel 03/10/2012
                 Case "IMPORTECTROCOSTOS":
                    Dim acount2, cencosto2 As String
                    acount2 = Mid(CStr(rs_Procesos!Cuenta), 1, 10)
                    cencosto2 = Mid(CStr(rs_Procesos!Cuenta), 11, Len(rs_Procesos!Cuenta))
                    Call ImporteCtroCostos(rs_Procesos!vol_cod, acount2, cencosto2, rs_Procesos!Monto, cadena, rs_Items!itemiclong, rs_Procesos!dh)
                 'hasta aca
                    
                 'sebastian stremel 29/01/2013
                 'Case "TOTALGRUPO(1,10)":
                 Case "TOTALGRUPO,0,0" To "TOTALGRUPO,99,99":
                    Dim acount3, cencosto3 As String
                    Dim Cuenta() As String
                    Dim longDesde As Integer
                    Dim longHasta As Integer
                    Cuenta = Split(programa, ",")
                    longDesde = Cuenta(1)
                    longHasta = Cuenta(2)
                    If (Not (IsNumeric(longDesde)) Or (Not IsNumeric(longHasta))) Then
                        Flog.writeline "El programa TOTALGRUPO esta mal configurado"
                        Exit Sub
                    End If
                    acount3 = Mid(CStr(rs_Procesos!Cuenta), longDesde, longHasta)
                    'cencosto2 = Mid(CStr(rs_Procesos!Cuenta), 11, Len(rs_Procesos!Cuenta))
                    Call ImporteGrupo(rs_Procesos!vol_cod, acount3, rs_Procesos!Monto, cadena, rs_Items!itemiclong, rs_Procesos!dh)
                 'hasta aca
                    
                 Case "IMPORTEF" To "IMPORTEF Z":
                    If Len(programa) > 8 Then
                        posicion = Trim(Mid(programa, 10, 1))
                        completa = (UCase(posicion) = "S")
                    Else
                        'completa = True
                        completa = False
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call Importe_Format(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena, "", "")
                    cadena = cadena & CStr(rs_Procesos!masinro)
                 Case "DIMPORTEF" To "DIMPORTEF Z":
                    If Len(programa) > 8 Then
                        posicion = Trim(Mid(programa, 10, 1))
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = True
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call Importe_Format_decimal(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena, "", SeparadorDecimales)
                    cadena = cadena & CStr(rs_Procesos!masinro)
                Case "IMPORTEPAR" To "IMPORTEPAR Z": 'JAZ Agregado 22-07-11 ver 1.33
                    If Len(programa) > 10 Then
                        posicion = Mid(programa, 12, 1)
                        completa = (UCase(posicion) = "C")
                    Else
                        completa = False
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteABS(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                    If rs_Procesos!dh Then
                    Else
                        cadena = "(" & cadena & ")"
                    End If
                    EsImporte = True
                    
                Case "IMPORTEABS" To "IMPORTEABS Z":
                    If Len(programa) > 10 Then
                        posicion = Mid(programa, 12, 1)
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = True
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteABS(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                
                Case "IMPORTEABSSD" To "IMPORTEABSSD Z":
                    If Len(programa) > 12 Then
                        posicion = Mid(programa, 14, 1)
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = False
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteABSSD(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                
                Case "IMPORTEABSP" To "IMPORTEABSP Z":
                    If Len(programa) > 11 Then
                        posicion = Mid(programa, 13, 1)
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = True
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteABS_2(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                Case "FECHA" To "FECHA YYYYYYYYYY"
                    If Len(programa) >= 7 Then
                        Formato = Mid(programa, 7, Len(programa) - 6)
                    Else
                        Formato = "DDMMYYYY"
                    End If
                
                    Select Case Formato
                    Case "YYYDDD":
                        Call Fecha1(rs_Procesos!vol_fec_asiento, cadena)
                    Case Else
                        Call Fecha_Estandar(rs_Procesos!vol_fec_asiento, Formato, True, rs_Items!itemiclong, cadena)
                    End Select
                Case "FECHAPROCVOL" To "FECHAPROCVOL YYYYYYYY"
                    If Len(programa) >= 14 Then
                        Formato = Mid(programa, 14, Len(programa) - 6)
                    Else
                        Formato = "DDMMYYYY"
                    End If
                
                    Select Case Formato
                    Case "YYYDDD":
                        Call Fecha1(rs_Procesos!vol_fec_proc, cadena)
                    Case "YYMM":
                        cadena = Format(rs_Procesos!vol_fec_proc, "YYMM")
                    Case Else
                        Call Fecha_Estandar(rs_Procesos!vol_fec_proc, Formato, True, rs_Items!itemiclong, cadena)
                    End Select
                Case "FECHALIQH" To "FECHALIQH YYYYYYYY"
                    If Len(programa) >= 11 Then
                        Formato = Mid(programa, 11, Len(programa) - 6)
                    Else
                        Formato = "DDMMYYYY"
                    End If
                
                    Select Case Formato
                    Case "YYYDDD":
                        Call Fecha1(rs_Periodo!pliqhasta, cadena)
                    Case Else
                        Call Fecha_Estandar(rs_Periodo!pliqhasta, Formato, True, rs_Items!itemiclong, cadena)
                    End Select
                Case "PROCESO":
                    Cantidad = CLng(rs_Items!itemiclong)
                    Call Leyenda(rs_Procesos!vol_desc, 1, CInt(Cantidad), True, rs_Items!itemiclong, cadena)
                Case "SIPROCESO": 'NG - 27/01/2016
                    Cantidad = CLng(Len(rs_Procesos!vol_desc))
                    cadena = rs_Procesos!vol_desc
                    Call LeyendaAux(rs_Procesos!vol_desc, 1, CInt(Cantidad), False, rs_Items!itemiclong, rs_Procesos!Linea, ListaLineas, cadena)
                
                Case "LEYENDA":
                    Cantidad = CLng(rs_Items!itemiclong)
                    Call Leyenda(rs_Procesos!descLinea, 1, CInt(Cantidad), True, rs_Items!itemiclong, cadena)
                Case "MODELO,LEYENDA":
                    Cantidad = CLng(rs_Items!itemiclong)
                    Call Leyenda1(rs_Procesos!masinro, rs_Procesos!Linea, rs_Procesos!descLinea, 1, rs_Items!itemiclong, True, rs_Items!itemiclong, cadena)
                Case "LINEA":
                    Call nroLinea(Nro, True, rs_Items!itemiclong, cadena)
                Case "AGRUPADOR":
                    Call nroLinea(NroL, True, rs_Items!itemiclong, cadena)
                Case "ASIENTO":
                    Call NroAsiento(Asinro, True, rs_Items!itemiclong, cadena)
                Case "PERIODO" To "PERIODO 99":
                    If Len(programa) > 8 Then
                        posicion = Mid(programa, 9, 2)
                        Call NroPeriodo(rs_Periodo!pliqmes, CLng(posicion), True, rs_Items!itemiclong, cadena)
                    Else
                        posicion = "0"
                        Call NroPeriodo(rs_Periodo!pliqmes, CLng(posicion), True, rs_Items!itemiclong, cadena)
                    End If
                Case "DEBEHABER" To "DEBEHABER ZZ,ZZ":
                    pos = CLng(InStr(1, programa, ","))
                    If dh = "" Then
                        debeCod = Mid(programa, 11, pos - 11)
                        haberCod = Mid(programa, pos + 1, Len(programa) - pos)
                    Else
                        If dh = "D" Then
                            debeCod = Mid(programa, 11, pos - 11)
                            haberCod = Mid(programa, 11, pos - 11)
                        End If
                        
                        If dh = "H" Then
                            debeCod = Mid(programa, pos + 1, Len(programa) - pos)
                            haberCod = Mid(programa, pos + 1, Len(programa) - pos)
                        
                        End If
                        
                    End If
                    Call debehaber(rs_Procesos!dh, debeCod, haberCod, True, rs_Items!itemiclong, cadena)
                    dh = ""
                Case "FECHAACTUAL" To "FECHAACTUAL YYYYYYYY"
                    If Len(programa) >= 13 Then
                        Formato = Mid(programa, 13, Len(programa) - 6)
                    Else
                        Formato = "DDMMYYYY"
                    End If
                
                    Select Case Formato
                    Case "YYYDDD":
                        Call Fecha1(Date, cadena)
                    Case Else
                        Call Fecha_Estandar(Date, Formato, True, rs_Items!itemiclong, cadena)
                    End Select
                Case "MODELO":
                    Cantidad = CLng(rs_Items!itemiclong)
                    Call Leyenda2(rs_Procesos!masinro, rs_Procesos!Linea, rs_Procesos!descLinea, 1, rs_Items!itemiclong, True, rs_Items!itemiclong, cadena)
                    
                Case "MODELOPERIODO":
                    Cantidad = CLng(rs_Items!itemiclong)
                    Call Leyenda3(rs_Procesos!masinro, rs_Procesos!Linea, rs_Procesos!descLinea, 1, rs_Items!itemiclong, True, rs_Items!itemiclong, rs_Periodo!pliqmes, rs_Periodo!pliqanio, cadena)
                 
'                ' LA - 17-01-2007 - Roche
'                Case "PROFIT":
'                    Aux_str = Mid(Trim(rs_Procesos!Cuenta), 10, 2)
'
'                    Select Case Aux_str
'                    Case "11":
'                        cadena = "0000110000"
'                    Case "45":
'                        cadena = "0000453000"
'                    Case Else
'                        cadena = String(10, " ")
'                    End Select
                Case "CC_PROFIT":  'FGZ - 06/02/2007
                    Cantidad = CLng(rs_Items!itemiclong)
                    Call CC_Profit(rs_Procesos!masinro, rs_Procesos!Linea, rs_Procesos!descLinea, rs_Procesos!Cuenta, 1, rs_Items!itemiclong, True, rs_Items!itemiclong, rs_Periodo!pliqmes, rs_Periodo!pliqanio, cadena)
                    '(rs_Procesos!masinro, rs_Procesos!Linea, rs_Procesos!desclinea, rs_Procesos!Cuenta, 1, rs_Items!itemiclong, True, rs_Items!itemiclong, rs_Periodo!pliqmes, rs_Periodo!pliqanio, cadena)
                    
                'MAF - 21/02/2007 - Para la union de paris
                Case "SICUENTA" To "SICUENTA 99,99,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ":
                    If Len(programa) > 9 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 10, pos - 10)
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Cantidad = Mid(programa, pos + 1, pos2 - pos - 1)
                        IgualA = Mid(programa, pos2 + 1, Len(programa) - pos2)
                        Call ComparaCTA(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), IgualA, rs_Procesos!vol_fec_proc, rs_Items!itemiclong, cadena)
                    Else
                        cadena = " ERROR "
                        Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & ". Son 3 parametros, posicion inicial, longitud y Comparacion."
                    End If
                 '------------JAZ - CAS-12398 - Sykes CR - Nuevo Item Exportación Asiento ------------
                 Case "SICUENTAS" To "SICUENTAS 99,99,ZZZZZZZZZZZZZZZZZZZZ,99,99":
                    If Len(programa) > 10 Then
                        posicion = Trim(Mid(programa, 10, Len(programa)))
                        ArrC = Split(posicion, ",")
                        Call ComparaCADS(rs_Procesos!Cuenta, ArrC(), cadena)
                    Else
                        cadena = " ERROR "
                        Flog.writeline "No hay parámetros en el Item SICUENTAS"
                    End If
                                 
                 Case "CTAREPLACE" To "CTAREPLACE ZZZZZZZZZZ,ZZZZZZZZZZ,99,99":
                    If Len(programa) > 11 Then
                        'Miro que tenga parametros
                        If InStr(programa, ",") <> 0 Then
                            'guardo los parametros en arreglo
                            ArrPar = Split(Mid(rs_Items!itemicprog, 12, Len(programa)), ",", -1)
                            If UBound(ArrPar) = 3 Then
                                
                                'Controlo que el parametro DesdeMostrar sea numerico
                                If IsNumeric(ArrPar(2)) Then
                                        'Controlo que el parametro sea numerico
                                        If IsNumeric(ArrPar(3)) Then
                                            Call CuentaReemplaza(rs_Procesos!Cuenta, ArrPar(0), ArrPar(1), ArrPar(2), ArrPar(3), rs_Items!itemiclong, cadena)
                                        Else
                                            Flog.writeline Espacios(Tabulador * 2) & "En el Item " & rs_Items!itemicnro & " el parametro CantMostrar debe ser numerico."
                                        End If
                                Else
                                    Flog.writeline Espacios(Tabulador * 2) & "En el Item " & rs_Items!itemicnro & " el parametro DesdeMostrar debe ser numerico."
                                End If
                                
                            Else
                                Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & ". Son 2 parametros, cadena a buscar, cadena reemplazar, desdemostrar, cantmostrar."
                            End If
                        Else
                            Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & ". Son 2 parametros, cadena a buscar, cadena reemplazar, desdemostrar, cantmostrar."
                        End If
                    Else
                        Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & ". Son 2 parametros, cadena a buscar, cadena reemplazar, desdemostrar, cantmostrar."
                    End If
               
                '/////caso especial de CCU (Cia. Industrial Cervecera) - start //////////////
                Case "SICTA" To "SICTA 99,99,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ,99,99,99,99,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ":
                    If Len(programa) > 6 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 7, pos - 7)
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Cantidad = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        IgualA = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        CambiaPor = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        LongIgualAdesde = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        LongIgualAhasta = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Posicion2 = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Posicion3 = Mid(programa, pos + 1, pos2 - pos - 1)
                        Relleno = Mid(programa, pos2 + 1, Len(programa) - pos2)
                        Call ComparaCTAZ(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), IgualA, CambiaPor, CLng(LongIgualAdesde), CLng(LongIgualAhasta), CLng(Posicion2), CLng(Posicion3), Relleno, rs_Items!itemiclong, cadena)
                    Else
                        cadena = " ERROR "
                        Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & ". Son 7 parametros, posición inicial, longitud, Comparación, reemplazo, longitud de la cadena, posición 2, longitud desde la posición 2."
                    End If
                
                Case "IMPORTEREP" To "IMPORTEREP Z":
                    If Len(programa) > 10 Then
                        posicion = Mid(programa, 12, 1)
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = True
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteABS_4(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, rs_Procesos!Cuenta, cadena)
                    
                '/////caso especial de CCU (Cia. Industrial Cervecera) - end //////////////
                
                Case "ASIENTOZ":
                    Call NroAsientoZ(Asinro, True, rs_Items!itemiclong, cadena)
                Case "MODELO_NRO"
                    Call Modelo_Nro(rs_Procesos!masinro, rs_Procesos!Linea, rs_Procesos!descLinea, 1, rs_Items!itemiclong, True, rs_Items!itemiclong, cadena)
                'LED - 04/07/2012
                Case "ENTER"
                    cadena = Enter
                    vinoEnter = True
                'sebastian stremel 27/02/2013
                Case "PRIMERCAMPO,AAAAA,AAAAA" To "PRIMERCAMPO,ZZZZZ,ZZZZZ"
                    Dim texto1 As String
                    Dim texto2 As String
                    Dim cadena1
                    cadena1 = Split(programa, ",")
                    texto1 = cadena1(1)
                    texto2 = cadena1(2)
                    If Len(texto1) > 5 Then
                        Flog.writeline " ERROR, la longitud del parametro texto 1 es mayor a 5"
                    Else
                        If Len(texto2) > 5 Then
                            Flog.writeline " ERROR, la longitud del parametro texto 2 es mayor a 5"
                        Else
                            Call primerCampo(texto1, texto2, cadena)
                        End If
                    End If
                Case "CUENTA_DH_VAR 00,00-00 " To "CUENTA_DH_VAR 99,99-99":
                 'EAM (v1.57)
                    If (rs_Procesos!dh = -1) Then
                        pos = CLng(InStr(1, programa, " "))
                        posicion = Mid(programa, pos, CLng(InStr(pos, programa, ",")) - pos)
                        Cantidad = Mid(programa, CLng(InStr(pos, programa, ",")) + 1, CLng(InStr(pos, programa, "-")) - (CLng(InStr(pos, programa, ",")) + 1))
                        
                        Call NroCuentaVariable(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    Else
                        pos = CLng(InStr(1, programa, " "))
                        posicion = Mid(programa, pos, CLng(InStr(pos, programa, ",")) - pos)
                        Cantidad = Mid(programa, CLng(InStr(pos, programa, "-")) + 1, (CLng(Len(programa)) - CLng(InStr(pos, programa, "-"))))
                        
                        Call NroCuentaVariable(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    End If
                    
                'RAD(v1.87)
                Case "CUENTA_DH_VAR2 00,00-00,00 " To "CUENTA_DH_VAR2 99,99-99,99":
                    'verifico si es variable
                    StrSql = "SELECT linaD_H FROM mod_linea "
                    StrSql = StrSql & "WHERE masinro = " & rs_Procesos!masinro
                    StrSql = StrSql & " AND mod_linea.linaorden = " & rs_Procesos!Linea
                    OpenRecordset StrSql, rs_ItemsAux
                    If Not rs_ItemsAux.EOF Then
                        If rs_ItemsAux!linaD_H > 1 Then ' Variable o variable invertida
                            'busco la cantidad de digitos
                            Dim cant
                            posicion = 1
                            cant = Split(programa, ",")
                            If UBound(cant) > 1 Then
                                Cantidad = cant(2)
                                Call NroCuentaVariable(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                            Else
                                Flog.writeline "La configuracion del programa es incorrecta, falta un parametro"
                                Exit Sub
                            End If
                            
                        Else
                            'no es variable, aplica la logica anterior
                            If (rs_Procesos!dh = -1) Then
                                pos = CLng(InStr(1, programa, " "))
                                posicion = Mid(programa, pos, CLng(InStr(pos, programa, ",")) - pos)
                                Cantidad = Mid(programa, CLng(InStr(pos, programa, ",")) + 1, CLng(InStr(pos, programa, "-")) - (CLng(InStr(pos, programa, ",")) + 1))
                                
                                Call NroCuentaVariable(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                            Else
                                pos = CLng(InStr(1, programa, " "))
                                posicion = Mid(programa, pos, CLng(InStr(pos, programa, ",")) - pos)
                                Cantidad = Mid(programa, CLng(InStr(pos, programa, "-")) + 1, (CLng(Len(programa)) - CLng(InStr(pos, programa, "-"))))
                                
                                Call NroCuentaVariable(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                            End If
                        End If
                    End If
                    rs_ItemsAux.Close
                    
                'RAD 1.89
                Case "CUENTATEXT 00,00 " To "CUENTATEXT 99,99":
                    Dim cuenta_original As String
                     Call CuentaL(rs_Procesos!Linea, rs_Procesos!masinro, cuenta_original)
                     pos = Mid(programa, 12, (CLng(InStr(1, programa, ",")) - 12))
                     Cantidad = Mid(programa, (CLng(InStr(1, programa, ",")) + 1), (CLng(Len(programa)) - CLng(InStr(pos, programa, "-"))))
                     cadena = Mid(cuenta_original, pos, Cantidad)
                    
                'LED - 04/11/2014 - v1.83
                Case "CUENTA_DH_FIJ_VAR 00-00,00-00 " To "CUENTA_DH_FIJ_VAR 99-99,99-99":
                    If (rs_Procesos!dh = -1) Then
                        pos = CLng(InStr(1, programa, " "))
                        posicion = Mid(programa, pos, CLng(InStr(pos, programa, "-")) - pos)
                        Cantidad = Mid(programa, CLng(InStr(pos, programa, "-")) + 1, CLng(InStr(programa, ",") - 1) - CLng(InStr(pos, programa, "-")))
                        
                        Call NroCuentaVariable(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    Else
                        'POS = CLng(InStr(1, Programa, " "))
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, pos + 1, CLng(InStr(pos, programa, "-") - 1) - pos)
                        Cantidad = Mid(programa, CLng(InStr(pos, programa, "-")) + 1, (CLng(Len(programa)) - CLng(InStr(pos, programa, "-"))))
                        
                        Call NroCuentaVariable(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    End If
                'Fin LED - 04/11/2014 - v1.83
                
                'LED 16/01/2014 - CAS-23341 - ZOETIS COLOMBIA - Item exportación asiento contable
                Case "FINCUENTA":
                    Call finCuentaZoetis(rs_Procesos!Cuenta, True, rs_Items!itemiclong, rs_Procesos!vol_fec_asiento, cadena)
                'Fin - 'LED 16/01/2014 - CAS-23341 - ZOETIS COLOMBIA - Item exportación asiento contable
                
    
                'LED 19 / 6 / 2013 - 19041
                Case "CUENTA2F" To "CUENTA2F 99,99":
                    If Len(programa) > 8 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 9, pos - 9)
                        Cantidad = Mid(programa, pos + 1, Len(programa) - pos)
                        Call NroCuenta(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), False, rs_Items!itemiclong, cadena)
                    Else
                        posicion = "1"
                        Cantidad = rs_Items!itemiclong
                        Call NroCuenta(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), False, rs_Items!itemiclong, cadena)
                    End If
                'Fin - LED 19/06/2013 - 19041
                
                'Inicio
                Case "CUENTAF" To "CUENTAF 99,99,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ,99,99,99,99,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ":
                    If Len(programa) > 8 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 9, pos - 9)
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Cantidad = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        IgualA = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        CambiaPor = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        LongIgualAdesde = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        LongIgualAhasta = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Posicion2 = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Posicion3 = Mid(programa, pos + 1, pos2 - pos - 1)
                        Relleno = Mid(programa, pos2 + 1, Len(programa) - pos2)
                        Call ComparaCuentaF(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), IgualA, CambiaPor, CLng(LongIgualAdesde), CLng(LongIgualAhasta), CLng(Posicion2), CLng(Posicion3), Relleno, rs_Items!itemiclong, cadena)
                    Else
                        cadena = " ERROR "
                        Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & ". Son 7 parametros, posición inicial, longitud, Comparación, reemplazo, longitud de la cadena, posición 2, longitud desde la posición 2."
                    End If
               ' Fin
                'Inicio
                Case "CUENTAFF" To "CUENTAFF 99,99,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ,99,99,99,99,ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ,Z":
                    If Len(programa) > 8 Then
                        pos = CLng(InStr(1, programa, ","))
                        posicion = Mid(programa, 9, pos - 9)
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Cantidad = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        IgualA = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        CambiaPor = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        LongIgualAdesde = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        LongIgualAhasta = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Posicion2 = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Posicion3 = Mid(programa, pos + 1, pos2 - pos - 1)
                        pos = pos2
                        pos2 = CLng(InStr(pos + 1, programa, ","))
                        Relleno = Mid(programa, pos + 1, pos2 - pos - 1)
                        caracter = Mid(programa, pos2 + 1, Len(programa) - pos2)
                                              
                        
                        Call ComparaCuentaFF(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), IgualA, CambiaPor, CLng(LongIgualAdesde), CLng(LongIgualAhasta), CLng(Posicion2), CLng(Posicion3), Relleno, rs_Items!itemiclong, caracter, cadena)
                    Else
                        cadena = " ERROR "
                        Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & ". Son 7 parametros, posición inicial, longitud, Comparación, reemplazo, longitud de la cadena, posición 2, longitud desde la posición 2."
                    End If
               ' Fin
                Case "LINEAS" To "LINEAS Z"
                    Call lineas(programa, Nro, rs_Items!itemiclong, cadena)

                Case "IMPORTECTA" To "IMPORTECTA Z"
                    If Len(programa) > 10 Then
                        posicion = Trim(Mid(programa, 12, 1))
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = False
                    End If
                    
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call IMPORTECTA(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                    
                Case "IMPORTEABSSR" To "IMPORTEABSSR Z":
                    If Len(programa) > 12 Then
                        posicion = Mid(programa, 14, 1)
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = False
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteABSSR(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena)
                    
                Case "IMPORTEABSCR3" To "IMPORTEABSCR3 Z":
                    If Len(programa) > 13 Then
                        posicion = Mid(programa, 15, 1)
                        completa = (UCase(posicion) = "S")
                    Else
                        completa = False
                    End If
                    If EsUltimoRegistroItem(rs_Items) Then
                        EsUltimoItem = True
                    End If
                    Call ImporteABSCR3(rs_Procesos!Monto, rs_Procesos!dh, completa, rs_Items!itemiclong, cadena, dh)
                    
                '//// caso especial de Bco. Industrial Azul - BIA - start //////
                Case "COMPROBANTE":
                    Sucursal = Mid(rs_Procesos!Cuenta, 1, 3)
                    If UCase(Sucursal_Ant) <> UCase(Sucursal) Then
                        Sucursal_Ant = Sucursal
                        NroSucursal = NroSucursal + 1
                        NroSucursalInterno = 1
                    Else
                        NroSucursalInterno = NroSucursalInterno + 1
                    End If
                    
                    If IsNumeric(Sucursal) Then
                        TotSuc = TotSuc + CLng(Sucursal)
                    End If
                    
                    posicion = "1"
                    Cantidad = rs_Items!itemiclong
                    'licho - 16/12/2014 - NroSucursal * NroSucursalInterno
                    
                    'MDF-----------
                    
                    'Call Comprobante(CStr(NroSucursalInterno), CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    Call Comprobante(CStr(NroSucursal), CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                    
                    '16/06/2015 - Carmen Quintero
                    'TotComprobante = TotComprobante + CLng(NroSucursal)
                    TotComprobante = TotComprobante + CLng(NroSucursalInterno)
                    'FIN
                
                Case "LINEA2":
                    Call nroLinea(NroSucursalInterno, True, rs_Items!itemiclong, cadena)
                    
                Case ""
                '//// caso especial de Bco. Industrial Azul - BIA - end //////
                Case "SUBCTA_LEGAJO" To "SUBCTA_LEGAJO 99,99,0000000000,##########":
                    Dim param
                    param = Split(programa, ",")
                    pos = Right(param(0), 2)
                    cant = param(1)
                    Call subcta_legajo(rs_Procesos!Cuenta, pos, cant, False, rs_Items!itemiclong, param(2), param(3), rs_Procesos!descLinea, cadena)
                    
                Case "SUBCTA_CC" To "SUBCTA_CC 99,99,0000,####":
                    param = Split(programa, ",")
                    pos = Right(param(0), 2)
                    cant = param(1)
                    Call subcta_cc(rs_Procesos!Cuenta, pos, cant, False, rs_Items!itemiclong, param(2), param(3), rs_Procesos!descLinea, cadena)
                
                Case "NRODOC" To "NRODOC 99,99":
                    param = Split(programa, ",")
                    pos = Right(param(0), 2)
                    cant = param(1)
                    Call nro_doc(rs_Procesos!Cuenta, pos, cant, False, rs_Items!itemiclong, cadena)
               
                Case "LEG_CC" To "LEG_CC 99,99,99,99":
                    param = Split(programa, ",")
                    Dim posLeg As Integer
                    Dim cantLeg As Integer
                    Dim posCC As Integer
                    Dim cantCC As Integer
                    posLeg = Right(param(0), 2)
                    cantLeg = param(1)
                    posCC = Right(param(2), 2)
                    cantCC = param(3)
                    Call leg_cc(rs_Procesos!Cuenta, rs_Procesos!Linea, rs_Procesos!masinro, posLeg, cantLeg, posCC, cantCC, False, rs_Items!itemiclong, cadena)
                
                Case Else
                    cadena = " ERROR "
                    Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
                End Select
            End If
                
            'If Mid(cadena, 1, 2) <> "RR" Or primero Then 'Comentado versión 1.31
            If primero Then
                If Aux_Linea = "" Then
                    Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
                Else
                    If cadena <> Enter And vinoEnter Then
                        Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
                        vinoEnter = False
                    Else
                        Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
                    End If
                End If
            Else
                'Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)   'Comentado versión 1.31
                If EsImporte Then 'JAZ Agregado 22-07-11 ver 1.33
                    Aux_Linea = Aux_Linea & separadorCampos & cadena 'JAZ Agregado 22-07-11 ver 1.33
                    EsImporte = False
                Else
                    If cadena <> Enter And vinoEnter Then
                        Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
                        vinoEnter = False
                    Else
                        Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
                    End If
                
                    'Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong) 'Agregado ver. 1.31
                End If
            End If
            primero = False
            'End If seba 12/10/2012
   
            rs_Items.MoveNext
        Loop
            
        ' ------------------------------------------------------------------------
        'Escribo en el archivo de texto
        'Aux_Relleno = Space(256 - Len(Aux_Linea))
        'FGZ - 07/02/2012 - le agregué esta condicion para CARDIF
        If porcc Or porcc1 Then
            If Trim(Aux_Linea) <> "" Then
                'fExport.writeline Aux_Linea '& Aux_Relleno
                fAuxiliarDetalle.writeline Aux_Linea
            End If
        Else
            fAuxiliarDetalle.writeline Aux_Linea
            'fExport.writeline Aux_Linea '& Aux_Relleno
        End If

        

        
        'fExport.writeline Aux_Linea '& Aux_Relleno
        'fAuxiliarDetalle.writeline Aux_Linea
        primero = True
        'Actualizo el progreso del Proceso
        Progreso = Progreso + IncPorc
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    '10/02/2012 Gonzalez N.|Formateo Aux_linea
    Aux_Linea = ""
                
    'Siguiente proceso
    rs_Procesos.MoveNext
Loop

rs_Procesos.MoveFirst
'------------------------------------------------------
'seba 28/12/2015 -se pasa el encabezado a otro modulo
'ByVal rs_Procesos As Recordset, ByVal rs_Periodo As Recordset
Call encabezado(rs_Items, cadena, rs_Procesos, rs_Periodo)
'hasta aca
'------------------------------------------------------
'------------------------------------------------------
'aca codigo viejo

'hasta aca
'------------------------------------------------------
'licho desde
''------------------------------------------------------------------------
'' Agregado ver. 1.36 - JAZ - CAS-14674-------------
'' Genero los totales para realizar el PORCC
''------------------------------------------------------------------------
'' Realizo las consultas para ver si existe el Item en el Detalle o Pie
'
'' Por Pie en rs_ItemsPie
'StrSql = "SELECT * FROM confitemicpie "
'StrSql = StrSql & " INNER JOIN itemintcont ON confitemicpie.itemicnro = itemintcont.itemicnro "
'If ModSalidaAsiento <> 0 Then
'    StrSql = StrSql & " AND confitemicpie.moditenro = " & ModSalidaAsiento
'End If
'StrSql = StrSql & " WHERE itemicprog = 'PORCC'"
'StrSql = StrSql & " ORDER BY confitemicpie.confitemicorden "
'OpenRecordset StrSql, rs_ItemsPie
'
''Por Detalle en rs_Items
'StrSql = "SELECT * FROM confitemic "
'StrSql = StrSql & " INNER JOIN itemintcont ON confitemic.itemicnro = itemintcont.itemicnro "
'If ModSalidaAsiento <> 0 Then
'    StrSql = StrSql & " AND confitemic.moditenro = " & ModSalidaAsiento
'End If
'StrSql = StrSql & " WHERE itemicprog = 'PORCC'"
'StrSql = StrSql & " ORDER BY confitemic.confitemicorden "
'OpenRecordset StrSql, rs_Items
'
'' Verifico Previamente si se realiza por Porcc
'If Not rs_Items.EOF Or Not rs_ItemsPie.EOF Then
'    porcc = True
'End If
'' Si existe el Item, Genero los Totales
'If porcc Then
'    Do While Not rs_Procesos.EOF
'        Call AlmCuentaCC(rs_Procesos!cuenta, rs_Procesos!Monto, rs_Procesos!linea, rs_Procesos!desclinea, rs_Procesos!dh)
'        rs_Procesos.MoveNext
'    Loop
'End If
'licho hasta


'------------------------------------------------------------------------
' Genero el pie de la exportacion
'------------------------------------------------------------------------
Flog.writeline Espacios(Tabulador * 1) & "-------------------------------------"
Flog.writeline Espacios(Tabulador * 1) & "Exportando datos del pie del proceso de volcado "
Flog.writeline

Cantidad_Warnings = 0
Nro = Nro + 1 'Contador de Lineas
       

StrSql = "SELECT * FROM confitemicpie "
StrSql = StrSql & " INNER JOIN itemintcont ON confitemicpie.itemicnro = itemintcont.itemicnro "
If ModSalidaAsiento <> 0 Then
    StrSql = StrSql & " AND confitemicpie.moditenro = " & ModSalidaAsiento
End If
StrSql = StrSql & " ORDER BY confitemicpie.confitemicorden "
OpenRecordset StrSql, rs_Items
Aux_Linea = ""
If porcc Then
    rs_Procesos.MoveFirst
    Aux_Linea = ""
    primero = True
    EsUltimoLineaCuenta = False
    EsUltimoLineaCuenta1 = False
    Do While Not rs_Procesos.EOF
    
        If EsUltimoRegistroCuenta(rs_Procesos, rs_Procesos!Linea) Then
            EsUltimoLineaCuenta = True
            EsUltimoLineaCuenta1 = True
        End If
        If Not rs_Items.EOF Then 'sebastian stremel 23/10/2012
            rs_Items.MoveFirst
        End If
        cadena = ""
        Aux_Linea = ""
        primero = True
      'If Len(rs_Procesos!Cuenta) > 11 Then 'Imprime sólo cuentas con CC -> Tamaño de Cuenta = 10 caracteres - Agregado ver 1.37
          Do While Not rs_Items.EOF
            cadena = ""
            If rs_Items!itemicfijo Then
                If rs_Items!itemicvalorfijo = "" Then
                    cadena = String(256, " ")
                Else
                    cadena = rs_Items!itemicvalorfijo
                End If
            Else
                programa = UCase(rs_Items!itemicprog)
                Select Case programa
                Case "PIESAP":
                    cadena = "*****;;0.00"
                
                Case "ESPACIOS":
                    cadena = String(rs_Items!itemiclong, " ")
                
                Case "PERIODO" To "PERIODO 99":
                    If Len(programa) > 8 Then
                        posicion = Mid(programa, 9, 2)
                        Call NroPeriodo(rs_Periodo!pliqmes, CLng(posicion), True, rs_Items!itemiclong, cadena)
                    Else
                        posicion = "0"
                        Call NroPeriodo(rs_Periodo!pliqmes, CLng(posicion), True, rs_Items!itemiclong, cadena)
                    End If
                
                Case "FECHAACTUAL" To "FECHAACTUAL YYYYYYYY"
                    If Len(programa) >= 13 Then
                        Formato = Mid(programa, 13, Len(programa) - 6)
                    Else
                        Formato = "DDMMYYYY"
                    End If
                
                    Select Case Formato
                    Case "YYYDDD":
                        Call Fecha1(Date, cadena)
                    Case Else
                        Call Fecha_Estandar(Date, Formato, True, rs_Items!itemiclong, cadena)
                    End Select
                Case "CUENTZ" To "CUENTZ 99,99": ' Agregado Ver 1.37
                        If Len(programa) > 7 Then
                            pos = CLng(InStr(1, programa, ","))
                            posicion = Mid(programa, 8, pos - 8)
                            Cantidad = Mid(programa, pos + 1, Len(programa) - pos)
                            Call NroCuenta_1(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                        Else
                            posicion = "1"
                            Cantidad = "10"
                            Call NroCuenta_1(rs_Procesos!Cuenta, CLng(posicion), CLng(Cantidad), True, rs_Items!itemiclong, cadena)
                        End If
        
                Case "IMPORTETOTAL":
                    Call ImporteTotal(True, rs_Items!itemiclong, cadena)
                'JAZ Agregado 22-07-11 ver 1.33
                Case "IMPORTETOTALPAR" To "IMPORTETOTALPAR C":
                    If Len(programa) > 15 Then
                      posicion = Mid(programa, 15, 1)
                      completa = (UCase(posicion) = "C")
                    Else
                      completa = False
                    End If
                    Call ImporteTotal(completa, rs_Items!itemiclong, cadena)
                    If Int(cadena) < 0 Then
                        cadena = Replace(cadena, "-", "")
                        cadena = "(" & cadena & ")"
                    End If
                    EsImporte = True
                Case "IMPORTETOTALDH D", "IMPORTETOTALDH H":
                    If Len(programa) > 15 Then
                        If Mid(programa, 16, 1) = "D" Then
                            Call ImporteTotalDH(True, rs_Items!itemiclong, True, cadena)
                        Else
                            Call ImporteTotalDH(True, rs_Items!itemiclong, False, cadena)
                        End If
                    Else
                        cadena = " ERROR "
                        Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & " o esta mal definido. Se debe indicar si el total es D (debe) o H (haber)."
                    End If
                Case "PORCC": 'CAS-14674 Agregado ver 1.36 - JAZ
                    'Dim acount, cencosto As String
                    acount = Mid(CStr(rs_Procesos!Cuenta), 1, 10)
                    cencosto = Mid(CStr(rs_Procesos!Cuenta), 11, Len(rs_Procesos!Cuenta))
                    Call Porcentaje_CC(rs_Procesos!vol_cod, acount, cencosto, rs_Procesos!Monto, cadena)
                    'Call Porcentaje_CC(rs_Procesos!vol_cod, rs_Procesos!dh, completa, acount, cencosto, rs_Procesos!Monto, cadena)
                
                'sebastian stremel 03/10/2012
                Case "IMPORTECTROCOSTOS":
                    'Dim acount2, cencosto2 As String
                    acount2 = Mid(CStr(rs_Procesos!Cuenta), 1, 10)
                    cencosto2 = Mid(CStr(rs_Procesos!Cuenta), 11, Len(rs_Procesos!Cuenta))
                    Call ImporteCtroCostos(rs_Procesos!vol_cod, acount2, cencosto2, rs_Procesos!Monto, cadena, rs_Items!itemiclong, rs_Procesos!dh)
                'hasta aca
                
                Case "TOTALREG" To "TOTALREG Y":
                    'Call totalRegistros(Nro - 1, True, rs_Items!itemiclong, cadena)
                    If Len(programa) > 8 Then
                        posicion = Mid(programa, 10, 1)
                        completa = (UCase(posicion) = "S")
                    Else
                       completa = True
                    End If
                    Call totalRegistrosCompletar(Nro - 1, completa, rs_Items!itemiclong, cadena)
                Case "MODELO_NRO"
                    'rs_Procesos.MoveFirst Modificado JAZ ver 1.36 - CAS-14674
                    Call Modelo_Nro(rs_Procesos!masinro, rs_Procesos!Linea, rs_Procesos!descLinea, 1, rs_Items!itemiclong, True, rs_Items!itemiclong, cadena)
                'LED - 04/07/2012
                Case "ENTER"
                    cadena = Enter
                    vinoEnter = True
                'FGZ - 01/08/2013 -------------------------
                Case "TAB 1" To "TAB 9":
                    If Len(programa) > 4 Then
                        Cantidad = Mid(programa, 5, 1)
                    Else
                        Cantidad = "1"
                    End If
                    cadena = String(CLng(Cantidad), Chr(9))
                'FGZ - 01/08/2013 -------------------------
                'FGZ - 13/09/2012 ------------
                Case "TOTALDEBEHABER A,A" To "TOTALDEBEHABER Z,Z":
                    Fecha = Trim(Mid(programa, 16, 1))
                    completa = (UCase(Mid(programa, 18, 1)) = "S")
                    Select Case Fecha
                        Case "D":
                            'Call ImporteTotalDebeHaber(True, Completa, rs_Items!itemiclong, cadena)
                            Call ImporteTotalDebeHaber(True, completa, rs_Items!itemiclong, cadena, nroliq, ProcVol, Asinro)
                        Case "H":
                            'Call ImporteTotalDebeHaber(False, Completa, rs_Items!itemiclong, cadena)
                            Call ImporteTotalDebeHaber(False, completa, rs_Items!itemiclong, cadena, nroliq, ProcVol, Asinro)
                    End Select
                
                
                '///// casos especiales Bco. Industrial Azul - BIA - start ////////
                Case "TOTALSUC":
                    Call totalRegistros(TotSuc, True, rs_Items!itemiclong, cadena)
                    
                Case "TOTALCOMPROB":
                    Call totalRegistros(TotComprobante, True, rs_Items!itemiclong, cadena)
        
                '///// casos especiales Bco. Industrial Azul - BIA - end ////////
                    
                'LED - CAS-30304 - BANCO INDUSTRIAL - CUSTOM ITEM SUMATORIA DEBE Y HABER
                Case "IMPORTETOTALDH2 A,A" To "IMPORTETOTALDH2 Z,Z"
                    If Len(programa) > 16 Then
                        'ultimo parametro es "S" completa con 0's
                        If Mid(programa, 19, 1) = "S" Then
                            'primer parametro es "S" usa separador decimal
                            If Mid(programa, 17, 1) = "S" Then
                                Call ImporteTotalDH2(True, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro, True)
                            Else
                                Call ImporteTotalDH2(True, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro, False)
                            End If
                            'Call ImporteTotalDH(True, rs_Items!itemiclong, True, cadena)
                        Else
                            'primer parametro es "S" usa separador decimal
                            If Mid(programa, 17, 1) = "S" Then
                                Call ImporteTotalDH2(False, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro, True)
                            Else
                                Call ImporteTotalDH2(False, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro, False)
                            End If
                        End If
                    Else
                        cadena = " ERROR "
                        Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & " o esta mal definido. Se debe indicar si el total es D (debe) o H (haber)."
                    End If
                'Fin - LED - CAS-30304 - BANCO INDUSTRIAL - CUSTOM ITEM SUMATORIA DEBE Y HABER
                
                Case "IMPORTETOTALHABERN A,A" To "IMPORTETOTALHABERN Z,Z"
                    Call Validacion(programa, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro)
                   
                Case Else
                    cadena = " ERROR "
                    Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
                End Select
            End If
                
            'If Mid(cadena, 1, 2) <> "RR" Or primero Then 'Comentado versión 1.31
            If primero Then
                If Aux_Linea = "" Then
                    Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
                Else
                    If cadena <> Enter And vinoEnter Then
                        Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
                        vinoEnter = False
                    Else
                        Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
                    End If
                    
                    'Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)

                End If
            Else
                If EsImporte Then 'JAZ Agregado 22-07-11 ver 1.33
                   Aux_Linea = Aux_Linea & separadorCampos & cadena 'JAZ Agregado 22-07-11 ver 1.33
                    EsImporte = False
                Else 'Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)   'Comentado versión 1.31
                    If cadena <> Enter And vinoEnter Then
                        Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
                        vinoEnter = False
                    Else
                        Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
                    End If
                    
                    'Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong) 'Agregado ver. 1.31
                End If
            End If
            primero = False
            
            rs_Items.MoveNext
          Loop
      'End If 'de Cuenta con CC - Agregado ver 1.37
      'Escribo en el archivo de texto
      'Aux_Relleno = Space(256 - Len(Aux_Linea))
      If Trim(Aux_Linea) <> "" Then
        'fExport.writeline Aux_Linea '& Aux_Relleno
        fAuxiliarPie.writeline Aux_Linea
      End If
      rs_Procesos.MoveNext
      If Not primero Then
            Progreso = Progreso + 1
            TiempoAcumulado = GetTickCount
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                     ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                     "' WHERE bpronro = " & NroProcesoBatch
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
      End If
    Loop
Else
     Do While Not rs_Items.EOF
        cadena = ""
        If rs_Items!itemicfijo Then
            If rs_Items!itemicvalorfijo = "" Then
                cadena = String(256, " ")
            Else
                cadena = rs_Items!itemicvalorfijo
            End If
        Else
            programa = UCase(rs_Items!itemicprog)
            Select Case programa
            Case "PIESAP":
                cadena = "*****;;0.00"
            
            Case "ESPACIOS":
                cadena = String(rs_Items!itemiclong, " ")
            
            Case "PERIODO" To "PERIODO 99":
                If Len(programa) > 8 Then
                    posicion = Mid(programa, 9, 2)
                    Call NroPeriodo(rs_Periodo!pliqmes, CLng(posicion), True, rs_Items!itemiclong, cadena)
                Else
                    posicion = "0"
                    Call NroPeriodo(rs_Periodo!pliqmes, CLng(posicion), True, rs_Items!itemiclong, cadena)
                End If
            
            Case "FECHAACTUAL" To "FECHAACTUAL YYYYYYYY"
                If Len(programa) >= 13 Then
                    Formato = Mid(programa, 13, Len(programa) - 6)
                Else
                    Formato = "DDMMYYYY"
                End If
            
                Select Case Formato
                Case "YYYDDD":
                    Call Fecha1(Date, cadena)
                Case Else
                    Call Fecha_Estandar(Date, Formato, True, rs_Items!itemiclong, cadena)
                End Select
    
            Case "IMPORTETOTAL":
                Call ImporteTotal(True, rs_Items!itemiclong, cadena)
            'JAZ Agregado 22-07-11 ver 1.33
            Case "IMPORTETOTALPAR" To "IMPORTETOTALPAR C":
                If Len(programa) > 15 Then
                  posicion = Mid(programa, 15, 1)
                  completa = (UCase(posicion) = "C")
                Else
                  completa = False
                End If
                Call ImporteTotal(completa, rs_Items!itemiclong, cadena)
                If Int(cadena) < 0 Then
                    cadena = Replace(cadena, "-", "")
                    cadena = "(" & cadena & ")"
                End If
                EsImporte = True
            Case "IMPORTETOTALDH D", "IMPORTETOTALDH H":
                If Len(programa) > 15 Then
                    If Mid(programa, 16, 1) = "D" Then
                        Call ImporteTotalDH(True, rs_Items!itemiclong, True, cadena)
                    Else
                        Call ImporteTotalDH(False, rs_Items!itemiclong, False, cadena)
                    End If
                Else
                    cadena = " ERROR "
                    Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & " o esta mal definido. Se debe indicar si el total es D (debe) o H (haber)."
                End If
                
            Case "TOTALREG" To "TOTALREG Y":
                'Call totalRegistros(Nro - 1, True, rs_Items!itemiclong, cadena)
                 If Len(programa) > 8 Then
                     posicion = Mid(programa, 10, 1)
                     completa = (UCase(posicion) = "S")
                 Else
                    completa = True
                 End If
                 Call totalRegistrosCompletar(Nro - 1, completa, rs_Items!itemiclong, cadena)
                 
            Case "MODELO_NRO"
                'rs_Procesos.MoveFirst Modificado JAZ ver 1.36 - CAS-14674
                Call Modelo_Nro(rs_Procesos!masinro, rs_Procesos!Linea, rs_Procesos!descLinea, 1, rs_Items!itemiclong, True, rs_Items!itemiclong, cadena)
                            'LED - 04/07/2012
            Case "ENTER"
                 cadena = Enter
                 vinoEnter = True
                 
            'LED - CAS-30304 - BANCO INDUSTRIAL - CUSTOM ITEM SUMATORIA DEBE Y HABER
            Case "IMPORTETOTALDH2 A,A" To "IMPORTETOTALDH2 Z,Z"
                If Len(programa) > 16 Then
                    'ultimo parametro es "S" completa con 0's
                    If Mid(programa, 19, 1) = "S" Then
                        'primer parametro es "S" usa separador decimal
                        If Mid(programa, 17, 1) = "S" Then
                            Call ImporteTotalDH2(True, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro, True)
                        Else
                            Call ImporteTotalDH2(True, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro, False)
                        End If
                        'Call ImporteTotalDH(True, rs_Items!itemiclong, True, cadena)
                    Else
                        'primer parametro es "S" usa separador decimal
                        If Mid(programa, 17, 1) = "S" Then
                            Call ImporteTotalDH2(False, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro, True)
                        Else
                            Call ImporteTotalDH2(False, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro, False)
                        End If
                    End If
                Else
                    cadena = " ERROR "
                    Flog.writeline Espacios(Tabulador * 2) & "Faltan Parámetros en el Item " & rs_Items!itemicnro & " o esta mal definido. Se debe indicar si el total es D (debe) o H (haber)."
                End If
            'Fin - LED - CAS-30304 - BANCO INDUSTRIAL - CUSTOM ITEM SUMATORIA DEBE Y HABER
                 
             Case "IMPORTETOTALHABERN A,A" To "IMPORTETOTALHABERN Z,Z"
                Call Validacion(programa, rs_Items!itemiclong, cadena, rs_Procesos!pliqnro, rs_Procesos!vol_cod, Asinro)
                 
            'FGZ - 01/08/2013 -------------------------
            Case "TAB 1" To "TAB 9":
                If Len(programa) > 4 Then
                    Cantidad = Mid(programa, 5, 1)
                Else
                    Cantidad = "1"
                End If
                cadena = String(CLng(Cantidad), Chr(9))
            'FGZ - 01/08/2013 -------------------------
                'FGZ - 13/09/2012 ------------
            Case "TOTALDEBEHABER A,A" To "TOTALDEBEHABER Z,Z":
                Fecha = Trim(Mid(programa, 16, 1))
                completa = (UCase(Mid(programa, 18, 1)) = "S")
                Select Case Fecha
                    Case "D":
                        'Call ImporteTotalDebeHaber(True, Completa, rs_Items!itemiclong, cadena)
                        Call ImporteTotalDebeHaber(True, completa, rs_Items!itemiclong, cadena, nroliq, ProcVol, Asinro)
                    Case "H":
                        'Call ImporteTotalDebeHaber(False, Completa, rs_Items!itemiclong, cadena)
                        Call ImporteTotalDebeHaber(False, completa, rs_Items!itemiclong, cadena, nroliq, ProcVol, Asinro)
                End Select
                
            '///// casos especiales Bco. Industrial Azul - BIA - start ////////
            Case "TOTALSUC":
                Call totalRegistros(TotSuc, True, rs_Items!itemiclong, cadena)
                    
            Case "TOTALCOMPROB":
                Call totalRegistros(TotComprobante, True, rs_Items!itemiclong, cadena)
        
            '///// casos especiales Bco. Industrial Azul - BIA - end ////////
                
            Case Else
                cadena = " ERROR "
                Flog.writeline Espacios(Tabulador * 2) & "Programa inexistente o error de Sintaxis en programa. Item " & rs_Items!itemicnro
            End Select
        End If
            
        'If Mid(cadena, 1, 2) <> "RR" Or primero Then 'Comentado versión 1.31
        If primero Then
            If Aux_Linea = "" Then
                Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
            Else
                If cadena <> Enter And vinoEnter Then
                    Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
                    vinoEnter = False
                Else
                    Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
                End If
                
                'Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
            End If
        Else
            If EsImporte Then 'JAZ Agregado 22-07-11 ver 1.33
                Aux_Linea = Aux_Linea & separadorCampos & cadena 'JAZ Agregado 22-07-11 ver 1.33
                EsImporte = False
            Else 'Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)   'Comentado versión 1.31
                If cadena <> Enter And vinoEnter Then
                    Aux_Linea = Aux_Linea & Mid(cadena, 1, rs_Items!itemiclong)
                    vinoEnter = False
                Else
                    Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong)
                End If
                
                'Aux_Linea = Aux_Linea & separadorCampos & Mid(cadena, 1, rs_Items!itemiclong) 'Agregado ver. 1.31
            End If
        End If
        primero = False
        rs_Items.MoveNext
    Loop
    If Not primero Then
        Progreso = Progreso + 1
        TiempoAcumulado = GetTickCount
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & _
                 ", bprctiempo ='" & CStr((TiempoAcumulado - TiempoInicialProceso)) & _
                 "' WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    End If
    'Escribo en el archivo de texto
    'Aux_Relleno = Space(256 - Len(Aux_Linea))
    If Trim(Aux_Linea) <> "" Then
        'fExport.writeline Aux_Linea '& Aux_Relleno
        fAuxiliarPie.writeline Aux_Linea
    End If
End If

fAuxiliarEncabezado.Close
fAuxiliarDetalle.Close
fAuxiliarPie.Close

If separarArchivo <> -1 Then
'leo el archivos temporal de encabezado y escribo en el de salida.
    On Error Resume Next
    intentos = 0
    Err.Number = 1
    Do Until Err.Number = 0 Or intentos = 10
        Err.Number = 0
        Set fAuxiliarEncabezado = fs.GetFile(directorio & "\fencab.tmp")
        If fAuxiliarEncabezado.Size = 0 Then
            Err.Number = 1
            intentos = intentos + 1
        End If
    Loop
    On Error GoTo CE
   strLinea2 = ""
    If Not intentos = 10 Then
       'Abro el archivo
        On Error GoTo CE
        Set fAuxiliarEncabezado = fs.OpenTextFile(directorio & "\fencab.tmp", ForReading, TristateFalse)
        
        'sebastian stremel
        Set fAuxiliarDetalle = fs.GetFile(directorio & "\fdet.tmp")
        Set fAuxiliarPie = fs.GetFile(directorio & "\fpie.tmp")
        
        If fAuxiliarDetalle.Size = 0 And fAuxiliarPie.Size = 0 Then
            Do While Not fAuxiliarEncabezado.AtEndOfStream
                strLinea = fAuxiliarEncabezado.ReadLine
                If Not fAuxiliarEncabezado.AtEndOfStream Then
                    strLinea2 = strLinea2 & strLinea & Chr(13) & Chr(10)
                    fExport.writeline strLinea
                Else
                    fExport.Write strLinea
                    strLinea2 = strLinea2 & strLinea
                End If
            Loop
        Else
            Do While Not fAuxiliarEncabezado.AtEndOfStream
                
                strLinea = fAuxiliarEncabezado.ReadLine
                fExport.writeline strLinea
                strLinea2 = strLinea2 & strLinea & Chr(13) & Chr(10)
                'fExport.Write strLinea
            Loop
        End If
        fAuxiliarEncabezado.Close
        'fAuxiliarPie.Close
        'fAuxiliarDetalle.Close
        'hasta aca
        
        'Do While Not fAuxiliarEncabezado.AtEndOfStream
        '    strLinea = fAuxiliarEncabezado.ReadLine
        '    fExport.writeline strLinea
            'fExport.Write strLinea
        'Loop
        'fAuxiliarEncabezado.Close
    End If
    'fExport.Close

    'Borro el auxiliar
    fs.DeleteFile directorio & "\fencab.tmp", True
'Fin de auxiliar de encabezado.
' ------------------------------------------------------------------------
'leo el archivos temporal de detalle y escribo en el de salida.
    On Error Resume Next
    intentos = 0
    Err.Number = 1
    Do Until Err.Number = 0 Or intentos = 10
        Err.Number = 0
        Set fAuxiliarDetalle = fs.GetFile(directorio & "\fdet.tmp")
        If fAuxiliarDetalle.Size = 0 Then
            Err.Number = 1
            intentos = intentos + 1
        End If
    Loop
    On Error GoTo CE
   
   If Not intentos = 10 Then
       'Abro el archivo
        On Error GoTo CE
        Set fAuxiliarDetalle = fs.OpenTextFile(directorio & "\fdet.tmp", ForReading, TristateFalse)
    
        'sebastian stremel
       
        Set fAuxiliarPie = fs.GetFile(directorio & "\fpie.tmp")
        
        If fAuxiliarPie.Size = 0 Then
            Do While Not fAuxiliarDetalle.AtEndOfStream
                strLinea = fAuxiliarDetalle.ReadLine
                If Not fAuxiliarDetalle.AtEndOfStream Then
                    fExport.writeline strLinea
                    strLinea2 = strLinea2 & strLinea & Chr(13) & Chr(10)
                Else
                    fExport.Write strLinea
                    strLinea2 = strLinea2 & strLinea
                End If
            Loop
        Else
            Do While Not fAuxiliarDetalle.AtEndOfStream
                strLinea = fAuxiliarDetalle.ReadLine
                strLinea2 = strLinea2 & strLinea & Chr(13) & Chr(10)
                fExport.writeline strLinea
                'fExport.Write strLinea
            Loop
        End If
        fAuxiliarDetalle.Close
        'fAuxiliarPie.Close
    
        'Do While Not fAuxiliarDetalle.AtEndOfStream
        '    strLinea = fAuxiliarDetalle.ReadLine
        '    fExport.writeline strLinea
        '    'fExport.Write strLinea
        'Loop
        'fAuxiliarDetalle.Close
    End If
    'fExport.Close

    'Borro el auxiliar
    fs.DeleteFile directorio & "\fdet.tmp", True
'Fin de auxiliar de detalle.
' ------------------------------------------------------------------------
'leo el archivos temporal de detalle y escribo en el de salida.
    On Error Resume Next
    intentos = 0
    Err.Number = 1
    Do Until Err.Number = 0 Or intentos = 10
        Err.Number = 0
        Set fAuxiliarPie = fs.GetFile(directorio & "\fpie.tmp")
        If fAuxiliarPie.Size = 0 Then
            Err.Number = 1
            intentos = intentos + 1
        End If
    Loop
    On Error GoTo CE
   
   If Not intentos = 10 Then
       'Abro el archivo
        On Error GoTo CE
        Set fAuxiliarPie = fs.OpenTextFile(directorio & "\fpie.tmp", ForReading, TristateFalse)

        Do While Not fAuxiliarPie.AtEndOfStream
            strLinea = fAuxiliarPie.ReadLine
            'sebastian stremel - 28/02/2013 - CAS-16908 - GESTION COMPARTIDA - Custom en Exportación de Asiento - 4
            ' si es la ultima linea del pie no hago salto de carro
            If Not fAuxiliarPie.AtEndOfStream Then
                fExport.writeline strLinea
                strLinea2 = strLinea2 & strLinea & Chr(13) & Chr(10)
            Else
                fExport.Write strLinea
                strLinea2 = strLinea2 & strLinea
            End If

            'fExport.Write strLinea
        Loop
        fAuxiliarPie.Close
       
    End If
    'fExport.Close

    'Borro el auxiliar
    fs.DeleteFile directorio & "\fpie.tmp", True
'Fin de auxiliar de Pie.
' ------------------------------------------------------------------------
End If
primero = True


'Cierro el archivo creado
fExport.Close
 If TipoArchivo = 17 Then
           Set fExport2 = CreateObject("ADODB.Stream") 'Create the stream
           fExport2.Type = adTypeText
           fExport2.Charset = "UTF-8" 'Indicate the charactor encoding
           fExport2.Open 'Initialize the stream
            fExport2.Position = 0 'Reset the position
            fExport2.WriteText strLinea2
           Archivo = "\" & Replace(Archivo, "\\", "\")
                  'Archivo = Replace(Archivo, "\\", "\")
            Flog.writeline "Archivo a convertir: " & Archivo
            fExport2.SaveToFile Archivo, 2 'Save the stream to a file
            fExport2.Close
            Set fExport2 = Nothing
         End If


'Si existe una direccion en la tabla sistema campo sis_expseguridad copio los archivos en la direccion
StrSql = " SELECT sis_expseguridad FROM sistema "
OpenRecordset StrSql, rs_Sistema

If Not EsNulo(rs_Sistema!sis_expseguridad) Then
    Flog.writeline "Los archivos se moveran a la carpeta: " & rs_Sistema!sis_expseguridad
    moverArchivos directorio, rs_Sistema!sis_expseguridad, True
Else
    Flog.writeline "No existe carpeta configurada en la tabla sistema, los archivos no se moveran. "
End If
'Fin de la transaccion
MyCommitTrans


If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
If rs_Items.State = adStateOpen Then rs_Items.Close
If rs_Sistema.State = adStateOpen Then rs_Sistema.Close

Set rs_Procesos = Nothing
Set rs_Periodo = Nothing
Set rs_Modelo = Nothing
Set rs_Items = Nothing

Exit Sub
CE:
    Flog.writeline Espacios(Tabulador * 1) & "Error: " & Err.Description
    HuboError = True
    MyRollbackTrans

    If rs_Procesos.State = adStateOpen Then rs_Procesos.Close
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If rs_Modelo.State = adStateOpen Then rs_Modelo.Close
    If rs_Items.State = adStateOpen Then rs_Items.Close
    
    Set rs_Procesos = Nothing
    Set rs_Periodo = Nothing
    Set rs_Modelo = Nothing
    Set rs_Items = Nothing
End Sub

Public Sub LevantarParamteros(ByVal bpronro As Long, ByVal parametros As String)
' --------------------------------------------------------------------------------------------
' Descripcion: Procedimiento para levantar los parametros pasados en batch_proceso en bprcparam
' Autor      : FGZ
' Fecha      :
' Ult. Mod   :
' Fecha      :
' --------------------------------------------------------------------------------------------
Dim pos1 As Integer
Dim pos2 As Integer
Dim separador As String

Dim Periodo As Long
Dim asiento As String
Dim Empresa As Long
Dim TipoArchivo As Long
Dim ProcVol As Long
Dim ModSalidaAsiento As Long
Dim separarArchivo As Integer
Dim OrdenarPor As Long

'Orden de los parametros
'pliqnro
'Asinro, lista separada por comas
'tipoarchivo
'proceso de volcado, 0=todos
Flog.writeline "Entro a levantar parametros "

separador = "@"
' Levanto cada parametro por separado
If Not IsNull(parametros) Then
    If Len(parametros) >= 1 Then
        pos1 = 1
        pos2 = InStr(pos1, parametros, separador) - 1
        Periodo = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, separador) - 1
        asiento = Mid(parametros, pos1, pos2 - pos1 + 1)
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, separador) - 1
        TipoArchivo = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, separador) - 1
        ProcVol = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        'If InStr(pos1, parametros, Separador) > 0 Then
                    
        pos1 = pos2 + 2
            'pos2 = Len(parametros)
        pos2 = InStr(pos1, parametros, separador) - 1
        ModSalidaAsiento = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
        'Else
        '    pos2 = Len(parametros)
        '    ProcVol = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
            
        '    ModSalidaAsiento = CLng(0)
        'End If
        
        pos1 = pos2 + 2
        pos2 = InStr(pos1, parametros, separador) - 1
        Empresa = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        'sebastian stremel 02/11/2012
        
'        pos1 = pos2 + 2
'        pos2 = Len(parametros)
'        Empresa = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        'hasta aca
        
        pos1 = pos2 + 2
        'pos2 = Len(parametros)
        pos2 = InStr(pos1, parametros, separador) - 1
        separarArchivo = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
                
        pos1 = pos2 + 2
        pos2 = Len(parametros)
        OrdenarPor = CLng(Mid(parametros, pos1, pos2 - pos1 + 1))
        
    End If
    
End If
Call Generacion(bpronro, Periodo, asiento, Empresa, TipoArchivo, ProcVol, ModSalidaAsiento, separarArchivo, OrdenarPor)
End Sub






