Attribute VB_Name = "mdlValidarBD"
Option Explicit
'*********************************************************************************************************************
'***********************************************                       ***********************************************
'*********************************************** -- CONSIDERACIONES -- ***********************************************
'***********************************************                       ***********************************************
'*********************************************************************************************************************
'1 - SI NO EXISTE EL MODULO DEL PAIS, SE DEBE CREAR --> NOMENCLATURA: mdlNombrepais
'2 - LA POLITICA 1515 ES OBLIGATORIA | PERMITE UTILIZAR MÁS DE UN MODELO (SOLO COMO INDIVIDUAL O ESTRUCTURA)
'3 - SI SE AGREGA UN MODELO NUEVO - AGREGAR EN CASE politica1515()
'4 - VALIDAR VERSIONES EN FUNCION -->  ValidarVBD()
'5 - AGREGAR ETIQUETAS CON MULTILENGUAJE (ANTES DE CREAR UNA ETIQUETA VERIFICAR QUE NO EXISTA UNA SIMILAR)

'En todos los procedimientos de calculo de dias correspondientes se deben validar las políticas:
'-> 1501 - Proporcióon de días
'-> 1502 - Escala de Vacaciones
'-> 1508 - Licencias por Maternidad
'-> 1513 - Días trabajados en el último año
'-> 1511 - Vacaciones Acordadas


' MODELOS DE PAISES/CUSTOM (Al agregar Nuevo modelo verificar fuentes ASP de vacaciones en shared/inc/vacaciones_XXX.inc)
' y en politica1515()
'0 - Argentina | 1 - Uruguay | 2 - Chile | 3 - Colombia | 4 - Costa Rica 5 - Portugal | 6 - Paraguay | 7 - Perú

'****************************************                              ******************************************
'****************************************                              ******************************************
'Version: 1.01  'Inicial

'Const Version = 1.01    'Version Inicial
'Const FechaVersion = "17/10/2005"

'Const Version = 1.02    'Version con otra conexion para el progreso
'Const FechaVersion = "23/11/2005"

'Const Version = 2.01     'Revision general
'Const FechaVersion = "01/12/2005"

'Const Version = 2.02     'Correccion en el insert de las licencias
'Const FechaVersion = "07/12/2005"

'Const Version = 2.03     'Detalles de logs en el insert de la licencia
'Const FechaVersion = "22/02/2006"

'Const Version = 2.04    'Correccion error al obtener ultimo insertado
                        'Controlar que exista fase activa al insertar licencia
'Const FechaVersion = "06/03/2006"

'Const Version = 2.05    'Se agrego la tabla gti_justificacion a los procesos de vacaciones
'Const FechaVersion = "14/11/2006"

'Const Version = 2.06
'Const FechaVersion = "17/04/2007"
''Autor: FGZ
''Modificacion:
''       Se toma un parametro nuevo () para configurar si corto las licencias cuando hay un feriado.
''       Antes cortaba siempre las licencias cuando se excluia los feriados.
''       Ahora en el parametro


'Const Version = "2.07"
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

'Const Version = "2.08"
'Const FechaVersion = "18/04/2008" 'FGZ
'Se en el sub Licencias3
'       Cuando habia una licencia superpuesta que no era de vacaciones generaba las licencias
'                Solo andaba bien el chequeo de superposicion cuando
'                el tipo de licencia superpuesto era  de vacaciones
'                para cualquier otro tipo de licencia las generaba igual

'Const Version = 2.09    'Se agrego la politica 1509 - No genera situacion de revista si se configura - agd
'                        ' Se agrego la encriptacion al string de conexion
'Const FechaVersion = "12/06/2009" 'Lisandro Moro


'Const Version = "2.10"
'Const FechaVersion = "24/06/2009" 'FGZ
''       Se agrego la llamada a la politica 1509 qe faltaba
'

'Const Version = "2.11"
'Const FechaVersion = "26/06/2009"   'FGZ
''           Nueva Politica 1510 - Licencias se gozan por días hábiles.
''               Esta politica se tiene en cuenta para el calculo de dias pedidos y dias de licencia.


'Const Version = "2.12"
'Const FechaVersion = "02/10/2009"   'FGZ
''           Se recompiló la version por una modificacion en un modulo compartido


'------------------------------------------------------------------------------------
'Const Version = "3.00"
'Const FechaVersion = "14/04/2010" 'FGZ
''           Ahora los periodos de vacaciones ahora pueden tener alcance por estructura


'Const Version = "3.01"
'Const FechaVersion = "11/08/2010" 'FGZ
''           Los feriados ahora pueden ser laborables. Ahora las licencias se separan dependiendo de las 2 marcas configuradas.
''           Feriados NO Laborables (la marca actual)
''           Feriados Laborables (la nueva marca)

'Const Version = "3.02"
'Const FechaVersion = "06/09/2010" 'Margiotta, Emanuel
''           Se cambio el número de alcance por estructura de los periodos de vacaciones por el 21
''           Se agrego la funcion para el Control de versiones

'Const Version = "3.03"
'Const FechaVersion = "12/07/2011" 'Margiotta, Emanuel
''           Se modularizo el proceso de dias correspondiente segun el modelo de vacaciones
''           Estos cambios se realizaron por el caso de Syke - Costa Rica

'Const Version = "3.04"
'Const FechaVersion = "06/12/2011" 'Margiotta, Emanuel
''           Se agrego el circuito de firmas a la generacion de licencias - Syke - Costa Rica

'Const Version = "3.05"
'Const FechaVersion = "06/12/2011" 'Margiotta, Emanuel
''           Se corrigio cuando levantaba los parametros, en autorizadas estaba cortando mal el caracter  - Syke - Costa Rica

'Const Version = "3.06"
'Const FechaVersion = "15/05/2013" 'Lisandro Moro
'           Se corrigio cuando levantaba los parametros, en autorizadas estaba cortando mal el caracter  - Syke - Costa Rica (IDEM)

'Const Version = "3.07"
'Const FechaVersion = "26/09/2013" 'Sebastian Stremel
'           Se agrego la vesion para el salvador case 7

'Const Version = "3.08"
'Const FechaVersion = "07/10/2013" 'Fernando Favre
'           CAS-21253 - AMIA -Error en Licencia Vacaciones_ GIV.

'Const Version = "3.09"
'Const FechaVersion = "04/11/2013" 'Carmen Quintero
'           CAS-21383 - VSO - Pollpar - Error GIV [Entrega 4] - Se agregó condición en la funcion SepararLicencias
'           , para el caso cuando el primer dia de la licencia es un feriado.
                                       
                                         
'Global Const Version = "3.10"
'Global Const FechaVersion = "11/10/2013"

'Global Const Version = "3.11"
'Global Const FechaVersion = "17/12/2013" ' Fernandez, Matias - CAS-22354 - AMIA - BUG EN CARGA DE VACACIONES - se corrigio la parte dde se excluyen feriados no laborables y tiene que partir la licencia.
'Global Const Version = "3.12"
'Global Const FechaVersion = "09/01/2014" ' Fernandez, Matias - CAS-22354 - AMIA - BUG EN CARGA DE VACACIONES - se corrigio  a partir de que dia arranca la licencia cuando se separan....

'Global Const Version = "3.13"
'Global Const FechaVersion = "version r3 para giv r4"

'Global Const Version = "3.14"
'Global Const FechaVersion = "31/03/2015- EAM- CAS-16645 - Se ajusto el proceso para argentina version r4"

'Global Const Version = "3.15"
'Global Const FechaVersion = "05/06/2015- MDZ- CAS-26028 - se agrego firmas al generar licencias con circuito activo"

Global Const Version = "3.16"
Global Const FechaVersion = "15/02/2016 - MDZ - CAS-32780 - Se ajusto el proceso para operara con periodos de GIV R4"

'--------------------------------------
'Pendiente de liberar

'Global Const Version = "3.13"
'Global Const FechaVersion = "15/01/2014" ' Margiotta, Emanuel - CAS-21472 - Sykes El Salvador - desarrollo de GIV - Se modifico la funciona del modelo 7 el salvador para que llame Licencias3() igual que argentina ya que usa el complemento de vacaciones


Public Function ValidarVBD(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer, ByVal codPais As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : Gonzalez Nicolás
' Fecha      : 04/05/2012
' Modificado : 17/09/2012 - Se agregó versión 6 PARAGUAY
'            : 11/10/2012 - Se agregó mje. de error de versiones (Codigo del pais con error)
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True
' CODIGOS DE PAISES
 '0 - Argentina
 '1 - Uruguay
 '2 - Chile
 '3 - Colombia
 '4 - Costa Rica
 '5 - Portugal
 '6 - Paraguay
 '7 - Perú
If Version >= "3.10" Then
    'vacdiascor.venc
    Texto = "Revisar los campos: vacdiascor.venc"
    StrSql = "Select venc from vacdiascor where ternro = 1"
    OpenRecordset StrSql, rs
    V = True
    
    'Días correspondientes
    Texto = "Revisar los campos: "
    Texto = Texto & " vacdiascor.vdiascorcantcorr,  vacdiascor.tipvacnrocorr"
    StrSql = "Select vacdiascor.vdiascorcantcorr,  vacdiascor.tipvacnrocorr FROM vacdiascor WHERE ternro = 1"
    OpenRecordset StrSql, rs
    V = True


    Select Case codPais
        Case 4: 'COSTA RICA - SYKES
            'Días correspondientes
            Texto = "Revisar los campos: "
            Texto = Texto & " vacdiascor.vdiasfechasta"
            StrSql = "Select vacdiascor.vdiasfechasta FROM vacdiascor WHERE ternro = 1"
            OpenRecordset StrSql, rs
            Texto = "Revisar los campos: "
            Texto = Texto & " vacacion.ternro"
            StrSql = "Select vacacion.ternro FROM vacacion WHERE ternro = 1"
            OpenRecordset StrSql, rs
            V = True
        Case 6: 'PARAGUAY
            'Días correspondientes
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
    End Select
End If


ValidarVBD = V
    
If V = True Then
    Texto = "Versión correcta"
    Exit Function
End If




ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & "Revisar estructura de Base de Datos para el modelo N°: " & codPais
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function
