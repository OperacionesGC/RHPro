Attribute VB_Name = "mdlValidarBD"
Option Explicit
'*********************************************************************************************************************
'***********************************************                       ***********************************************
'*********************************************** -- CONSIDERACIONES -- ***********************************************
'***********************************************                       ***********************************************
'*********************************************************************************************************************
'1 - SI NO EXISTE EL MODULO DEL PAIS, SE DEBE CREAR --> NOMENCLATURA: mdlPed_Nombrepais
'2 - LA POLITICA 1515 ES OBLIGATORIA | PERMITE UTILIZAR MÁS DE UN MODELO (SOLO COMO INDIVIDUAL O ESTRUCTURA)
'3 - SI SE AGREGA UN MODELO NUEVO - AGREGAR EN CASE politica1515()
'4 - VALIDAR VERSIONES EN FUNCION -->  ValidarVBD()
'5 - AGREGAR ETIQUETAS CON MULTILENGUAJE (ANTES DE CREAR UNA ETIQUETA VERIFICAR QUE NO EXISTA UNA SIMILAR)

'En todos los procedimientos de calculo de dias correspondientes se deben validar las políticas:
'-> 1515 -
'-> 1510 -

' MODELOS DE PAISES/CUSTOM (Al agregar Nuevo modelo verificar fuentes ASP de vacaciones en shared/inc/vacaciones_XXX.inc)
' y en politica1515()
'0 - Argentina | 1 - Uruguay | 2 - Chile | 3 - Colombia | 4 - Costa Rica 5 - Portugal | 6 - Paraguay | 7 - Perú

'****************************************                              ******************************************
'**************************************** -- ULT. VERSION AL INICIO -- ******************************************
'****************************************                              ******************************************

'Version: 1.01  'Inicial

'Const Version = 1.01    'Version Inicial
'Const FechaVersion = "17/10/2005"

'Const Version = 1.02    'Version con otra conexion para el progreso
'Const FechaVersion = "23/11/2005"

'Const Version = 2.01     'Revision general
'Const FechaVersion = "01/12/2005"

'Const Version = 2.02     'Se agrego validacion que la fase del empleado este activa a
'                         'a la fecha del pedido
'                         'No restauraba la varible Aux_Cant_dias cuando ciclaba por empleado
'Const FechaVersion = "06/03/2006"

'Const Version = "2.03"
'Const FechaVersion = "13/11/2007" 'FGZ
''Se cambió la fecha para la cual se resuelve el alcance por estructura de las politicas (sub politica)
''               Se cambió el uso de fecha_desde en los querys por aux_fecha
''                If fecha_desde > Date Then
''                    Aux_fecha = fecha_desde
''                Else
''                    If fecha_hasta > Date Then
''                        Aux_fecha = Date
''                    Else
''                        Aux_fecha = fecha_hasta
''                    End If
''                End If

'Const Version = "2.04"
'Const FechaVersion = "22/06/2009"
'       'FGZ - Se agrego la encriptacion al string de conexion.

'Const Version = "2.05"
'Const FechaVersion = "26/06/2009"   'FGZ
''           Nueva Politica 1510 - Licencias se gozan por días hábiles.
''               Esta politica se tiene en cuenta para el calculo de dias pedidos y dias de licencia.

'Const Version = "2.06"
'Const FechaVersion = "23/10/2009"   'FGZ
''           Nueva Politica 1512 - Vencimiento de Vacaciones.

'Const Version = "2.07"
'Const FechaVersion = "16/11/2009" 'FGZ
''           Problema con la funcion de validacion de version.


'Const Version = "3.00"
'Const FechaVersion = "14/04/2010" 'FGZ
'           Ahora los periodos de vacaciones ahora pueden tener alcance por estructura


'Const Version = "3.01"
'Const FechaVersion = "09/08/2010" 'Marigotta, Emanuel
''           Se agrego en el pedido de dias, que los haga sobre los periodos que estan abiertos.


'Const Version = "3.02"
'Const FechaVersion = "05/10/2010" 'Marigotta, Emanuel
''           Se cambio el Nro de alcance por estructura para los periodos al 21
''           Se agregó la validación de que el periodo esté abierto


'Const Version = "3.03"
'Const FechaVersion = "26/11/2010" 'Marigotta, Emanuel
'           Se corrigió el cálculo para el pedido de días.

'Global Const Version = "3.04"
'Global Const FechaVersion = "11/11/2013" ' Gonzalez Nicolás - CAS-19425 - H&A - Mapeo GIV multi-pais R4v1
                                          ' Se agregó MdlVac_Paraguay,MdlVac_Argentina y Mdlidioma.
                                          ' Se agregó Política 1515
                                          ' Se agrega validación para política 1515 (Solo para PY)
                                          ' Se movieron los comentarios a mdlValidarBD
                                          
                                          
                                          
'Global Const Version = "3.05"
'Global Const FechaVersion = "23/05/2014" ' Fernandez, Matias - CAS-25560 - H&A - GIV (Paraguay) - Bugs en Generar dias pedidos
                                         'Se controla el salto a la etiqueta siguiente que no haga loop sobre registros no abiertos


Global Const Version = "3.06"
Global Const FechaVersion = "02/07/2014" ' Fernandez, Matias - CAS-25560 - H&A - GIV (Paraguay) - Bugs en Generar dias pedidos
                                         'se comentaron las variables globales  fecha_desde y fecha_hasta aparte se corrigio cuando el alcance de la politica es por
                                         'estructura





'---------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------FIN VERSIONES ------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ValidarVBD(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer, ByVal codPais As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion:
' Autor      : Gonzalez Nicolás
' Fecha      : 08/11/2012
' Modificado :
'            :
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
 
  If Version >= "2.06" Then
    'Revisar los campos
    'vacdiascor.venc
    Texto = "Revisar los campos: vacdiascor.venc"
        
    StrSql = "Select venc from vacdiascor where ternro = 1"
    OpenRecordset StrSql, rs

    V = True
    End If
 

    If Version >= "3.04" Then
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
