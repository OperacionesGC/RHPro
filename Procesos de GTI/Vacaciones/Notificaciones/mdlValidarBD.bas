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

'Const Version = 2.02     'FAF - Los dias de anticipacion es configurable por confrep
'Const FechaVersion = "10/08/2006"

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


'------------------------------------------------------------------------------------
'Const Version = "3.00"
'Const FechaVersion = "14/04/2010" 'FGZ
'           Ahora los periodos de vacaciones ahora pueden tener alcance por estructura
'Global Const Version = "3.01"
'Global Const FechaVersion = "11/11/2013" ' Gonzalez Nicolás - CAS-19425 - H&A - Mapeo GIV multi-pais R4v1
                                          ' Se agregó Mdlidioma.
                                          ' Se agregó Política 1515
                                          ' Se agrega validación para política 1515 (Solo para PY)
                                          ' Se movieron los comentarios a mdlValidarBD
'Global Const Version = "3.02"
'Global Const FechaVersion = "12/03/2014" ' Gonzalez Nicolás - CAS-19425 - H&A - Mapeo GIV multi-pais R4v1
                                         ' Se corrigió error en recordset
'Global Const Version = "3.03"
'Global Const FechaVersion = "17/10/2014" ' Carmen Quintero - CAS-26298 - H&A - Actualización Contínua Junio y Julio ADP Teylem bug notificación de vacaciones
                                         ' Se modificó el insert que se realiza cuando se procesa por primera vez
'Global Const Version = "3.04"
'Global Const FechaVersion = "21/10/2014" ' Carmen Quintero - CAS-26298 - H&A - Actualización Contínua Junio y Julio ADP Teylem bug notificación de vacaciones [Entrega 2]
                                         ' Modificaciones varias
'
' Global Const Version = "3.05"
'Global Const FechaVersion = "20/02/2015" ' Fernandez, Matias - CAS-29578 - TEYLEM - No se Generan las Notificaciones
'                                         ' se corrige la sintaxis del insert, y el condicional para poder insertar
                                         
'Global Const Version = "3.06"
'Global Const FechaVersion = "01/04/2015" 'Margiotta, Emanuel - CAS 16645 se saco variable de fecha desde y hasta que etaban globales y duplicadas

Global Const Version = "3.07"
Global Const FechaVersion = "06/04/2015" 'Margiotta, Emanuel - CAS 16645 se quito la validacion de la política 1515 cuando busca la Licencia
                                          
                                          

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
 '7 - El salvador
 

    If Version >= "3.01" Then
        Select Case codPais
        Case 6: 'Paraguay
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
