Attribute VB_Name = "mdlVersiones"

Option Explicit


'Global Const Version = "1.00"
'Global Const FechaModificacion = " "
'Global Const UltimaModificacion = " " 'C testaseca - Version inicial

'Global Const Version = "1.01"
'Global Const FechaModificacion = "03/09/2007"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Toda datos de batch proceso
                                      'se agrego encabezado del log
                                      'En ejecutar ahora No preguntaban If objRs2.EOF y daba error si habian borrado el schedule
                                      'Chequear si correr la sql de oracle o no de acuerdo al valor del la config de empresa
                                      'Si confint = -1 de la tabla confper entonces SQL en Oracle

'Global Const Version = "1.02"
'Global Const FechaModificacion = "07/08/09"
'Global Const UltimaModificacion = " " 'Martin Ferraro - Se quito TOP de sql para que funcione en oracle
'                                      'Encriptacion de stringconn
'                                      'Se cambio el formato del nombre del log
                                      
'Global Const Version = "1.03"
'Global Const FechaModificacion = "25/05/2011"
'Global Const UltimaModificacion = " " 'FGZ - Se cambio orden de la llamada a cargarconfiguracionesBasicas
''                                     Se agregó un control por si la ejecucion fué manual o programada
''                                     Para que ejecute manualmente se debe poner cualquier valor distinto de cero cuando se ejecuta el indicardo manualmente

'Global Const Version = "1.04"
'Global Const FechaModificacion = "25/08/2011"
'Global Const UltimaModificacion = " " 'FGZ - Ahora si el indicador es detallado por empleado se guarda el resultado de su query en la tabla ind_historia_det

'Global Const Version = "1.05"
'Global Const FechaModificacion = "10/04/2012"
'Global Const UltimaModificacion = " " 'Deluchi Ezequiel - validacion varias y se chequean las fechas de planificacion de los indicadores y se arma correctamente para la funcion EjecutarAhora(indnro, schednro)

'Global Const Version = "1.06"
'Global Const FechaModificacion = "28/05/2012"
'Global Const UltimaModificacion = " " 'Deluchi Ezequiel - Correccion en la validacion de la fecha en la funcion ejecutar ahora.

'Global Const Version = "1.07"
'Global Const FechaModificacion = "29/06/2012"
'Global Const UltimaModificacion = " " 'Dimatz Rafael - 15513 - En el ultimo insert ind_historia se cambio l_Resu_Total por l_Resu


'Global Const Version = "1.08"
'Global Const FechaModificacion = "06/08/2012"
'Global Const UltimaModificacion = " " 'FGZ - Correcciones varias
''                                           Manejador de errores: se cambió el manejo.
''                                           Cuando no hay indicadores activos el proceso quedaba en estado procesando. Ahora lo deja en "Procesado"
''                                           Se le agregó un Funcion de control de versiones en un nuevo modulo del proceso (las versiones del proceso se movieron al nuevo modulo)


' ****************************************
''       Version NO LIBERADA AUN
'Global Const Version = "1.09"
'Global Const FechaModificacion = "01/11/2013"
'Global Const UltimaModificacion = " " 'FGZ - Correcciones
''                                           Con la correccion de la version 1.07 se rompió el detalle
''                                           En el ultimo insert ind_historia se cambio l_Resu por l_Resu_Total
''
''                                           Le agregué control sobre indice unico en tabla ind_historia



'Global Const Version = "1.10"
'Global Const FechaModificacion = "06/11/2013"
'Global Const UltimaModificacion = " " 'JPB  - CAS-13764 - H&A - Tablero de Comandos del empleado
'                                           Cuando el indicador es detallado por empleado, al insertar en el hitorico ahora guarda el detalle del indicador


'Global Const Version = "1.11"
'Global Const FechaModificacion = "21/10/2014"
'Global Const UltimaModificacion = " " 'Ruiz Miriam  - CAS-26972 - H&A - Bugs detectados en R4 - Error en el proceso de indicadores
'                                         Se corrigio el caso en que el indicador estuviera configurado con detalle y la query no trajera todos los parámetros
'                                         Se corrigió el caso en que el indicador tuviera null en el campo inddetalle

Global Const Version = "1.12"
Global Const FechaModificacion = "30/10/2015"
Global Const UltimaModificacion = " " 'Ruiz Miriam  - CAS-26972 - RH Pro (Producto) - Bug en el progreso del proceso RHProIndicador
'                                         Se corrigió progreso



' ==================================================================================================================


Public Function ValidarV(ByVal Version As String, ByVal TipoProceso As Long, ByVal TipoBD As Integer) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Validacion de estructura de BD
' Autor      : FGZ
' Fecha      : 06/08/2012
' ---------------------------------------------------------------------------------------------
Dim V As Boolean
Dim Texto As String
Dim rs As New ADODB.Recordset

On Error GoTo ME_Version

V = True

    If Version >= "1.04" Then
        'Revisar los campos

        '/*Agrega el campo inddetalle a la tabla indicador para especificar que el indicador se va a detallar por empleado */
        'ALTER TABLE indicador ADD inddetalle smallint
        Texto = "Revisar los campos: indicador.inddetalle"
        StrSql = "Select inddetalle from indicador where indnro = 1"
        OpenRecordset StrSql, rs

        '/* Esta tabla va a asociar el historico del indicador asociado a un empleado */
        'CREATE TABLE ind_historia_det(
        ' indnro int not null,
        'ternro int not null,
        'indhisvalor decimal(19,4),
        'indhisdetnro int IDENTITY (100,1),
        'indhisfec datetime,
        'indhishora varchar(4),
        'indhisdesabr varchar(50),
        'indhisnro int
        ')
        Texto = "Revisar los campos: indnro,ternro,indhisvalor,indhisdetnro,indhisfec,indhishora,indhisdesabr,indhisnro " & " " & " de la tabla ind_historia_det "
        StrSql = "Select indnro,ternro,indhisvalor,indhisdetnro,indhisfec,indhishora,indhisdesabr,indhisnro from ind_historia_det where indnro = 1"
        OpenRecordset StrSql, rs

        V = True
    End If

    ValidarV = V
Exit Function

ME_Version:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Estructura de BD incompatible con la version del proceso."
    Flog.writeline Espacios(Tabulador * 1) & Texto
    Flog.writeline
    V = False
End Function

