Attribute VB_Name = "mdlVersiones"
Option Explicit

'Const Version = 1.01    'Control para evitar sobreescribir archivos
'Const FechaVersion = "15/12/2005"


'Const Version = 1.02
'Const FechaVersion = "04/09/2006"   'FGZ
''                               Modificaciones: El proceso de alerta disparaba un proceso de mensajeria por cada tipo de mensaje a enviar pero....
''                                               Cada proceso de mesaje procesa todo lo que encuentre en el directorio de attachs por lo que no veo objeto en insertar varios procesos ==>
''                                               Ahora inserto un solo proceso.


'Const Version = 1.03
'Const FechaVersion = "05/09/2006"   'FGZ
''                               Modificaciones:
''                               Los nombres de los archivos para los mails de esta alerta empiezan con el bpronro de este proceso
''                               Set MsgFile = fs2.CreateTextFile(BProNro & "_" & AlertaFileName & ".msg", True)
''                   OBS: Dada esta modificacion, hay que modificar el proceso de mensajeria para que filtre los archivos a procesar

'Const Version = "1.04"
'Const FechaVersion = "21/06/2007"   'FGZ
''                               Modificaciones:
''                               SUB MailsAEmpleado: estaba faltando el segundo guion bajo y por eso el proceso de mensajeria no lo levanta
''                               AlertaFileName = dirsalidas & "\msg_" & BProNro & "_ale_" & Replace(FormatDateTime(Date, 2), "/", "-") & "_" & Replace(FormatDateTime(Time, 4), ":", "-") & "-" & String(2 - Len(Second(Now)), "0") & Second(Now) & Contador

'Const Version = "1.05"
'Const FechaVersion = "19/11/2007"   'FGZ
''                               Modificaciones:
''                               Se agregó un nuevo tipo de Notificaciones (Postulantes para Teleperformance)

'Const Version = "1.06"
'Const FechaVersion = "13/02/2009"   'FGZ
''                               Modificaciones:
''                                   Se agregó el progreso en 1 cuando arranca el proceso
''                                   Encriptacion de string de conexion
''                                   Alter Schema para Oracle

'Const Version = "1.07"
'Const FechaVersion = "13/02/2009"   'Lisandro Moro
''                               Modificaciones:
''                                   Se agregó el sub MailsAQueryDelAlerta
''                                   Se encarga de ejecutar la query correspondiente y mandar el resultado por mail a los distintos destinatarios
''                                   Se debe incluir el campo tercero el SELECT en primer lugar
'

'Const Version = "1.08"
'Const FechaVersion = "07/04/2009"   'FGZ
'                               Modificaciones:
'                                   Se agregó la opcion de agregar la alerta una sola vez, es decir,
'                                   que controle si ya habia enviado la alerta con el mismo resultado a la misma persona.

'Const Version = "1.09"
'Const FechaVersion = "14/04/2010"   'MB
''                               Modificaciones:
''                                   Se agregó el sub MailsAQueryDelAlertaEmp
'''                                   Se encarga de ejecutar la query correspondiente y mandar el resultado por mail a los distintos destinatarios
'''                                   Se debe incluir el campo tercero el SELECT en primer lugar y lo toma de mail del empleado.


'Const Version = "1.10"
'Const FechaVersion = "13/09/2010"   'Lisandro Moro
'                               Modificaciones:
'                                   Se agregaron los templates

'Const Version = "1.11"
'Const FechaVersion = "25/02/2011"  'Leticia A. - Se creo un nuevo mensaje de Alerta que muestra información de las Interfaces planificadas.


'Const Version = "1.12"
'Const FechaVersion = "10/06/2011"  'Leticia A. - Se modifico la forma de buscar el directorio de Templates, ahora se usa la dirección de sis_direntradas

'Public Const Version = "1.13"
'Public Const FechaVersion = "10/08/2012"  'FGZ - Se agregó un procedimiento para poder enviar un mail por cada resultado de la alerta
'                               Templates: Nuevos tags para poder utilizar
'                                   ññCxxx: se reemplazará por la columna xxx del resultado de la alerta , ejemplo ññC000 se reemplaza por el primer campo del resultado de la alerta
'                                   ññIxxx: se reemplazará por la imagen de tipo xxx asociado al empleado correspondiente al resultado de la alerta(el legajo debe estar en la primer columna del resultado)

'Public Const Version = "1.14"
'Public Const FechaVersion = "06/06/2013"  'MDZ - se modifico el procedimiento MailsAUsuario que generaba mal la tabla a adjuntar al mail

'Public Const Version = "1.15"
'Public Const FechaVersion = "20/08/2013"  'FGZ - Se mejoró el manejo de errores. Si hay errores y la Alarma es recurrente igual se replanifica (salvo que sean errores generales de version o directorios, etc).

'Public Const Version = "1.16"
'Public Const FechaVersion = "27/09/2013"  'FGZ - Se mejoró el manejo de errores. Se controla que el registro del proceso tenga como parametro el nro de Alerta.
'                                           Ademas se controla si se asociaron parametros al query y no estan resueltos.

'Public Const Version = "1.17"
'Public Const FechaVersion = "28/04/2014"  ' Dimatz Rafael - CAS 24393 - Se recupera el campo SSL de la Tabla
''                                           Verifica si el Mail tiene SSL

'Public Const Version = "1.18"
'Public Const FechaVersion = "09/06/2014"  ' FB - CAS-25868 - H&A - Bug Proceso Alertas
'                                           Se verifica que el estado del recordset (objRs2) no sea cerrado.
'                                           Se muestra una advertencia en el caso en que el recordset este cerrado
'                                           Se filtra el archivo Thumbs.db para que no se adjunte en los mails
'                                           Ademas se agregó el tag ññA003 para la descripcion extendida de la alerta


'Public Const Version = "1.19"
'Public Const FechaVersion = "22/07/2014"  ' Fernandez, Matias - CAS-25868 - H&A - Bug Proceso Alertas- replanifica a pesar de los errores.

'Public Const Version = "1.20"
'Public Const FechaVersion = "05/08/2014"  ' Mauricio Zwenger - CAS-24666 - Se agrego tipo de Notificacion Grupo de Usuarios (17)

'Public Const Version = "1.21"
'Public Const FechaVersion = "13/08/2014"  ' Mauricio Zwenger - CAS-24666 - Se corrigio bug. no se enviaban mail al tipo 17 si no tenia la opcion de enviar solo una vez

'Public Const Version = "1.22"
'Public Const FechaVersion = "13/08/2014"  ' Mauricio Zwenger - CAS-24666 -
'                                          Se corrigio error al generar nombre de archivo para body de email. Si el Alerta debia enviar mas de un mail por resultado y tiene configurado un template
'                                          se pisan los archivos. Se agrego parametro fileName en funciones


'Public Const Version = "1.23"
'Public Const FechaVersion = "04/12/2014"  'Fernandez, Matias -CAS-28334 - HyA - Error en replanificación de alerta mensual
'                                          'Se corrigio el calculo de la fecha de replanificacion

'Public Const Version = "1.24"
'Public Const FechaVersion = "22/12/2014"  'LED CAS-27150 - GALICIA SEG - Capacitación_varios
''                                          'Tipo de notificacion 18 para galicia seguros, con agrupamiento de datos, multiple attach's en el mail, campo alemultattach = -1 tabla alertas.

Public Const Version = "1.25"
Public Const FechaVersion = "24/11/2015"  'Dimatz Rafael - CAS-32369 - HYA - Modificación de asunto de mail en de alertas
'                                          'Se sacó la palabra fija Alerta del asunto del mail y pone la Descripcion de la Alerta



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

    If Version >= "1.13" Then
        'Revisar los campos

        'Alertas . campo que indica si se debe enviar un mail por cada resultado de la alerta
        'ALTER TABLE [dbo].[alertas] ADD [mailxresul] [smallint] NOT NULL DEFAULT 0

        Texto = "Revisar los campos: alertas.mailxresul"
        StrSql = "Select mailxresul from alertas where alenro = 1"
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

