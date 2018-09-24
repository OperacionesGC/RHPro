Attribute VB_Name = "MdlMail"
Option Explicit

'Const Version = 1.01    'Inicial
'Const FechaVersion = "13/10/2005"

'Const Version = 1.02    'Inicial
'Const FechaVersion = "12/12/2005"

'Const Version = 1.03    'no se estaba manejando bien la cantidad de reintentos
'Const FechaVersion = "18/05/2006"

'Const Version = 1.04    'Mariano Capriz - Se agrego log ya que kedaba el proceso
'                        'colgado y se detecto que el RHProappSrvDefaults.ini
'                        'debe ir en la misma estructura donde estan los procesos
'Const FechaVersion = "07/08/2006"

'Const Version = 1.05    'FGZ - primero creo el log y luego intento establecer la conexion
'                        '       NroProceso as integer por NroProceso as long
'Const FechaVersion = "24/08/2006"

'Const Version = 1.06
'Const FechaVersion = "01/09/2006" 'FGZ - Cierro los archivos de logs que no estaban
''                                        Pongo el end que no estaba

'Const Version = 1.07
'Const FechaVersion = "05/09/2006" 'FGZ - POr la version generada del proceso de alertas V1.03 se genera esta version
''                                  Dada esta modificacion, se modifica el proceso para que filtre los archivos a procesar
''                                   se procesaran solamente los mensajes cuyo nombre comiencen con el nro de proceso batch

'Const Version = 1.08
'Const FechaVersion = "07/09/2006" 'FGZ - Se sacó la generacion del archivo de log floge
''                                        Agregados de log cuando trata de eliminar los archivos

'Const Version = "1.09"
'Const FechaVersion = "22/04/2008" 'FGZ - se cambió el tipo de la variable de MailPort
''

'Const Version = "1.10"
'Const FechaVersion = "13/02/2009"   'FGZ
'                               Modificaciones:
'                                   Encriptacion de string de conexion
'                                   Alter Schema para Oracle

'Const Version = "1.11"
'Const FechaVersion = "29/07/2009"   'Fernando Favre
'                               Modificaciones:
'                                   Se integro una version de Martin - "12/12/2008" 'Martin Ferraro - Envio con copia oculta
'                                   Se modifico para que refleje el avance en el proceso (campo bprcprogreso de la tabla batch_proceso)

'Const Version = "1.12"
'Const FechaVersion = "12/08/2009"   'Fernando Favre
''                               Modificaciones:
''                                   Se modifico el conjunto de caracteres utilizado en la codificacion del mail (charset = "ISO-8859-1").

''------------------------------------------------------------------------------------------------------------------------
'Const Version = "2.00"
'Const FechaVersion = "01/07/2010"   'FGZ
''                               Modificaciones:
''                                   El proceso ahora Utiliza la libreria CDO de Microsoft (dejó de utilizar OSSMTP.dll)
''                                   Mejoras
''                                       Se puede enviar contenido html en el cuerpo del mail
''                                       Se agregaron las opciones de recipiente de CC
'------------------------------------------------------------------------------------------------------------------------
'Const Version = "2.01"
'Const FechaVersion = "13/09/2010"   'Lisandro Moro
''                               Modificaciones:
''                                   Se agregaron los templates.
'------------------------------------------------------------------------------------------------------------------------
'Const Version = "2.02"
'Const FechaVersion = "25/09/2012"   'Lisandro Moro
''                               Modificaciones:
''                                   Se valida con el contenido los attach o como header.
''------------------------------------------------------------------------------------------------------
'Const Version = "2.03"
'Const FechaVersion = "10/05/2013"   'Lisandro Moro
'                               Modificaciones:
'                                   Se agrego el schema en cargarconfiguracionesbasicas.

Const Version = "2.04"
Const FechaVersion = "28/04/2014"   'Dimatz Rafael - CAS 24393 - AMIA
'                               Modificaciones:
'                                   Se agrego el Control de Mail con SSL
'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------

Dim dirsalidas As String
Dim MailHost As String
Dim MailPort As String
Dim MailFromName As String
Dim MailFrom As String
Dim MailTo As String
Dim MailSubject As String
Dim MailBody As String
Dim MailHtmlBody As String
Dim HTMLMailHeader As String
Dim MailAttach As String
Dim MailUser As String
Dim MailPassword As String
Dim f1, fc
Dim HuboErrores As Boolean
Dim Maxretry As Integer
Dim Timeretry As Integer
Dim SinParemetros As Boolean
Dim MailBCC As String
Dim MailCC As String
Dim Nivel
Dim SSL

Sub CargarMensaje(ByVal archivo As String)
'carga las configuraciones basicas para los procesos
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim Tag

    Nivel = Nivel + 1
    Flog.writeline Espacios(Tabulador * Nivel) & "Cargar Mensaje " & f1.Name
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(archivo, ForReading, 0)
    
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If
    
    Nivel = Nivel + 1
    'FromName
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailFromName = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "FromName: " & MailFromName
    
    'Subject
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailSubject = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "Subject: " & MailSubject
    
    'Body1
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailBody = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "Body: " & MailBody
    
    
'    'Attachment
'    If Not f.AtEndOfStream Then
'        strline = f.ReadLine()
'        pos1 = InStr(1, strline, "=") + 1
'        pos2 = Len(strline)
'        MailAttach = Mid(strline, pos1, pos2)
'    End If
'    Flog.writeline Espacios(Tabulador * Nivel) & "Attachment: " & MailAttach
    
    'Esta linea es opcional (puede venir o no)----------------------------------------
    'HtmlBody o Attachment
    MailHtmlBody = ""
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        Tag = Mid(strline, 1, pos1 - 2)
        If UCase(Tag) = "HTMLBODY" Then
            MailHtmlBody = Mid(strline, pos1, pos2)
            Flog.writeline Espacios(Tabulador * Nivel) & "Archivo del Body: " & MailHtmlBody
            
            'Attachment
            If Not f.AtEndOfStream Then
                strline = f.ReadLine()
                pos1 = InStr(1, strline, "=") + 1
                pos2 = Len(strline)
                MailAttach = Mid(strline, pos1, pos2)
            End If
            Flog.writeline Espacios(Tabulador * Nivel) & "Attachment: " & MailAttach
        Else
            MailAttach = Mid(strline, pos1, pos2)
            Flog.writeline Espacios(Tabulador * Nivel) & "Attachment: " & MailAttach
        End If
    End If
    
    '---------------------------------------------------------------------------------
    
    
    'Recipients
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailTo = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "Recipients: " & MailTo
    
    'FromAddress
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailFrom = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "FromAddress: " & MailFrom
    
    'Host
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailHost = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "Host: " & MailHost
    
    'Port
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailPort = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "Port: " & MailPort
    
    'User
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailUser = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "User: " & MailUser
    
    'Password
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailPassword = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "Password: ********" '& MailPassword
       
    'Recipients CCO
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailBCC = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "CCO: " & MailBCC
    
    'Recipients CC
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailCC = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "CC: " & MailCC
    
    'HTMLBody
    If Not f.AtEndOfStream Then
        If MailHtmlBody = "" Then
            strline = f.ReadLine()
            pos1 = InStr(1, strline, "=") + 1
            pos2 = Len(strline)
            MailHtmlBody = Mid(strline, pos1, pos2)
        End If
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "HTMLBODY: " & MailHtmlBody
    
    'HTMLMailHeader
    If Not f.AtEndOfStream Then
        If MailHtmlBody <> "" Then
            strline = f.ReadLine()
            pos1 = InStr(1, strline, "=") + 1
            pos2 = Len(strline)
            HTMLMailHeader = Mid(strline, pos1, pos2)
        End If
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "HTMLMailHeader: " & HTMLMailHeader
    
   'SSL
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        SSL = Mid(strline, pos1, pos2)
    End If
    Flog.writeline Espacios(Tabulador * Nivel) & "SSL: " '& SSL
    
    f.Close
    
    'Leo el html del body
    Flog.writeline Espacios(Tabulador * Nivel) & "Cargando html del Body: " & MailHtmlBody
    If Not EsNulo(MailHtmlBody) Then
        MailBody = ""
        Set f = fs.OpenTextFile(MailHtmlBody, ForReading, 0)
    
        Do While Not f.AtEndOfStream
            MailBody = MailBody & f.ReadLine()
        Loop
    End If
    
    Nivel = Nivel - 1
    Flog.writeline Espacios(Tabulador * Nivel) & "Termina la carga mensaje"
    Nivel = Nivel - 1
End Sub

'
Sub CargarDefaults()
' carga las configuraciones para los mensajes
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer

    On Error GoTo ME_CargaDef
    
    Nivel = Nivel + 1
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.Path & "\RHProappSrvDefaults.ini", ForReading, 0)
    
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If
    ' Retries
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        Maxretry = CInt(Mid(strline, pos1, pos2 - pos1))
    End If
    ' Subject
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        Timeretry = CInt(Mid(strline, pos1, pos2 - pos1))
    End If
    Nivel = Nivel - 1
    f.Close

ME_CargaDef:
    Nivel = Nivel + 1
    Flog.writeline Espacios(Tabulador * Nivel) & "Alerta!!"
    Flog.writeline Espacios(Tabulador * Nivel) & "No se pudo leer el archivo de configuracion (" & App.Path & "\RHProappSrvDefaults.ini)."
    Maxretry = 3
    Timeretry = 1   'Minutos
    Flog.writeline Espacios(Tabulador * Nivel) & "Parametros por default"
    Flog.writeline Espacios(Tabulador * (Nivel + 1)) & "Maxima cantidad de reintentos de envio: " & Maxretry
    Flog.writeline Espacios(Tabulador * (Nivel + 1)) & "Tiempo entre reintentos de envio: " & Timeretry & " minuto."
    Flog.writeline
End Sub


Sub Main()
Dim strCmdLine As String
Dim Nombre_Arch As String
Dim Nombre_ArchA As String
Dim StrSql As String
Dim objRs As New ADODB.Recordset
Dim objRs2 As New ADODB.Recordset
Dim PID As String
Dim Retry As Integer
Dim NroProceso As Long
Dim ArrParametros

  
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
    
    TiempoInicialProceso = GetTickCount

    ' carga las configuraciones basicas, formato de fecha, string de conexion,
    ' tipo de BD y ubicacion del archivo de log
    Call CargarConfiguracionesBasicas
    
    
    HuboErrores = False
    Nombre_Arch = PathFLog & "Mensajes" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    
    
    'Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    'Flog.writeline
    Flog.writeline "Inicio Mensajeria : " & Now
    Flog.writeline
    Nivel = 1
    
    'desactivo momentaneamente
    On Error GoTo 0
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline Espacios(Tabulador * Nivel) & "Problemas en la conexión."
        Exit Sub
    End If
    On Error GoTo CE
    
    'Cambio el estado del proceso a Procesando
    Flog.writeline Espacios(Tabulador * Nivel) & "Cambio el estado del proceso a Procesando"
    StrSql = "UPDATE batch_proceso SET bprcprogreso = 1, bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline
    
    Flog.writeline Espacios(Tabulador * Nivel) & "Seteo de dafaults"
    Call CargarDefaults
    Flog.writeline
    
    'Obtengo los datos del proceso
    Flog.writeline Espacios(Tabulador * Nivel) & "Obtengo los datos del proceso"
    StrSql = "SELECT bprcparam FROM batch_proceso WHERE btprcnro = 25 AND bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       'Obtengo los parametros del proceso
       If Not EsNulo(objRs!bprcparam) Then
            Retry = objRs!bprcparam
        Else
            Retry = 1
        End If
        If Retry <> 1 Then
            Flog.writeline Espacios(Tabulador * Nivel) & "Reintento nro " & Retry
        End If
    Else
        Flog.writeline Espacios(Tabulador * Nivel) & "Error al buscar datos en el proceso: " & NroProceso
        GoTo Fin
    End If
       
    'Directorio Salidas
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel) & "Configuracion tabla sistema"
    StrSql = "select sis_dirsalidas from sistema"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        Nivel = Nivel + 1
        dirsalidas = objRs!sis_dirsalidas & "\attach"
        Flog.writeline Espacios(Tabulador * Nivel) & "Dir Salida: " & dirsalidas
    Else
        Nivel = Nivel + 1
        Flog.writeline Espacios(Tabulador * Nivel) & "No se encuentra configurado sis_dirsalidas."
        GoTo Fin
    End If
    Nivel = Nivel - 1
    
    'Busco Mensaje a enviar
    Dim fso, f, s
    Dim IncPorcAux
    Dim ProgresoAux
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(dirsalidas)
    Set fc = f.Files
    If fc.Count = 0 Then
        Nivel = Nivel + 1
        Flog.writeline Espacios(Tabulador * Nivel) & "No hay mensajes a enviar."
    End If
    Nivel = Nivel - 1
    
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = fc.Count
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = Format((99 / CEmpleadosAProc), "0.0000")
    IncPorcAux = Replace(CStr(IncPorc), ",", ".")
    
    'Actualizo el progreso
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & IncPorcAux & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    For Each f1 In fc
        If InStr(f1.Name, ".msg") And Left(f1.Name, Len(CStr(NroProceso)) + 5) = "msg_" & CStr(NroProceso) & "_" Then
            Flog.writeline
            CargarMensaje (f1.Path)
            If HuboError Then
                GoTo Fin
            End If
            Flog.writeline
            SndMail
            If HuboError Then
                GoTo Fin
            End If
        End If
    
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        ProgresoAux = Replace(CStr(Progreso), ",", ".")
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & ProgresoAux & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Next
    
Fin:
    If Not HuboErrores Then
        'Actualizo el estado del proceso
        StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        Nivel = Nivel + 1
        Flog.writeline Espacios(Tabulador * Nivel) & "Hubo Errores."
        If Retry <= Maxretry Then
            Flog.writeline Espacios(Tabulador * Nivel) & "Reintento. Paso a Pendiente."
            Flog.writeline Espacios(Tabulador * Nivel) & "Retry: " & Retry & " Maxretry: " & Maxretry
            Retry = Retry + 1
            StrSql = "UPDATE batch_proceso SET bprcprogreso =0 , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Pendiente' "
            StrSql = StrSql & ",bprcparam = '" & CInt(Retry + 1) & "' "
            StrSql = StrSql & ",bprcfecha = " & ConvFecha(DateValue(Now + (Timeretry / 1440))) & " "
            StrSql = StrSql & ",bprchora = '" & FormatDateTime(Now + (Timeretry / 1440), 4) & ":00' "
            StrSql = StrSql & "Where bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Incompleto' WHERE bpronro = " & NroProceso
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    Flog.writeline
    Flog.writeline "Fin :" & Now
    Flog.writeline "-------------------------------------------------"
    Flog.Close
    objConn.Close
    
    Set Flog = Nothing
End
CE:
    HuboErrores = True
    Nivel = Nivel + 1
    'Flog.writeline
    'Flog.writeline Espacios(Tabulador * Nivel) & "-------------------------------------------------"
    Flog.writeline Espacios(Tabulador * Nivel) & " Error: " & Err.Description & " - " & Now
    Flog.writeline Espacios(Tabulador * Nivel) & " Ultimo SQL: " & StrSql
    'Flog.writeline Espacios(Tabulador * Nivel) & "-------------------------------------------------"
    Flog.writeline
    Nivel = Nivel - 1
    GoTo Fin
End Sub

Sub SndMail()
      
    Dim oMail As New CDO.Message
    'Dim oMailConfig As New CDO.Configuration
    Dim tipoArchivo As String
    
    Nivel = Nivel + 1
    Flog.writeline Espacios(Tabulador * Nivel) & "SndMail()"
    
    On Error GoTo ME_SendMail:
    
    'Seteo la configuracion
    'MailHost, MailPort, MailUser, MailPassword
    Nivel = Nivel + 1
    Flog.writeline Espacios(Tabulador * Nivel) & "Configuracion Servidor de Correo"
    
    ' Indica el servidor Smtp para poder enviar el Mail (puede ser el nombre del servidor o su dirección IP)
    oMail.Configuration.Fields(cdoSMTPServer) = MailHost
    
    'Metodo de envio
    'cdoSendUsingPickup (1)
    'cdoSendUsingPort (2)
    'cdoSendUsingExchange (3)
    oMail.Configuration.Fields(cdoSendUsingMethod) = 2
    
    'Puerto del servidor, por defaut 25
    oMail.Configuration.Fields(cdoSMTPServerPort) = CLng(MailPort)
    
    'Tiempo máximo de espera en segundos para la conexión
    oMail.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    
    If EsNulo(MailUser) Or UCase(MailUser) = "ANONYMOUS" Then
        'Anonimo
        oMail.Configuration.Fields(cdoSMTPAuthenticate) = 0
    Else
        'Autentificado
        
        ' Configura las opciones para el login en el SMTP
        oMail.Configuration.Fields(cdoSMTPAuthenticate) = 1
        
        'Id de usuario del servidor Smtp ( por lo general debe ser la dir completa usuario@domino.com )
        oMail.Configuration.Fields(cdoSendUserName) = MailUser
        
        ' Password de la cuenta
        oMail.Configuration.Fields(cdoSendPassword) = MailPassword
        
        ' Indica si se usa SSL para el envío. por ahora NO (puede requierir otro puerto)
        If SSL = 0 Then
            oMail.Configuration.Fields(cdoSMTPUseSSL) = False
        Else
            oMail.Configuration.Fields(cdoSMTPUseSSL) = True
        End If
    End If
    
   
    Flog.writeline Espacios(Tabulador * Nivel) & "Datos del mail:"
    Nivel = Nivel + 1
        
    Flog.writeline Espacios(Tabulador * Nivel) & "From: " & MailFrom
    oMail.From = MailFrom
    
    Flog.writeline Espacios(Tabulador * Nivel) & "To: " & MailTo
    oMail.To = MailTo
    
    Flog.writeline Espacios(Tabulador * Nivel) & "Subject: " & MailSubject
    oMail.Subject = MailSubject
    
    Flog.writeline Espacios(Tabulador * Nivel) & "Cargo el body "
    Flog.writeline Espacios(Tabulador * Nivel) & "Body: " & MailBody
    oMail.HTMLBody = MailBody
        
    If HTMLMailHeader <> "" Then
        Dim HTMLMailHeaderArray() As String
        HTMLMailHeaderArray = Split(HTMLMailHeader, ";")
        Dim objImage As Object
        Dim nomArch As String
        Dim nomArchArray() As String
        Dim a As Integer
        For a = 0 To UBound(HTMLMailHeaderArray) - 1
            nomArchArray = Split(HTMLMailHeaderArray(a), "\")
            nomArch = nomArchArray(UBound(nomArchArray))
            'tipoArchivo = UCase(Split(nomArch, ".")(UBound(Split(nomArch, "."))))
            ' valido si el archivo forma parte del contenido, va como header, sino como attach (lm)
            If InStr(1, MailBody, "src=""" & nomArch & """") Or InStr(1, MailBody, "href=""" & nomArch & """") Or InStr(1, MailBody, "src='" & nomArch & "'") Or InStr(1, MailBody, "href='" & nomArch & "'") Then
                Set objImage = oMail.AddRelatedBodyPart(HTMLMailHeaderArray(a), nomArch, 1) '    CdoReferenceTypeName
                objImage.Fields.Item("urn:schemas:mailheader:Content-ID") = nomArch
                objImage.Fields.Update
            Else
                oMail.AddAttachment (HTMLMailHeaderArray(a))
            End If
        Next a
        Set objImage = Nothing
    End If
        
        
    Flog.writeline Espacios(Tabulador * Nivel) & "Attach: " & MailAttach
    If MailAttach <> "" Then
        oMail.AddAttachment (MailAttach)
    End If
    
    
    Flog.writeline Espacios(Tabulador * Nivel) & "CC: " & MailCC
    oMail.CC = MailCC
    
    Flog.writeline Espacios(Tabulador * Nivel) & "CCO: " & MailBCC
    oMail.BCC = MailBCC
    
    'Actualizo los datos antes de enviar
    oMail.Configuration.Fields.Update
    
    Nivel = Nivel - 1
    Flog.writeline Espacios(Tabulador * Nivel) & "Enviando..."
    oMail.Send

'''''    revisar:
'''''        metodo de envio (a definir, hoy fijo 2 (port))
'''''        SSL (falta modificar asp, por ahora siempre en false)
'''''        Autentificacion (resuelto)
    
    
    Nivel = Nivel + 1
    Flog.writeline Espacios(Tabulador * Nivel) & "Mail enviando..."
    
    'Elimino los archivos del mail
    Nivel = Nivel - 1
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel) & "Elimino los archivos del mail..."
    
    If Not oMail Is Nothing Then
        Set oMail = Nothing
    End If
    
    Call EliminarArchivo(dirsalidas & "\" & f1.Name)
    'FGZ - 03/09/2012 -------------
    'Call EliminarArchivo(MailAttach)
    If Not EsNulo(MailAttach) Then
        Call EliminarArchivo(MailAttach)
    End If
    'FGZ - 03/09/2012 -------------
    If Not EsNulo(MailHtmlBody) Then
        Call EliminarArchivo(MailHtmlBody)
    End If
    
    Nivel = Nivel - 1
    Flog.writeline Espacios(Tabulador * Nivel) & "Fin SndMail()"
    Nivel = Nivel - 1
Exit Sub
ME_SendMail:
    HuboError = True
    Nivel = Nivel + 1
    'Flog.writeline
    'Flog.writeline Espacios(Tabulador * Nivel) & "Error tratando de enviar el mail. Sub sndMail() "
    Flog.writeline Espacios(Tabulador * Nivel) & "Error " & Err.Number & " " & Err.Description
End Sub




Private Sub EliminarArchivo(ByVal FileName As String)
  Dim fso
  
  On Error GoTo ME_Eliminar
  
  Nivel = Nivel + 1
  Set fso = CreateObject("Scripting.FileSystemObject")
  fso.DeleteFile (FileName)
  Flog.writeline Espacios(Tabulador * Nivel) & "Archivo eliminado " & FileName
  
  Nivel = Nivel - 1
  
  Set fso = Nothing
Exit Sub

ME_Eliminar:
    Nivel = Nivel + 1
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel) & "Error tratando de eliminar el archivo: " & Err.Description
    Nivel = Nivel - 1
End Sub
