Attribute VB_Name = "mdlMail"
Option Explicit

Global MailHost As String
Global MailPort As String
Global MailFromName As String
Global MailFrom As String
Global MailTo As String
Global MailSubject As String
Global MailBody As String
Global MailHtmlBody As String
Global HTMLMailHeader As String
Global MailAttach As String
Global MailUser As String
Global MailPassword As String
Global MailBCC As String
Global MailCC As String



Sub CargarMensaje(ByVal Texto As String)
'carga las configuraciones basicas para los procesos

Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim Encontro As Boolean


    NotificaError = False
    ErrorNotificado = False
    
    MailFromName = ""   'From Name
    MailSubject = ""    'Asunto
    MailBody = ""       'Cuerpo del mail
    MailTo = ""         'Para
    MailCC = ""         'Copia
    MailBCC = ""        'Copia Oculta
    MailFrom = ""       'FromAddress
    MailHost = ""       'Host
    MailPort = ""       'Port
    MailUser = ""       'Usuario
    MailPassword = ""   'Pass
    MailHtmlBody = ""   'MailHtmlBody
    HTMLMailHeader = "" 'HTMLMailHeader
    MailAttach = ""     'Adjuntos
    
'Ejemplo
'FromName = [RHPro Saas]
'Subject = [RHDESA.DES_R3 - Servicio Detenido por error]
'Body = []
'TO=[fzwenger@rhpro.com]
'CC = [lmoro@rhpro.com; emargiotta@rhpro.com]
'CCO = []
'FromAddress = [rhpro@lisandromoro.com.ar]
'Host = [mail.lisandromoro.com.ar]
'Port = [25]
'USER = [rhpro@lisandromoro.com.ar]
'Password = [suipacha]
'HTMLBody = []
'HTMLMailHeader = []
'Attachment = []
       
    
    'On Error Resume Next
    On Error GoTo ME_Notif
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(App.path & "\notificacion.ini", ForReading, 0)
    
    
    'MailFromName
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailFromName = Mid(strline, pos1, pos2 - pos1)
    Else
        Flog.writeline
        Flog.writeline "No hay archivo de notificacion. Si hay errores no podran ser notificados"
    End If
    
    'MailSubject
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailSubject = Mid(strline, pos1, pos2 - pos1)
    End If
    
    'MailBody
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailBody = Mid(strline, pos1, pos2 - pos1)
    End If
    
    'MailTo
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailTo = Mid(strline, pos1, pos2 - pos1)
    End If

    'MailCC
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailCC = Mid(strline, pos1, pos2 - pos1)
    End If

    'MailBCC
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailBCC = Mid(strline, pos1, pos2 - pos1)
    End If


    'MailfromAdress
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailFrom = Mid(strline, pos1, pos2 - pos1)
    End If

    'MailHost
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailHost = Mid(strline, pos1, pos2 - pos1)
    End If

    'MailPort
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailPort = Mid(strline, pos1, pos2 - pos1)
    End If

    'MailUser
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailUser = Mid(strline, pos1, pos2 - pos1)
    End If

    'MailPassword
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailPassword = Mid(strline, pos1, pos2 - pos1)
    End If
    
    'MailHtmlBody
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailHtmlBody = Mid(strline, pos1, pos2 - pos1)
    End If
    
    'HTMLMailHeader
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        HTMLMailHeader = Mid(strline, pos1, pos2 - pos1)
    End If

    'MailAttach
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "[") + 1
        pos2 = InStr(1, strline, "]")
        MailAttach = Mid(strline, pos1, pos2 - pos1)
    End If
   
    
    f.Close
    NotificaError = True
Exit Sub
ME_Notif:
    Flog.writeline "No se puede levantar Notificacion.ini."
    NotificaError = False
End Sub




Sub SndMail(ByVal Texto As String)
      
    Dim oMail As New CDO.Message
    Dim tipoArchivo As String
    
    'On Error GoTo ME_SendMail:
    
   
    'Parametros dinamicos
    'MailSubject = Etiqueta & MailSubject
    'MailBody = MailBody & Texto
    MailBody = Texto
   
   
    'Seteo la configuracion
    'MailHost, MailPort, MailUser, MailPassword
   
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
        oMail.Configuration.Fields(cdoSMTPUseSSL) = False

    End If
   
    oMail.Sender = MailFromName
    oMail.From = MailFrom
    oMail.To = MailTo
    oMail.Subject = MailSubject
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
        
    If MailAttach <> "" Then
        oMail.AddAttachment (MailAttach)
    End If
    
    oMail.CC = MailCC
    oMail.BCC = MailBCC
    
    
    'Actualizo los datos antes de enviar
    oMail.Configuration.Fields.Update
    
    oMail.Send

Exit Sub
ME_SendMail:
    Flog.writeline "Error. No se pudo notitificar "
End Sub




Sub NotificarError(ByVal Asunto As String, ByVal Texto As String)
Dim Cuerpo As String

    On Error GoTo Me_Noti
    '---------------------------------------------------------
    'Mensaje
    MailSubject = Etiqueta & " - " & Asunto
    
    Cuerpo = Chr(13) + Chr(10)
    Cuerpo = Cuerpo & Texto & Chr(13) + Chr(10)
    Cuerpo = Cuerpo & Chr(13) + Chr(10)
    Cuerpo = Cuerpo & Chr(13) + Chr(10)
    Cuerpo = Cuerpo & Chr(13) + Chr(10)
    'Cuerpo = Cuerpo & "Atte. Administrador " & Chr(13) + Chr(10)
    '---------------------------------------------------------
    
    SndMail (Cuerpo)
    Exit Sub
Me_Noti:
    Flog.writeline "Error. No se pudo notitificar."
End Sub
