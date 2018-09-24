Attribute VB_Name = "mdlMsg"
Option Explicit

Public WithEvents oSMTP As OSSMTP.SMTPSession

Dim dirsalidas As String
Dim MailHost As String
Dim MailPort As String
Dim MailFromName As String
Dim MailFrom As String
Dim MailTo As String
Dim MailSubject As String
Dim MailBody As String
Dim MailAttach As String
Dim MailUser As String
Dim MailPassword As String
Dim f1, fc
Dim HuboErrores As Boolean
Dim Maxretry As Integer
Dim Timeretry As Integer
Dim SinParemetros As Boolean

Public Sub CargarMensaje(ByVal Archivo As String)
'-------------------------------------------------
'
'-------------------------------------------------
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer


    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(Archivo, ForReading, 0)
    
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
    End If
    
    Flog.Write "FromName "
    ' FromName
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailFromName = Mid(strline, pos1, pos2)
        Flog.writeline MailFromName
    End If
    
    Flog.Write "Subject "
    ' Subject
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailSubject = Mid(strline, pos1, pos2)
        Flog.writeline MailSubject
    End If
    
    Flog.writeline "Body1 "
    ' Body1
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailBody = Mid(strline, pos1, pos2)
        Flog.writeline MailBody
    End If
    
    Flog.writeline "Attachment "
    ' Attachment
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailAttach = Mid(strline, pos1, pos2)
        Flog.writeline MailAttach
    End If
    
    Flog.writeline "Recipients "
    ' Recipients
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailTo = Mid(strline, pos1, pos2)
        Flog.writeline MailTo
    End If
    
    Flog.writeline "FromAddress "
    ' FromAddress
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailFrom = Mid(strline, pos1, pos2)
        Flog.writeline MailFrom
    End If
    
    Flog.writeline "Host "
    ' Host
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailHost = Mid(strline, pos1, pos2)
        Flog.writeline MailHost
    End If
    
    Flog.writeline "Port "
    ' Port
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailPort = Mid(strline, pos1, pos2)
        Flog.writeline MailPort
    End If
    
    Flog.writeline "User "
    ' User
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailUser = Mid(strline, pos1, pos2)
        Flog.writeline MailUser
    End If
    
    Flog.writeline "Password "
    ' Password
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailPassword = Mid(strline, pos1, pos2)
        Flog.writeline MailPassword
    End If
    
    f.Close
    Flog.writeline "Termina la carga"
   
End Sub

Sub SndMail()
    
    On Error GoTo ME_SendMail:
    Flog.writeline "configurando ..." & Now
    Set oSMTP = New OSSMTP.SMTPSession
    'authentication
    If MailUser <> "" Then
        Flog.writeline "MailUser " & MailUser
        oSMTP.Username = MailUser
        
        Flog.writeline "MailPassword " & MailPassword
        oSMTP.Password = MailPassword
        
        Flog.writeline "AuthenticationType " & 2
        oSMTP.AuthenticationType = 2
    End If
    
    'attachments
    Dim oAttachment As Attachment
   
    'simplified syntax (without Attachment object):
    If MailAttach <> "" Then
        Flog.writeline "MailAttach " & MailAttach
        oSMTP.Attachments.Add MailAttach
    End If
    
    'without CustomHeader object:
    Flog.writeline "MailHost " & MailHost
    oSMTP.Server = MailHost
    
    Flog.writeline "MailPort " & MailPort
    oSMTP.Port = MailPort
    
    Flog.writeline "MailFromName " & MailFrom
    oSMTP.MailFrom = MailFromName & " <" & MailFrom & ">"
    
    Flog.writeline "MailTo " & MailTo
    oSMTP.SendTo = Replace(MailTo, ";", ",")
    '.CC = txtCC
    '.BCC = txtBCC
    Flog.writeline "MailSubject " & MailSubject
    oSMTP.MessageSubject = MailSubject
    
    Flog.writeline "MailBody " & MailBody
    oSMTP.MessageText = MailBody
    
    Flog.writeline "Enviando ..." & Now
    oSMTP.SendEmail
    Set oSMTP = Nothing
    Flog.writeline "Enviado " & Now
    Exit Sub
    
ME_SendMail:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error tratando de enviar el mail " & Err.Description
    
End Sub

Private Sub oSMTP_ErrorSMTP(ByVal Number As Integer, Description As String)
'error occured
Flog.writeline "------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 1) & "1 Mensaje con Error: " & Number & ": " & Description
Flog.writeline
Flog.writeline Espacios(Tabulador * 1) & "FromName " & MailFromName
Flog.writeline Espacios(Tabulador * 1) & "FromAddress " & MailFrom
Flog.writeline Espacios(Tabulador * 1) & "Recipients " & MailTo
Flog.writeline Espacios(Tabulador * 1) & "User " & MailUser
Flog.writeline "------------------------------------------------------------"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub oSMTP_SendSMTP()
  Dim fso
  Dim FileName As String
  
  On Error GoTo ME_Local
  
  Flog.writeline "1 Mensaje enviado OK."
  Set fso = CreateObject("Scripting.FileSystemObject")
  If MailAttach <> "" Then
    Flog.writeline "Borro attach " & MailAttach
    fso.DeleteFile (MailAttach)
  End If
  Flog.writeline "Borro mensaje " & f1.Path
  fso.DeleteFile (f1.Path)
  Set fso = Nothing
  Exit Sub
  
  
ME_Local:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 1) & "Error tratando de eliminar el archivo: " & Err.Description
End Sub

