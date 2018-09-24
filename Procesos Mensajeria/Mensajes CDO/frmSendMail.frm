VERSION 5.00
Begin VB.Form Mensajes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Mail Sender"
   ClientHeight    =   135
   ClientLeft      =   525
   ClientTop       =   1230
   ClientWidth     =   2835
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   135
   ScaleWidth      =   2835
End
Attribute VB_Name = "Mensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Const Version = "1.12"
Const FechaVersion = "12/08/2009"   'Fernando Favre
'                               Modificaciones:
'                                   Se modifico el conjunto de caracteres utilizado en la codificacion del mail (charset = "ISO-8859-1").

'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------

Public WithEvents oSMTP As OSSMTP.SMTPSession
Attribute oSMTP.VB_VarHelpID = -1

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
Dim MailBcc As String

Sub CargarMensaje(ByVal Archivo As String)
Flog.writeline "carga las configuraciones basicas para los procesos"
'carga las configuraciones basicas para los procesos
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer

'On Error GoTo CE

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
    
    Flog.writeline "Copia Oculta "
    ' Recipients
    If Not f.AtEndOfStream Then
        strline = f.ReadLine()
        pos1 = InStr(1, strline, "=") + 1
        pos2 = Len(strline)
        MailBcc = Mid(strline, pos1, pos2)
        Flog.writeline MailBcc
    End If
    
    f.Close
    Flog.writeline "Termina la carga"
    
'    Exit Sub
'
'CE:
'    HuboErrores = True
'    Flog.writeline
'    Flog.writeline "-------------------------------------------------"
'    Flog.writeline " Error: " & Err.Description & Now
'    Flog.writeline " Ultimo SQL " & StrSql
'    Flog.writeline "-------------------------------------------------"
'    Flog.writeline
End Sub

Sub CargarDefaults()
Flog.writeline "carga las configuraciones para los mensajes"
' carga las configuraciones para los mensajes
Const ForReading = 1
Const ForAppending = 8

Dim f, fs
Dim strline As String
Dim pos1 As Integer
Dim pos2 As Integer

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
    
    f.Close
    

End Sub

Private Sub Form_Load()

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

    'Nombre_ArchA = App.Path & "Mensajes" & ".log"
    'Set fs = CreateObject("Scripting.FileSystemObject")
    'Set FlogE = fs.CreateTextFile(Nombre_ArchA, True)

    'FlogE.writeline "Arrancamos ..."
    
    Mensajes.Hide
    TiempoInicialProceso = GetTickCount

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
    
    'FlogE.writeline "Creo el log"
    
    HuboErrores = False
    Nombre_Arch = PathFLog & "Mensajes" & "-" & NroProceso & ".log"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Flog = fs.CreateTextFile(Nombre_Arch, True)
    
    'FlogE.writeline "Abro la conexion"
    
    'desactivo momentaneamente
    On Error GoTo 0
    'Abro la conexion
    On Error Resume Next
    OpenConnection strconexion, objConn
    If Err.Number <> 0 Or Error_Encrypt Then
        Flog.writeline "Problemas en la conexion"
        Exit Sub
    End If
    On Error GoTo CE
    
    ' Obtengo el Process ID
    PID = GetCurrentProcessId
    Flog.writeline "-------------------------------------------------"
    Flog.writeline "Version                  : " & Version
    Flog.writeline "Fecha Ultima Modificacion: " & FechaVersion
    Flog.writeline "PID                      : " & PID
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    
    Flog.writeline "Inicio Mensajeria : " & Now
    Flog.writeline "Cambio el estado del proceso a Procesando"
    
    'Cambio el estado del proceso a Procesando
    StrSql = "UPDATE batch_proceso SET bprchorainicioej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecinicioej = " & ConvFecha(Date) & ", bprcestado = 'Procesando', bprcpid = " & PID & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    Flog.writeline "Obtengo los datos del proceso"
    
    Call CargarDefaults
    
    'Obtengo los datos del proceso
    StrSql = "SELECT bprcparam FROM batch_proceso WHERE btprcnro = 25 AND bpronro = " & NroProceso
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
       'Obtengo los parametros del proceso
       If Not EsNulo(objRs!bprcparam) Then
            Retry = objRs!bprcparam
        Else
            Retry = 1
        End If
        Flog.writeline "Reintento nro " & Retry
    Else
        Flog.writeline "Error al buscar datos en el proceso: " & NroProceso
        'Unload Me
        'End
        GoTo Fin
    End If
    If objRs.State = adStateOpen Then objRs.Close
       
    ' Directorio Salidas
    StrSql = "select sis_dirsalidas from sistema"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        dirsalidas = objRs!sis_dirsalidas & "\attach"
        Flog.writeline "Dir Salida: " & dirsalidas
    Else
        Flog.writeline "No se encuentra configurado sis_dirsalidas"
        'Unload Me
        'End
        GoTo Fin
    End If
    If objRs.State = adStateOpen Then objRs.Close
    
    Dim fso, f, s
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFolder(dirsalidas)
    Set fc = f.Files
    If fc.Count = 0 Then
        Flog.writeline "No hay mensajes a enviar."
    End If
    
    'Determino la proporcion de progreso
    Progreso = 0
    CEmpleadosAProc = fc.Count
    If CEmpleadosAProc = 0 Then
        CEmpleadosAProc = 1
    End If
    IncPorc = (99 / CEmpleadosAProc)
    
    'Actualizo el progreso
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & IncPorc & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
    For Each f1 In fc
        'If InStr(f1.Name, ".msg") Then
        'FGZ - 05/09/2006
        If InStr(f1.Name, ".msg") And Left(f1.Name, Len(CStr(NroProceso)) + 5) = "msg_" & CStr(NroProceso) & "_" Then
            Flog.writeline "Cargar Mensaje " & f1.Name
            CargarMensaje (f1.Path)
            SndMail
        End If
    
        'Actualizo el progreso
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    
    Next
    
Fin:
    If Not HuboErrores Then
        'Actualizo el estado del proceso
        StrSql = "UPDATE batch_proceso SET bprcprogreso =100 , bprchorafinej = '" & Format(Now, "hh:mm:ss ") & "', bprcfecfinej = " & ConvFecha(Date) & ", bprcestado = 'Procesado' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        Flog.writeline "Hubo Errores."
        If Retry <= Maxretry Then
            Flog.writeline "Reintento. 6Paso a Pendiente."
            Flog.writeline "Retry: " & Retry & " Maxretry: " & Maxretry
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
    
    
    Flog.writeline "Fin :" & Now
    Flog.Close
    'FlogE.writeline "Terminó"
    'FlogE.Close
    objConn.Close
    
    Set Flog = Nothing
    'Set FlogE = Nothing
End
CE:
    HuboErrores = True
    Flog.writeline
    Flog.writeline "-------------------------------------------------"
    Flog.writeline " Error: " & Err.Description & " - " & Now
    Flog.writeline " Ultimo SQL " & StrSql
    Flog.writeline "-------------------------------------------------"
    Flog.writeline
    GoTo Fin
End Sub

Sub SndMail()
    
    On Error GoTo ME_SendMail:
    'Flog.writeline "Send Mail"
    'FGZ - 01/09/2006
    Flog.writeline "configurando ..." & Now
    Set oSMTP = New OSSMTP.SMTPSession
    oSMTP.Charset = "ISO-8859-1"
    
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
'    Set oAttachment = New Attachment
'      oAttachment.FilePath = MailAttach
      'oAttachment.AttachmentName = ""
'      oAttachment.ContentType = "text/html"
'      oAttachment.ContentTransferEncoding = encBase64
'      oSMTP.Attachments.Add oAttachment
    
    'simplified syntax (without Attachment object):
    If MailAttach <> "" Then
        Flog.writeline "MailAttach " & MailAttach
        oSMTP.Attachments.Add MailAttach
    End If
    
    'without CustomHeader object:
    '.CustomHeaders.Add "Reply-To: errors@mydomain.com"
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
    '.MessageHTML = txtMessageHTML
    
    Flog.writeline "MailBody " & MailBody
    oSMTP.BCC = MailBody
    
    Flog.writeline "MailBcc " & MailBcc
    oSMTP.BCC = Replace(MailBcc, ";", ",")
    
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
  'HuboErrores = True
Flog.writeline "------------------------------------------------------------"
Flog.writeline Espacios(Tabulador * 1) & "1 Mensaje con Error: " & Number & ": " & Description
Flog.writeline
Flog.writeline Espacios(Tabulador * 1) & "FromName " & MailFromName
Flog.writeline Espacios(Tabulador * 1) & "FromAddress " & MailFrom
Flog.writeline Espacios(Tabulador * 1) & "Recipients " & MailTo
Flog.writeline Espacios(Tabulador * 1) & "User " & MailUser
Flog.writeline Espacios(Tabulador * 1) & "Bcc " & MailBcc
Flog.writeline "------------------------------------------------------------"

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If Not SinParemetros Then
'        Flog.writeline "Fin :" & Now
'        Flog.Close
'    End If
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




