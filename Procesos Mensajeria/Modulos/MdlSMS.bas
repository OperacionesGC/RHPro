Attribute VB_Name = "MdlSMS"
Option Explicit

Public Sub SendSmsTxt(ByVal toNum As String, ByVal fromNum As String, ByVal toCompany As Integer, ByVal msg As String)
Dim rs_sms As New ADODB.Recordset
Dim Wserv As Boolean
Dim subFijo As String
Dim Para As String

    If Len(toCompany) = 0 Then toCompany = "0"
    
    'Busco cual proceso de sms esta activo
    StrSql = "SELECT cfgsmsnro, cfgsmsdesc, cfgsmsest, cfgsmsaux1, cfgsmsaux2, cfgsmsuser,"
    StrSql = StrSql & " cfgsmspassword, cfgsmscodext, cfgsmsfrom"
    StrSql = StrSql & " FROM conf_sms"
    StrSql = StrSql & " WHERE ((cfgsmsnro = 1)AND (cfgsmsest = -1))"
    StrSql = StrSql & " OR (cfgsmscodext = " & toCompany & ")"
    StrSql = StrSql & " ORDER BY cfgsmsnro"
    OpenRecordset StrSql, rs_sms
        
    If Not rs_sms.EOF Then
        
        'Si el from del procedimiento es vacio todo el de la tabla como default
        If Len(fromNum) <> 0 Then
            Para = fromNum
        Else
            Para = IIf(EsNulo(rs_sms!cfgsmsfrom), "", rs_sms!cfgsmsfrom)
        End If
        
        If (CLng(rs_sms!cfgsmsnro) = 1) Then
            'Envio el sms por Web Service
            Wserv = True
            Call smsWebServ(rs_sms!cfgsmsuser, rs_sms!cfgsmspassword, toNum, toCompany, Para, msg)
        Else
            'Envio el sms como un mail
            subFijo = rs_sms!cfgsmsaux1
        End If
    Else
        Flog.writeline "No se encuentra configurado el envio de SMS."
    End If
    
    If rs_sms.State = adStateOpen Then rs_sms.Close
    Set rs_sms = Nothing
    
End Sub


Public Sub smsWebServ(ByVal user, ByVal password, ByVal toNum, ByVal toCompany, ByVal Para, ByVal msg)
Dim wsClient As New MSSOAPLib30.SoapClient30
Dim objXMLdsPubs As IXMLDOMSelection

    wsClient.MSSoapInit "http://app2.intertronmobile.com/sendmessages/WSMessage.asmx"
    Set objXMLdsPubs = wsClient.Execute("SittInterface.Reintegros", auxi)
    
    Set objXMLdsPubs = Nothing
    Set wsClient = Nothing
End Sub
