VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Web Service con VB6"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   3450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reintegros"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Licencias Estados"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Horas Pactadas"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim empresa As String
Dim FechaInicio As String
Dim FechaFin As String
Dim Parametros() As String
Dim Linea As String

Private Sub Command1_Click()
    ' Licencias Estados
    Dim wsClient As New MSSOAPLib30.SoapClient30
    Dim auxi As String
    Dim objXMLdsPubs As IXMLDOMSelection
    
    ' Conecxion al Web Service
    wsClient.MSSoapInit "http://192.168.106.17/WSSittMGOExt/rhpro.asmx?WSDL", "RHPro", "RHProSoap"
    'wsClient.MSSoapInit "http://www.sittnet.com:8000/wssittmgoext/rhpro.asmx?WSDL", "RHPro", "RHProSoap"

    ' Parametros auxiliares del Servicio
    auxi = "<Params>" & _
    "<Empresa>" & empresa & "</Empresa>" & _
    "<FechaDesde>" & FechaInicio & "</FechaDesde>" & _
    "<FechaHasta>" & FechaFin & "</FechaHasta>" & _
    "</Params>"

    ' Executa el metodo execute con la accion SittInterface.LicenciasEstados
    ' y parametros de la variable auxi

    Set objXMLdsPubs = wsClient.Execute("SittInterface.LicenciasEstados", auxi)

    ' Verifica si el resultado del envio de datos fue OK
    If objXMLdsPubs.Item(0).selectNodes("Result").Item(0).selectSingleNode("IsOk").Text = "true" Then

        Set fs = CreateObject("Scripting.FileSystemObject")
        Set FLog = fs.CreateTextFile(".\LicenciasEstados.csv", True)
        
        FLog.writeline "Empresa;Legajo;Estado;Fecha"
    
       ' Recorre todo el XML correspondiente solo a los Datos
       For i = 0 To objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").length - 1
           auxi = objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(i).selectSingleNode("Empresa").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(i).selectSingleNode("Legajo").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(i).selectSingleNode("Estado").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("LicenciasEstados").Item(i).selectSingleNode("Fecha").Text
           FLog.writeline auxi
       Next i
       FLog.Close
       Set FLog = Nothing
       Set fs = Nothing
       MsgBox "Licencias Estados Exportados"
    Else
        MsgBox "Error: " & objXMLdsPubs.Item(0).xml & " Parametros: " & auxi
    End If
    Set objXMLdsPubs = Nothing
    Set wsClient = Nothing
End Sub

Private Sub Command2_Click()
    ' Horas Pactadas
    Dim wsClient As New MSSOAPLib30.SoapClient30
    Dim auxi As String
    Dim objXMLdsPubs As IXMLDOMSelection
    
    ' Conecxion al Web Service
    wsClient.MSSoapInit "http://192.168.106.17/WSSittMGOExt/rhpro.asmx?WSDL", "RHPro", "RHProSoap"
    'wsClient.MSSoapInit "http://www.sittnet.com:8000/wssittmgoext/rhpro.asmx?WSDL", "RHPro", "RHProSoap"

    ' Parametros auxiliares del Servicio
    auxi = "<Params>" & _
    "<Empresa>" & empresa & "</Empresa>" & _
    "<FechaDesde>" & FechaInicio & "</FechaDesde>" & _
    "<FechaHasta>" & FechaFin & "</FechaHasta>" & _
    "</Params>"

    ' Executa el metodo execute con la accion SittInterface.HorasPactadas
    ' y parametros de la variable auxi
    Set objXMLdsPubs = wsClient.Execute("SittInterface.HorasPactadas", auxi)

    ' Verifica si el resultado del envio de datos fue OK
    If objXMLdsPubs.Item(0).selectNodes("Result").Item(0).selectSingleNode("IsOk").Text = "true" Then

        Set fs = CreateObject("Scripting.FileSystemObject")
        Set FLog = fs.CreateTextFile(".\HorasPactadas.csv", True)
        
        FLog.writeline "Empresa;Dia;Legajo;HorasPactadas"
       
    
       ' Recorre todo el XML correspondiente solo a los Datos
       For i = 0 To objXMLdsPubs.Item(0).selectNodes("HorasPactadas").length - 1
           auxi = objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(i).selectSingleNode("Empresa").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(i).selectSingleNode("Dia").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(i).selectSingleNode("Legajo").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("HorasPactadas").Item(i).selectSingleNode("HorasPactadas").Text
           FLog.writeline auxi
       Next i
       FLog.Close
       Set FLog = Nothing
       Set fs = Nothing
       MsgBox "Horas Pactadas Exportadas"
    Else
        MsgBox "Error: " & objXMLdsPubs.Item(0).xml & " Parametros: " & auxi
    End If
    Set objXMLdsPubs = Nothing
    Set wsClient = Nothing
End Sub

Private Sub Command3_Click()
    ' Reintegros
    Dim wsClient As New MSSOAPLib30.SoapClient30
    Dim auxi As String
    Dim objXMLdsPubs As IXMLDOMSelection
    
    ' Conecxion al Web Service
    wsClient.MSSoapInit "http://192.168.106.17/WSSittMGOExt/rhpro.asmx?WSDL", "RHPro", "RHProSoap"
    'wsClient.MSSoapInit "http://www.sittnet.com:8000/wssittmgoext/rhpro.asmx?WSDL", "RHPro", "RHProSoap"

    ' Parametros auxiliares del Servicio
    auxi = "<Params>" & _
    "<Empresa>" & empresa & "</Empresa>" & _
    "<FechaDesde>" & FechaInicio & "</FechaDesde>" & _
    "<FechaHasta>" & FechaFin & "</FechaHasta>" & _
    "</Params>"

    ' Executa el metodo execute con la accion SittInterface.Reintegros
    ' y parametros de la variable auxi
    Set objXMLdsPubs = wsClient.Execute("SittInterface.Reintegros", auxi)

    ' Verifica si el resultado del envio de datos fue OK
    If objXMLdsPubs.Item(0).selectNodes("Result").Item(0).selectSingleNode("IsOk").Text = "true" Then

       Set fs = CreateObject("Scripting.FileSystemObject")
       Set FLog = fs.CreateTextFile(".\Reintegros.csv", True)
        
       FLog.writeline "Empresa;Dia;ReintegroNro;Legajo;FaltantesRendicion;Total"
    
       ' Recorre todo el XML correspondiente solo a los Datos
       For i = 0 To objXMLdsPubs.Item(0).selectNodes("Reintegros").length - 1
           auxi = objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("Empresa").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("Dia").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("ReintegoNumero").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("Legajo").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("FaltantesRendicion").Text & ";"
           auxi = auxi & objXMLdsPubs.Item(0).selectNodes("Reintegros").Item(i).selectSingleNode("Total").Text
           FLog.writeline auxi
       Next i
       FLog.Close
       Set FLog = Nothing
       Set fs = Nothing
       MsgBox "Reintegros Exportados"
    Else
        MsgBox "Error: " & objXMLdsPubs.Item(0).xml & " Parametros: " & auxi
    End If
    Set objXMLdsPubs = Nothing
    Set wsClient = Nothing
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()

Linea = Command()
'Linea = "POP@03/10/2005@03/11/2005"
Parametros = Split(Linea, "@")
empresa = Parametros(0)
FechaInicio = Parametros(1)
FechaFin = Parametros(2)

End Sub
