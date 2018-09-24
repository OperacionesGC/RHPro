Attribute VB_Name = "Module1"
Option Explicit

'Public eval As AnalisadorSintactico
'Public eval2 As AnalisadorSintactico
Public gdatServerStarted As Date

Sub Main()
   ' Código que se ejecutará cuando se inicie el
   '   componente, como respuesta a la primera
   '   solicitud de objeto.
   gdatServerStarted = Now
   Debug.Print "Ejecutando Sub Main"
End Sub

' Función para proporcionar identificadores únicos
'    para los objetos.
Public Function GetDebugID() As Long
   Static lngDebugID As Long
   lngDebugID = lngDebugID + 1
   GetDebugID = lngDebugID
End Function


