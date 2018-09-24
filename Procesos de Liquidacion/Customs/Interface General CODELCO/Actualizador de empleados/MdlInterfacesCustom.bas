Attribute VB_Name = "MdlInterfacesCustom"
Option Explicit
Global datosEsp(0 To 9) As Long
Global datosEltoAna(0 To 1, 0 To 113) As Long

Public Sub Insertar_Linea_Segun_Modelo_Custom(ByVal Linea As String)

' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento llamador de acurdo al modelo
' Autor      : FGZ
' Fecha      : 30/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
' Interfaces Customisadas
Dim Ok As Boolean

MyBeginTrans
    HuboError = False
    Select Case NroModelo
'        Case 216: 'Acumuladores de Agencia para Citrusvil
'            Call LineaModelo_216(Linea)
'        Case 239: 'Interfase Deloitte
'            Call LineaModelo_239(Linea)
'        Case 240: 'LiqPro04
'            Call LineaModelo_240(Linea)
'        Case 241: 'Interfase Dabra
'            Call LineaModelo_241(Linea)
'        Case 247: 'Interfase Acumulado de Horas TELEPERFORMANCE
'            Call LineaModelo_247(Linea)
'        Case 300: 'Interfase de empleados para Teleperformance
'            Ok = True
'            Call LineaModelo_300(Linea, Ok)
    End Select
    
MyCommitTrans
If Not HuboError Then
    Flog.Writeline Espacios(Tabulador * 1) & "Transaccion Cometida"
Else
    Flog.Writeline Espacios(Tabulador * 1) & "Transaccion Abortada"
End If

End Sub

