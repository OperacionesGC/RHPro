Attribute VB_Name = "MdlFechas"
Option Explicit
' ---------------------------------------------------------------------------------------------
' Descripcion:Modulo Fechas. Procedimientos y Funciones de Fechas
' Autor      :FGZ
' Fecha      :05/08/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


Public Sub Dif_Fechas(ByVal FAlta As Date, ByVal FBaja As Date, ByRef dd As Integer, ByRef mm As Integer, ByRef aa As Integer)
' Calcula la diferencia entre dos fechas en Dias, Meses y Años

    dd = DateDiff("d", FAlta, FBaja)
    mm = DateDiff("m", FAlta, FBaja)
    aa = DateDiff("yyyy", FAlta, FBaja)
End Sub



Public Sub DIF_FECHAS2(ByVal F1 As Date, ByVal F2 As Date, ByRef dias As Integer, ByRef Meses As Integer, ByRef anios As Integer)
'/*----------------------------------------------------------------------
'CALCULA LA DIFERENCIA ENTRE DOS FECHAS EN MESES Y A¤OS. f2>f1
'-----------------------------------------------------------------------*/

Dim numdiasmes As Integer
Dim d1 As Date
Dim d2 As Date

d1 = CDate("01/" & Month(F1) & "/" & Year(F1))
Meses = (Month(F1) Mod 12) + 1
anios = Year(F1) + Int(Month(F1) / 12)
d2 = CDate("01/" & Meses & "/" & anios)

numdiasmes = d2 - d1

Meses = 0
anios = 0

dias = Day(F2) - Day(F1)
Meses = Month(F2) - Month(F1)
anios = Year(F2) - Year(F1)
If dias < 0 Then
    Meses = Meses - 1
    dias = dias + numdiasmes
End If
If Meses < 0 Then
    anios = anios - 1
    Meses = Meses + 12
End If

End Sub
