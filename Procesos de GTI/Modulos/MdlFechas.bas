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



Public Sub DIF_FECHAS2(ByVal F1 As Date, ByVal F2 As Date, ByRef dias As Integer, ByRef meses As Integer, ByRef anios As Integer)
'/*----------------------------------------------------------------------
'CALCULA LA DIFERENCIA ENTRE DOS FECHAS EN MESES Y A¤OS. f2>f1
'-----------------------------------------------------------------------*/

Dim numdiasmes As Integer
Dim d1 As Date
Dim d2 As Date

d1 = CDate("01/" & Month(F1) & "/" & Year(F1))
meses = Month(F1) Mod 12 + 1
anios = Year(F1) + Int(Month(F1) / 12)
d2 = CDate("01/" & meses & "/" & anios)

numdiasmes = d2 - d1

meses = 0
anios = 0

dias = Day(F2) - Day(F1)
meses = Month(F2) - Month(F1)
anios = Year(F2) - Year(F1)
If dias < 0 Then
    meses = meses - 1
    dias = dias + numdiasmes
End If
If meses < 0 Then
    anios = anios - 1
    meses = meses + 12
End If

End Sub

Function FormatoFecha(ByVal fecha As Date, ByVal formato As String)
'Dada una fecha de entrada y un formato retorna un cadena con la fecha en el formato requerido
Dim salida As String
    
    salida = fecha
    Select Case UCase(formato)
        Case "AAAAMMDD"
            salida = Right(fecha, 4) & Mid(fecha, 4, 2) & Left(fecha, 2)
        Case "AAAA-MM-DD"
            salida = Right(fecha, 4) & "-" & Mid(fecha, 4, 2) & "-" & Left(fecha, 2)
        Case "AAAA/MM/DD"
            salida = Right(fecha, 4) & "/" & Mid(fecha, 4, 2) & "/" & Left(fecha, 2)
        Case "DDMMAAAA"
            salida = Left(fecha, 2) & Mid(fecha, 4, 2) & Right(fecha, 4)
        Case "DD-MM-AAAA"
            salida = Left(fecha, 2) & "-" & Mid(fecha, 4, 2) & "-" & Right(fecha, 4)
        Case "DD/MM/AAAA"
            salida = Left(fecha, 2) & "/" & Mid(fecha, 4, 2) & "/" & Right(fecha, 4)
        Case "MMDDAAAA"
            salida = Mid(fecha, 4, 2) & Left(fecha, 2) & Right(fecha, 4)
        Case "MM-DD-AAAA"
            salida = Mid(fecha, 4, 2) & "-" & Left(fecha, 2) & "-" & Right(fecha, 4)
        Case "MM/DD/AAAA"
            salida = Mid(fecha, 4, 2) & "/" & Left(fecha, 2) & "/" & Right(fecha, 4)
        Case "AAMMDD"
            salida = Right(fecha, 2) & Mid(fecha, 4, 2) & Left(fecha, 2)
        Case "AA-MM-DD"
            salida = Right(fecha, 2) & "-" & Mid(fecha, 4, 2) & "-" & Left(fecha, 2)
        Case "AA/MM/DD"
            salida = Right(fecha, 2) & "/" & Mid(fecha, 4, 2) & "/" & Left(fecha, 2)
        Case "DDMMAA"
            salida = Left(fecha, 2) & Mid(fecha, 4, 2) & Right(fecha, 2)
        Case "DD-MM-AA"
            salida = Left(fecha, 2) & "-" & Mid(fecha, 4, 2) & "-" & Right(fecha, 2)
        Case "DD/MM/AA"
            salida = Left(fecha, 2) & "/" & Mid(fecha, 4, 2) & "/" & Right(fecha, 2)
        Case "MMDDAA"
            salida = Mid(fecha, 4, 2) & Left(fecha, 2) & Right(fecha, 4)
        Case "MM-DD-AA"
            salida = Mid(fecha, 4, 2) & "-" & Left(fecha, 2) & "-" & Right(fecha, 2)
        Case "MM/DD/AA"
            salida = Mid(fecha, 4, 2) & "/" & Left(fecha, 2) & "/" & Right(fecha, 2)
    
    End Select
        FormatoFecha = salida
End Function
