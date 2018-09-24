Attribute VB_Name = "MdlFechas"
Option Explicit
' ---------------------------------------------------------------------------------------------
' Descripcion:Modulo Fechas. Procedimientos y Funciones de Fechas
' Autor      :FGZ
' Fecha      :05/08/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


Public Sub Dif_Fechas(ByVal FAlta As Date, ByVal FBaja As Date, ByRef dd As Long, ByRef mm As Long, ByRef aa As Long)
' Calcula la diferencia entre dos fechas en Dias, Meses y Años

    dd = DateDiff("d", FAlta, FBaja)
    mm = DateDiff("m", FAlta, FBaja)
    aa = DateDiff("yyyy", FAlta, FBaja)
End Sub



Public Sub DIF_FECHAS2(ByVal F1 As Date, ByVal F2 As Date, ByRef dias As Long, ByRef Meses As Long, ByRef anios As Long)
'/*----------------------------------------------------------------------
'CALCULA LA DIFERENCIA ENTRE DOS FECHAS EN MESES Y A¤OS. f2>f1
'-----------------------------------------------------------------------*/

Dim numdiasmes As Long
Dim d1 As Date
Dim d2 As Date

'd1 = c_date("01/" & Month(F1) & "/" & Year(F1))
'Meses = (Month(F1) Mod 12) + 1
'anios = Year(F1) + Int(Month(F1) / 12)
'd2 = c_date("01/" & Meses & "/" & anios)
'
'numdiasmes = d2 - d1

d1 = C_Date("01/" & Month(F1) & "/" & Year(F1))
Meses = (Month(F1) Mod 12) + 1
anios = Year(F1) + Int(Month(F1) / 12)
d2 = C_Date("01/" & Meses & "/" & anios)

numdiasmes = d2 - d1

Meses = 0
anios = 0

dias = IIf(Day(F2) = 31, 30, Day(F2)) - Day(F1)
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

Public Sub DIF_FECHAS3(ByVal F1 As Date, ByVal F2 As Date, ByRef dias As Long, ByRef Meses As Long, ByRef anios As Long)
'/*----------------------------------------------------------------------
'CALCULA LA DIFERENCIA ENTRE DOS FECHAS EN MESES Y A¤OS. f2>f1
'-----------------------------------------------------------------------*/

Dim numdiasmes As Long
Dim d1 As Date
Dim d2 As Date

'd1 = c_date("01/" & Month(F1) & "/" & Year(F1))
'Meses = (Month(F1) Mod 12) + 1
'anios = Year(F1) + Int(Month(F1) / 12)
'd2 = c_date("01/" & Meses & "/" & anios)
'
'numdiasmes = d2 - d1

d1 = C_Date("01/" & Month(F1) & "/" & Year(F1))
Meses = (Month(F1) Mod 12) + 1
anios = Year(F1) + Int(Month(F1) / 12)
d2 = C_Date("01/" & Meses & "/" & anios)

numdiasmes = d2 - d1

If numdiasmes > 30 Then
   numdiasmes = 30
End If

Meses = 0
anios = 0

dias = (IIf(Day(F2) = 31, 30, Day(F2)) - Day(F1)) + 1
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
If dias = 30 Then
   dias = 0
   Meses = Meses + 1
End If
If Meses = 12 Then
   Meses = 0
   anios = anios + 1
End If

End Sub

Public Sub DIF_FECHAS4(ByVal F1 As Date, ByVal F2 As Date, ByRef dias As Long, ByRef Meses As Long, ByRef anios As Long)
'/*----------------------------------------------------------------------
'CALCULA LA DIFERENCIA ENTRE DOS FECHAS EN MESES Y A¤OS. f2>f1
'-----------------------------------------------------------------------*/

Dim numdiasmes As Long
Dim DiasDelmes As Long
Dim d1 As Date
Dim d2 As Date

d1 = C_Date("01/" & Month(F1) & "/" & Year(F1))
Meses = (Month(F1) Mod 12) + 1
anios = Year(F1) + Int(Month(F1) / 12)
d2 = C_Date("01/" & Meses & "/" & anios)

numdiasmes = d2 - d1
DiasDelmes = numdiasmes

If numdiasmes > 30 Then
   numdiasmes = 30
End If

Meses = 0
anios = 0

dias = (Day(F2)) - Day(F1) + 1
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

'If dias = 30 Then
If dias = DiasDelmes Then
   dias = 0
   Meses = Meses + 1
End If
If Meses = 12 Then
   Meses = 0
   anios = anios + 1
End If

End Sub

Public Sub DIF_FECHAS5(ByVal res As String, ByVal F1 As Date, ByVal F2 As Date, ByRef dias As Long, ByRef Meses As Long, ByRef anios As Long)
'/*----------------------------------------------------------------------
'CALCULA LA DIFERENCIA ENTRE DOS FECHAS EN DIAS, MESES Y AÑOS. f2>f1
'-----------------------------------------------------------------------*/

Dim numdiasmes As Long
Dim DiasDelmes As Long
Dim d1 As Date
Dim d2 As Date

d1 = C_Date("01/" & Month(F1) & "/" & Year(F1))
Meses = (Month(F1) Mod 12) + 1
anios = Year(F1) + Int(Month(F1) / 12)
d2 = C_Date("01/" & Meses & "/" & anios)

numdiasmes = d2 - d1
DiasDelmes = numdiasmes

If numdiasmes > 30 Then
   numdiasmes = 30
End If

Meses = 0
anios = 0

'dias = (Day(F2)) - Day(F1) + 1
dias = (Day(F2)) - Day(F1)
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

Public Sub DIF_FECHAS6(ByVal res As String, ByVal F1 As Date, ByVal F2 As Date, ByRef dias As Long, ByRef Meses As Long, ByRef anios As Long)
'/*----------------------------------------------------------------------
'CALCULA LA DIFERENCIA ENTRE DOS FECHAS EN DIAS, MESES Y AÑOS. f2>f1
'-----------------------------------------------------------------------*/

Dim numdiasmes As Long
Dim DiasDelmes As Long
Dim d1 As Date
Dim d2 As Date

d1 = C_Date("01/" & Month(F1) & "/" & Year(F1))
Meses = (Month(F1) Mod 12) + 1
anios = Year(F1) + Int(Month(F1) / 12)
d2 = C_Date("01/" & Meses & "/" & anios)

numdiasmes = d2 - d1
DiasDelmes = numdiasmes

If numdiasmes > 30 Then
   numdiasmes = 30
End If

Meses = 0
anios = 0

dias = (Day(F2)) - Day(F1) + 1
'Dias = (Day(F2)) - Day(F1)
Meses = Month(F2) - Month(F1)
anios = Year(F2) - Year(F1)
'FGZ - 11/05/2015 -------------
If dias > 30 Then
   dias = 30
End If
'FGZ - 11/05/2015 -------------
If dias < 0 Then
    Meses = Meses - 1
    dias = dias + numdiasmes
End If
If Meses < 0 Then
    anios = anios - 1
    Meses = Meses + 12
End If

End Sub

Public Sub sigMes(ByRef Anio As Long, ByRef Mes As Long)
'/*----------------------------------------------------------------------
'CALCULA El Siguiente Mes dado un mes y un año
'-----------------------------------------------------------------------*/
    If (Mes = 12) Then
        Mes = 1
        Anio = Anio + 1
    Else
        Mes = Mes + 1
    End If
End Sub

Public Sub sigQuin(ByRef Anio As Long, ByRef Mes As Long, ByRef quin As Long)
'/*----------------------------------------------------------------------
'CALCULA la Siguiente quincena dado un mes y un año
'-----------------------------------------------------------------------*/
    If (quin = 2) Then
        quin = 1
        If (Mes = 12) Then
            Mes = 1
            Anio = Anio + 1
        Else
            Mes = Mes + 1
        End If
    Else
        quin = quin + 1
    End If
End Sub


Public Function UltimoDiaMes(ByVal Anio As Integer, ByVal Mes As Integer) As Date
'/*----------------------------------------------------------------------
'CALCULA El ultimo dia del mes y año pasado por parametro
'-----------------------------------------------------------------------*/
Dim Aux_Fecha As Date

    If (Mes = 12) Then
        Mes = 1
        Anio = Anio + 1
    Else
        Mes = Mes + 1
    End If
    
    Aux_Fecha = CDate("01/" & Mes & "/" & Anio) - 1
    UltimoDiaMes = Aux_Fecha
End Function

