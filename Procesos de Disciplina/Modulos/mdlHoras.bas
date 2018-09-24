Attribute VB_Name = "mdlHoras"
Option Explicit


Public Sub ValorenHs(ByVal Dias As Integer, ByVal Horas As Integer, ByVal Min As Integer, ByRef ValHora As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna un string con la cantidad de hs y minutos
' Autor      : FGZ
' Fecha      : 05/11/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim total As Integer
Dim cantdias  As Integer
Dim canthoras As Integer
Dim Dia   As Integer '  cantidad de minutos en un dia
Dim Hora As Integer   ' cantidad de minutos en una hora

    canthoras = Dias * 24 + Horas
    ValHora = Format(CStr(canthoras), "#####00") & ":" & Format(CStr(Min), "00")
End Sub

Public Function CHoras(ByVal Cantidad As Single, ByVal Dur As Single) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna un string con la cantidad de hs y minutos a partir de un valor decimal
' Autor      : FGZ
' Fecha      : 09/11/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Minutos As Single
Dim Horas As Single
    If Dur = 0 Then
        Dur = 60
    End If
    
    Cantidad = Cantidad * Dur
    Horas = Int(Cantidad / Dur)
    Minutos = Cantidad Mod Dur
    CHoras = "'" & Format(Horas, "#####00") & ":" & Format(Minutos, "00") & "'"
End Function

Public Function CHorasSF(ByVal Cantidad As Single, ByVal Dur As Single) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna un string con la cantidad de hs y minutos a partir de un valor decimal
' Autor      : FGZ
' Fecha      : 09/11/2009
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Minutos As Single
Dim Horas As Single
    If Dur = 0 Then
        Dur = 60
    End If
    
    Cantidad = Cantidad * Dur
    Horas = Int(Cantidad / Dur)
    Minutos = Cantidad Mod Dur
    CHorasSF = Format(Horas, "#####00") & ":" & Format(Minutos, "00")
End Function



Public Sub Redondeo_enHorasMinutos(ByVal HoraIni As String, ByVal Tip_Red As Integer, ByVal Duracion As Single, ByRef HoraR As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que redondea la cantidad de hs y minutos segun el redondeo de minutos configurado en horas y minutos.
' Autor      : FGZ
' Fecha      :  05/11/2009
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim objRs As New ADODB.Recordset
Dim Horas As Integer
Dim Minutos As Integer
Dim Aux As Single

If Duracion = 0 Then
    Duracion = 60
End If
   
Horas = Int(Mid(HoraIni, 1, 2))
Minutos = Int(Mid(HoraIni, 4, 2))
StrSql = "SELECT * FROM tipredondeo INNER JOIN tipreddet ON tipredondeo.trdnro = tipreddet.trdnro WHERE (tipredondeo.trdnro = " & Tip_Red & ") AND "
StrSql = StrSql & "(tipreddet.trdetdesde <= " & Minutos & ") AND "
StrSql = StrSql & "(tipreddet.trdethasta >= " & Minutos & ")"
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    Minutos = objRs!trdetvalor
    
    If Minutos >= Duracion Then
        Horas = Horas + 1
        Minutos = Minutos - Duracion
    End If
End If
HoraR = Format(CStr(Horas), "#####00") & ":" & Format(CStr(Minutos), "00")
    
If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing
End Sub




Public Sub SHoras(ByVal Hora1 As String, ByVal Hora2 As String, ByRef HoraR As String)
'------------------------------------------------------------------------------------------
'Descripcion:   Suma dos horas expresadas en HH:mm
'Autor:         FGZ
'Fecha:         11/11/2009
'Ult. Modif:
'------------------------------------------------------------------------------------------
Dim HH1
Dim HH2

Dim HR As Long   'cantidad de horas del resultado
Dim MR As Long   'cantidad de minutos del resultado

    HH1 = Split(Hora1, ":")
    HH2 = Split(Hora2, ":")
    HR = CLng(HH1(0)) + CLng(HH2(0))
    MR = CLng(HH1(1)) + CLng(HH2(1))
    
    If MR >= 60 Then
        HR = HR + 1
        MR = MR - 60
    End If
    HoraR = Format(CStr(HR), "#####00") & ":" & Format(CStr(MR), "00")
End Sub

Public Sub RHoras(ByVal Hora1 As String, ByVal Hora2 As String, ByRef HoraR As String)
'------------------------------------------------------------------------------------------
'Descripcion:   Resta dos horas expresadas en HH:mm
'Autor:         FGZ
'Fecha:         11/11/2009
'Ult. Modif:
'------------------------------------------------------------------------------------------
Dim HH1
Dim HH2

Dim HR As Long   'cantidad de horas del resultado
Dim MR As Long   'cantidad de minutos del resultado

    HH1 = Split(Hora1, ":")
    HH2 = Split(Hora2, ":")
    HR = CLng(HH1(0)) + CLng(HH2(0))
    MR = CLng(HH1(1)) + CLng(HH2(1))
    
    'Si los minutos de r1 son menores que los de r2 ==>
    If HH1(1) < HH2(1) Then
        HH1(1) = HH1(1) + 60
        HH1(0) = HH1(0) - 1
    End If
    MR = CLng(HH1(1)) - CLng(HH2(1))
    HR = CLng(HH1(0)) - CLng(HH2(0))

    HoraR = Format(CStr(HR), "#####00") & ":" & Format(CStr(MR), "00")
End Sub


