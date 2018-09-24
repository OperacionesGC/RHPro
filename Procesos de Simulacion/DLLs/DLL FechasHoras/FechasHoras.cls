VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FechasHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Property Set Conexion(ByRef cn As ADODB.Connection)
    Set objConn = cn
End Property


Public Sub Convertir_A_Hora(minutos As Long, ByRef hora As String)
Dim horas As Single
    
    horas = Int(minutos / 60)
    minutos = minutos Mod 60
    hora = Format(horas, "00") & Format(minutos, "00")

End Sub


Public Sub RestaXHoras(ByVal fecini As Date, ByVal horaini As String, ByVal horamas As String, ByRef Mifecha As Date, ByRef MiHora As String)
'Mas conocido como gtir1h01
' Le resta "horamas" horas a (fecini, horaini)
Dim totminutos As Long
Dim minutos As Long
Dim minutosmas As Long
Dim cantdias  As Long
Dim canthoras As Long
Const dia = 1440  '/* cantidad de minutos en un dia   */
Const hora = 60   '/* cantidad de minutos en una hora */

    minutos = Val(Mid(horaini, 1, 2)) * hora + _
                Val(Mid(horaini, 3, 2))
    minutosmas = Val(Mid(horamas, 1, 2)) * hora + _
                Val(Mid(horamas, 3, 2))
    totminutos = (minutos - minutosmas)
    Mifecha = fecini
    Do While (totminutos < 0)
    '      Voy restando de a un dia a la fecha inicial
        Mifecha = DateAdd("d", -1, Mifecha)
        totminutos = totminutos + dia
    Loop
    'Hay que pasar los minutos a horas y minutos
    Do While (totminutos >= hora)
        canthoras = canthoras + 1
        totminutos = totminutos - hora
    Loop
    MiHora = Format(canthoras, "00") & ":" & Format(totminutos, "00")
    
End Sub

Public Sub RestaEntreHoras(ByVal fecini As Date, ByVal horaini As String, ByVal horamas As String, Mifecha As Date, MiHora As String)
'Mas conocido como gtir2f01
'Probar
Dim totminutos As Long
Dim minutos As Long
Dim minutosmas As Long
Dim cantdias  As Long
Dim canthoras As Long
Const dia = 1440  '/* cantidad de minutos en un dia   */
Const hora = 60   '/* cantidad de minutos en una hora */

minutos = Val(Mid(horaini, 1, 2)) * hora + _
            Val(Mid(horaini, 3, 2))
minutosmas = Val(Mid(horamas, 1, 2)) * hora + _
            Val(Mid(horamas, 3, 2))
totminutos = (minutos + minutosmas)
Mifecha = fecini
Do While (totminutos > dia)
'      fecini = fecini + 1
    Mifecha = DateAdd("d", 1, Mifecha)
    totminutos = totminutos - dia
Loop
MiHora = Format(Str(totminutos * hora), "HH:MM")
End Sub


Public Function ValidarHora(ByRef hora As String) As Boolean
    Dim hs As String
    Dim min As String
        
        ValidarHora = False
        hs = Left(hora, 2)
        If Mid(hora, 3, 1) = ":" Then min = Mid(hora, 4) Else min = Mid(hora, 3)
        If IsNumeric(hs) And IsNumeric(min) Then
            If Int(hs) <= 24 And Int(min) <= 59 And Int(hs & min) <= 2400 Then
                hora = hs & min
                ValidarHora = True
            End If
        End If

End Function

Public Sub SumoHoras(fecini As Date, horaini As String, horamas As String, ByRef Mifecha, ByRef MiHora)
'Mas conocido como gtis1h01
'Suma de Horas
'La hora debe estar en formato 9999
Const hora = 60
Const dia = 1440
Dim totminutos As Long
Dim minutos As Long
Dim minutosmas As Long
Dim canthoras As Long

Dim fecha_aux As Date

    fecha_aux = fecini
    minutos = (Val(Mid(horaini, 1, 2)) * hora + _
                Val(Mid(horaini, 3, 2)))
    minutosmas = (Val(Mid(horamas, 1, 2)) * hora + _
                Val(Mid(horamas, 3, 2)))
    totminutos = (minutos + minutosmas)
    Do While (totminutos >= dia)
    '      fecini = fecini + 1
        fecha_aux = DateAdd("d", 1, fecha_aux)
        totminutos = totminutos - dia
    Loop
    'Hay que pasar los minutos a horas y minutos
    Do While (totminutos >= hora)
    '      fecini = fecini + 1
        canthoras = canthoras + 1
        totminutos = totminutos - hora
    Loop
    Mifecha = fecha_aux
    MiHora = Format(canthoras, "00") & ":" & Format(totminutos, "00")
    
End Sub


Public Sub Redondeo_Horas_Tipo(horaini As String, tip_red As Long, ByRef hora1 As Single)
  Dim StrSql As String
  Dim objRs As New ADODB.Recordset
  Dim minutos As Long
  Const dia = 1440
  Const hora = 60

  minutos = Int(Mid(horaini, 3, 2))

  StrSql = "SELECT * FROM tipredondeo INNER JOIN tipreddet ON tipredondeo.trdnro = tipreddet.trdnro WHERE (tipredondeo.trdnro = " & tip_red & ") AND " & _
           "(tipreddet.trdetdesde <= " & minutos & ") AND " & _
           "(tipreddet.trdethasta >= " & minutos & ")"
  OpenRecordset StrSql, objRs
  If Not objRs.EOF Then
      minutos = (Int(Mid(horaini, 1, 2)) * hora) + objRs!trdetvalor
      hora1 = minutos / hora
  Else
      minutos = (Int(Mid(horaini, 1, 2)) * hora) + Int(Mid(horaini, 3, 2))
      hora1 = minutos / hora
  End If
  
  If objRs.State = adStateOpen Then objRs.Close
  Set objRs = Nothing
  
End Sub


Public Function FraccionaHs(hora_inicio As String, tipo_frac As Long) As String
    Dim minutos As Long
    Dim dia As Long
    Dim hora As Long
    Dim result As String
    Dim dummy As Boolean
    
    dia = 1440
    hora = 60
    minutos = Int(Mid(hora_inicio, 3, 2))
    
    StrSql = "SELECT * FROM tipredondeo INNER JOIN tipreddet ON tipreddet.trdnro = tipredondeo.trdnro " & _
             " WHERE  (tipredondeo.trdnro = " & tipo_frac & " AND tipreddet.trdetdesde <= " & minutos & " ) AND " & _
             " (tipreddet.trdethasta >= " & minutos & ")"
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        If objRs!trdetvalor >= 60 Then
            result = Format(Mid(hora_inicio, 1, 2) + Int(objRs!trdetvalor / 60), "00") & Format(Int(objRs!trdetvalor Mod 60), "00")
        Else
            result = Format(Mid(hora_inicio, 1, 2), "00") & Format(objRs!trdetvalor, "00")
        End If
    Else
        result = hora_inicio
    End If
    If result = "00:00" And minutos = 1440 Then result = "24:00"
    dummy = ValidarHora(result)
    FraccionaHs = result
       
End Function

Public Sub RestaHs(Fecha_inicio As Date, hora_inicio As String, Fecha_Fin As Date, hora_fin As String, ByRef TotDias, ByRef tothoras As Long, ByRef TotMin As Long)
Dim total As Long
Dim cantdias  As Long
Dim canthoras As Long
Dim dia   As Long '  cantidad de minutos en un dia
Dim hora As Long   ' cantidad de minutos en una hora

    dia = 1440
    hora = 60
    canthoras = (Int(Mid(hora_fin, 1, 2)) * hora + _
                   Int(Mid(hora_fin, 3, 2))) - _
                  (Int(Mid(hora_inicio, 1, 2)) * hora + _
                   Int(Mid(hora_inicio, 3, 2)))
    cantdias = DateDiff("d", Fecha_inicio, Fecha_Fin) * dia
    
    total = cantdias + canthoras
    TotDias = Int(total / dia)
    tothoras = Int((total Mod dia) / hora)
    TotMin = (total Mod hora)
End Sub




