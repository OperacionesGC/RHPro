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


Public Sub Convertir_A_Hora(ByVal Minutos As Long, ByRef Hora As String)
'FGZ - 03/11/2008 - puse el parametro Minutos por valor porque estaba por referencia (si no se poner byval por default toma byref)
Dim Horas As Single
    
    Horas = Int(Minutos / 60)
    Minutos = Minutos Mod 60
    Hora = Format(Horas, "000") & Format(Minutos, "00")

End Sub

Public Sub Convertir_A_Hora_cDuracion(ByVal Minutos As Single, ByVal Dur As Single, ByRef Hora As String)
'FGZ - 03/11/2008 - puse el parametro Minutos por valor porque estaba por referencia (si no se poner byval por default toma byref)
'                   Ademas le agregué la duracion de la hora
Dim Horas As Single
    
    If Dur = 0 Then
        Dur = 60
    End If

    Minutos = Minutos * Dur
    Horas = Int(Minutos / Dur)
    Minutos = Minutos Mod Dur
    'Hora = Format(Horas, "00") & Format(Minutos, "00")
    Hora = Format(Horas, "000") & Format(Minutos, "00")
End Sub

Public Sub Convertir_A_Hora_cDuracion2(ByVal Minutos As Single, ByVal Dur As Single, ByRef Hora As String)
'FGZ - 03/11/2008 - puse el parametro Minutos por valor porque estaba por referencia (si no se poner byval por default toma byref)
'                   Ademas le agregué la duracion de la hora
Dim Horas As Single
    
    If Dur = 0 Then
        Dur = 60
    End If


    Minutos = Minutos * Dur

    Horas = Int(Minutos / Dur)
    Minutos = Minutos Mod Dur
    Hora = Format(Horas, "00") & Format(Minutos, "00")
End Sub


Public Sub Convertir_A_Hora2(ByVal Minutos As Long, ByRef Hora As String)
Dim Horas As Single
    Horas = Int(Minutos / 60)
    Minutos = Minutos Mod 60
    Hora = Format(Horas, "00") & Format(Minutos, "00")
End Sub


Public Sub RestaXHoras(ByVal fecini As Date, ByVal HoraIni As String, ByVal horamas As String, ByRef Mifecha As Date, ByRef MiHora As String)
'Mas conocido como gtir1h01
' Le resta "horamas" horas a (fecini, horaini)
Dim totminutos As Integer
Dim Minutos As Integer
Dim minutosmas As Integer
Dim cantdias  As Integer
Dim canthoras As Integer
Const Dia = 1440  '/* cantidad de minutos en un dia   */
Const Hora = 60   '/* cantidad de minutos en una hora */

    Minutos = Val(Mid(HoraIni, 1, 2)) * Hora + _
                Val(Mid(HoraIni, 3, 2))
    minutosmas = Val(Mid(horamas, 1, 2)) * Hora + _
                Val(Mid(horamas, 3, 2))
    totminutos = (Minutos - minutosmas)
    Mifecha = fecini
    Do While (totminutos < 0)
    '      Voy restando de a un dia a la fecha inicial
        Mifecha = DateAdd("d", -1, Mifecha)
        totminutos = totminutos + Dia
    Loop
    'Hay que pasar los minutos a horas y minutos
    Do While (totminutos >= Hora)
        canthoras = canthoras + 1
        totminutos = totminutos - Hora
    Loop
    MiHora = Format(canthoras, "00") & ":" & Format(totminutos, "00")
    
End Sub

Public Sub RestaEntreHoras(ByVal fecini As Date, ByVal HoraIni As String, ByVal horamas As String, Mifecha As Date, MiHora As String)
'Mas conocido como gtir2f01
'Probar
Dim totminutos As Integer
Dim Minutos As Integer
Dim minutosmas As Integer
Dim cantdias  As Integer
Dim canthoras As Integer
Const Dia = 1440  '/* cantidad de minutos en un dia   */
Const Hora = 60   '/* cantidad de minutos en una hora */

Minutos = Val(Mid(HoraIni, 1, 2)) * Hora + _
            Val(Mid(HoraIni, 3, 2))
minutosmas = Val(Mid(horamas, 1, 2)) * Hora + _
            Val(Mid(horamas, 3, 2))
totminutos = (Minutos + minutosmas)
Mifecha = fecini
Do While (totminutos > Dia)
'      fecini = fecini + 1
    Mifecha = DateAdd("d", 1, Mifecha)
    totminutos = totminutos - Dia
Loop
MiHora = Format(Str(totminutos * Hora), "HH:MM")
End Sub


Public Function ValidarHora(ByRef Hora As String) As Boolean
    Dim hs As String
    Dim Min As String
        
        ValidarHora = False
        hs = Left(Hora, 2)
        If Mid(Hora, 3, 1) = ":" Then Min = Mid(Hora, 4) Else Min = Mid(Hora, 3)
        If IsNumeric(hs) And IsNumeric(Min) Then
            If Int(hs) <= 24 And Int(Min) <= 59 And Int(hs & Min) <= 2400 Then
                Hora = hs & Min
                ValidarHora = True
            End If
        End If
        
End Function

Public Function ValidarHoraLarga(ByRef Hora As String) As Boolean
    Dim hs As String
    Dim Min As String
    Dim seg As String
    
        ValidarHoraLarga = False
        hs = Left(Hora, 2)
        Min = Mid(Hora, 4, 2)
        seg = Mid(Hora, 7, 2)
        
        If IsNumeric(hs) And IsNumeric(Min) And IsNumeric(seg) Then
            If Int(hs) <= 24 And Int(Min) <= 59 And Int(seg) <= 59 And Int(hs & Min) <= 2400 Then
                'Hora = hs & min & seg
                ValidarHoraLarga = True
            End If
        End If
        
End Function

Public Function ValidarHoraAMPM(ByRef Hora As String) As Boolean
Dim hs As String
Dim Min As String
    
    ValidarHoraAMPM = False
    hs = Left(Hora, 2)
    If Mid(Hora, 3, 1) = ":" Then
        Min = Mid(Hora, 4, 2)
    Else
        Min = Mid(Hora, 3, 2)
    End If
    If IsNumeric(hs) And IsNumeric(Min) Then
        If Int(hs) <= 24 And Int(Min) <= 59 And Int(hs & Min) <= 2400 Then
            Hora = hs & Min
            ValidarHoraAMPM = True
        End If
    End If
End Function

Public Sub SumoHoras_old(fecini As Date, HoraIni As String, horamas As String, ByRef Mifecha, ByRef MiHora)
'Mas conocido como gtis1h01
'Suma de Horas
'La hora debe estar en formato 9999
Const Hora = 60
Const Dia = 1440
Dim totminutos As Integer
Dim Minutos As Integer
Dim minutosmas As Integer
Dim canthoras As Integer

Dim fecha_aux As Date

    fecha_aux = fecini
    Minutos = (Val(Mid(HoraIni, 1, 2)) * Hora + _
                Val(Mid(HoraIni, 3, 2)))
    minutosmas = (Val(Mid(horamas, 1, 2)) * Hora + _
                Val(Mid(horamas, 3, 2)))
    totminutos = (Minutos + minutosmas)
    Do While (totminutos >= Dia)
    '      fecini = fecini + 1
        fecha_aux = DateAdd("d", 1, fecha_aux)
        totminutos = totminutos - Dia
    Loop
    'Hay que pasar los minutos a horas y minutos
    Do While (totminutos >= Hora)
    '      fecini = fecini + 1
        canthoras = canthoras + 1
        totminutos = totminutos - Hora
    Loop
    Mifecha = fecha_aux
    MiHora = Format(canthoras, "00") & ":" & Format(totminutos, "00")
    
End Sub

Public Sub SumoHoras(ByVal fecini As Date, ByVal HoraIni As String, ByVal horamas As String, ByRef Mifecha, ByRef MiHora)
'Mas conocido como gtis1h01
'Suma de Horas
'La hora debe estar en formato 9999
Const Hora = 60
Const Dia = 1440
Dim totminutos As Integer
Dim Minutos As Integer
Dim minutosmas As Integer
Dim canthoras As Integer

Dim fecha_aux As Date

    fecha_aux = fecini
    Minutos = (Val(Mid(HoraIni, 1, 2)) * Hora + _
                Val(Mid(HoraIni, 3, 2)))
    minutosmas = (Val(Mid(horamas, 1, 2)) * Hora + _
                Val(Mid(horamas, 3, 2)))
    totminutos = (Minutos + minutosmas)
    Do While (totminutos >= Dia)
    '      fecini = fecini + 1
        fecha_aux = DateAdd("d", 1, fecha_aux)
        totminutos = totminutos - Dia
    Loop
    'Hay que pasar los minutos a horas y minutos
    Do While (totminutos >= Hora)
    '      fecini = fecini + 1
        canthoras = canthoras + 1
        totminutos = totminutos - Hora
    Loop
    Mifecha = fecha_aux
    MiHora = Format(canthoras, "00") & ":" & Format(totminutos, "00")
    
End Sub


Public Sub Redondeo_Horas_Tipo(HoraIni As String, Tip_Red As Integer, ByRef Hora1 As Single)
  
  Dim objRs As New ADODB.Recordset
  Dim Minutos As Single
  Const Dia = 1440
  Const Hora = 60

  Minutos = Int(Mid(HoraIni, 4, 2))

  StrSql = "SELECT * FROM tipredondeo INNER JOIN tipreddet ON tipredondeo.trdnro = tipreddet.trdnro WHERE (tipredondeo.trdnro = " & Tip_Red & ") AND " & _
           "(tipreddet.trdetdesde <= " & Minutos & ") AND " & _
           "(tipreddet.trdethasta >= " & Minutos & ")"
  OpenRecordset StrSql, objRs
  If Not objRs.EOF Then
      Minutos = (Int(Mid(HoraIni, 1, 3)) * Hora) + objRs!trdetvalor
      Hora1 = Minutos / Hora
  Else
      Minutos = (Int(Mid(HoraIni, 1, 3)) * Hora) + Int(Mid(HoraIni, 4, 2))
      Hora1 = Minutos / Hora
  End If
  
  If objRs.State = adStateOpen Then objRs.Close
  Set objRs = Nothing
  
End Sub


Public Sub Redondeo_Horas_Tipo2(ByVal HoraIni As String, ByVal Tip_Red As Integer, ByVal Duracion As Single, ByRef Hora1 As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que redondea la cantidad de hs segun el redondeo de minutos configurado
' Autor      : FGZ
' Fecha      :
' Ultima Mod.: 03/11/2008
' Descripcion: Ahora tiene en cuenta la durecion de la hora configurada en la FP
' ---------------------------------------------------------------------------------------------
'Const Dia = 1440

Dim objRs As New ADODB.Recordset
Dim Minutos As Single

If Duracion = 0 Then
    Duracion = 60
End If
   
Minutos = Int(Mid(HoraIni, 4, 2))
StrSql = "SELECT * FROM tipredondeo INNER JOIN tipreddet ON tipredondeo.trdnro = tipreddet.trdnro WHERE (tipredondeo.trdnro = " & Tip_Red & ") AND "
StrSql = StrSql & "(tipreddet.trdetdesde <= " & Minutos & ") AND "
StrSql = StrSql & "(tipreddet.trdethasta >= " & Minutos & ")"
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    Minutos = (Int(Mid(HoraIni, 1, 3)) * Duracion) + objRs!trdetvalor
    Hora1 = Minutos / Duracion
'Else
'    Minutos = (Int(Mid(HoraIni, 1, 3)) * Duracion) + Int(Mid(HoraIni, 4, 2))
'    Hora1 = Minutos / Duracion
End If
  
If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing
End Sub



Public Sub Redondeo_Horas_Tipo2_old(ByVal HoraIni As String, ByVal Tip_Red As Integer, ByVal Duracion As Single, ByRef Hora1 As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Procedimiento que redondea la cantidad de hs segun el redondeo de minutos configurado
' Autor      : FGZ
' Fecha      :
' Ultima Mod.: 03/11/2008
' Descripcion: Ahora tiene en cuenta la durecion de la hora configurada en la FP
' ---------------------------------------------------------------------------------------------
'Const Dia = 1440

Dim objRs As New ADODB.Recordset
Dim Minutos As Single

   
Minutos = Int(Mid(HoraIni, 4, 2))
StrSql = "SELECT * FROM tipredondeo INNER JOIN tipreddet ON tipredondeo.trdnro = tipreddet.trdnro WHERE (tipredondeo.trdnro = " & Tip_Red & ") AND "
StrSql = StrSql & "(tipreddet.trdetdesde <= " & Minutos & ") AND "
StrSql = StrSql & "(tipreddet.trdethasta >= " & Minutos & ")"
OpenRecordset StrSql, objRs
If Not objRs.EOF Then
    Minutos = (Int(Mid(HoraIni, 1, 3)) * Duracion) + objRs!trdetvalor
    Hora1 = Minutos / Duracion
'Else
'    Minutos = (Int(Mid(HoraIni, 1, 3)) * Duracion) + Int(Mid(HoraIni, 4, 2))
'    Hora1 = Minutos / Duracion
End If
  
If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing
End Sub

Public Function FraccionaHs(hora_inicio As String, tipo_frac As Integer) As String
    Dim objRs As New ADODB.Recordset
    
    Dim Minutos As Integer
    Dim Dia As Integer
    Dim Hora As Integer
    Dim result As String
    Dim dummy As Boolean
    
    Dia = 1440
    Hora = 60
    If Not EsNulo(hora_inicio) Then
        Minutos = Int(Mid(hora_inicio, 3, 2))
    Else
        Minutos = 0
    End If
    'minutos = IIf(Not EsNulo(hora_inicio), Int(Mid(hora_inicio, 3, 2)), 0)
    
    StrSql = "SELECT * FROM tipredondeo INNER JOIN tipreddet ON tipreddet.trdnro = tipredondeo.trdnro " & _
             " WHERE  (tipredondeo.trdnro = " & tipo_frac & " AND tipreddet.trdetdesde <= " & Minutos & " ) AND " & _
             " (tipreddet.trdethasta >= " & Minutos & ")"
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
    If result = "00:00" And Minutos = 1440 Then result = "24:00"
    dummy = ValidarHora(result)
    FraccionaHs = result
       
End Function

Public Sub RestaHs(Fecha_Inicio As Date, hora_inicio As String, Fecha_Fin As Date, hora_fin As String, ByRef TotDias, ByRef tothoras As Integer, ByRef TotMin As Integer)
Dim total As Integer
Dim cantdias  As Integer
Dim canthoras As Integer
Dim Dia   As Integer '  cantidad de minutos en un dia
Dim Hora As Integer   ' cantidad de minutos en una hora

    Dia = 1440
    Hora = 60
    canthoras = 0
    If Not EsNulo(hora_fin) And Not EsNulo(hora_inicio) Then
        canthoras = (Int(Mid(hora_fin, 1, 2)) * Hora + _
                       Int(Mid(hora_fin, 3, 2))) - _
                      (Int(Mid(hora_inicio, 1, 2)) * Hora + _
                       Int(Mid(hora_inicio, 3, 2)))
    End If
    cantdias = DateDiff("d", Fecha_Inicio, Fecha_Fin) * Dia
    
    total = cantdias + canthoras
    TotDias = Int(total / Dia)
    tothoras = Int((total Mod Dia) / Hora)
    TotMin = (total Mod Hora)
End Sub



Public Function Menor_Igual_Hora(ByVal F1 As Date, ByVal H1 As String, ByVal F2 As Date, ByVal H2 As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna TRUE si la hora 1 es <= que la hora 2.
' Autor      : FGZ
' Fecha      : 26/10/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim R As Boolean

    If F1 < F2 Then
        R = True
    Else
        If F1 > F2 Then
            R = False
        Else
            If H1 <= H2 Then
                R = True
            Else
                R = False
            End If
        End If
    End If
    
    Menor_Igual_Hora = R
End Function


Public Function Mayor_Igual_Hora(ByVal F1 As Date, ByVal H1 As String, ByVal F2 As Date, ByVal H2 As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna TRUE si la hora 1 es >= que la hora 2.
' Autor      : FGZ
' Fecha      : 26/10/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim R As Boolean

    If F1 > F2 Then
        R = True
    Else
        If F1 < F2 Then
            R = False
        Else
            If H1 >= H2 Then
                R = True
            Else
                R = False
            End If
        End If
    End If
    
    Mayor_Igual_Hora = R
End Function


Public Function Mayor_Hora(ByVal F1 As Date, ByVal H1 As String, ByVal F2 As Date, ByVal H2 As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna TRUE si la hora 1 es > que la hora 2.
' Autor      : FGZ
' Fecha      : 26/10/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------

    Mayor_Hora = Not Menor_Igual_Hora(F1, H1, F2, H2)
End Function

Public Function Menor_Hora(ByVal F1 As Date, ByVal H1 As String, ByVal F2 As Date, ByVal H2 As String) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna TRUE si la hora 1 es > que la hora 2.
' Autor      : FGZ
' Fecha      : 26/10/2005
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------

    Menor_Hora = Not Mayor_Igual_Hora(F1, H1, F2, H2)
End Function





Public Function HorasInterseccion(ByVal R1FD As Date, ByVal R1HD As String, ByVal R1FH As Date, ByVal R1HH As String, ByVal R2FD As Date, ByVal R2HD As String, ByVal R2FH As Date, ByVal R2HH As String) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de horas que caen en la interseccion de 2 rangos de pares fecha-hora
'       R1FD -----------------R1FH
'       R2FD -----------------R2FH
' Parametros entrada
'              R1FD --> Fecha desde del rango 1
'              R1FH --> Fecha hasta del rango 1
'              R2FD --> Fecha desde del rango 2
'              R2FH --> Fecha hasta del rango 2

'              R1HD --> Hora desde del rango 1
'              R1HH --> Hora hasta del rango 1
'              R2HD --> Hora desde del rango 2
'              R2HH --> Hora hasta del rango 2
' Autor      : FGZ
' Fecha      : 07/11/2008
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim TotHor As Single
Dim Tdias As Integer
Dim Thoras As Integer
Dim Tmin As Integer

TotHor = 0

'Rango1     [---------------]
'Rango2           (----------------)
If objFechasHoras.Menor_Hora(R1FD, R1HD, R2FD, R2HD) And objFechasHoras.Menor_Hora(R2FD, R2HD, R1FH, R1HH) And objFechasHoras.Menor_Igual_Hora(R1FH, R1HH, R2FH, R2HH) Then
    objFechasHoras.RestaHs R2FD, R2HD, R1FH, R1HH, Tdias, Thoras, Tmin
    'HCdesde = R2HD
    'HChasta = R1HH
    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
End If

'Rango1         [---------------]
'Rango2     (----------------)
If objFechasHoras.Mayor_Igual_Hora(R1FD, R1HD, R2FD, R2HD) And objFechasHoras.Menor_Hora(R1FD, R1HD, R2FH, R2HH) And objFechasHoras.Mayor_Hora(R1FH, R1HH, R2FH, R2HH) Then
    objFechasHoras.RestaHs R1FD, R1HD, R2FH, R2HH, Tdias, Thoras, Tmin
    'HCdesde = R1HD
    'HChasta = R2HH
    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
End If

'rango1         [---------------]
'rango2     (-----------------------)
If objFechasHoras.Mayor_Igual_Hora(R1FD, R1HD, R2FD, R2HD) And objFechasHoras.Menor_Hora(R1FD, R1HD, R2FH, R2HH) And objFechasHoras.Menor_Igual_Hora(R1FH, R1HH, R2FH, R2HH) And objFechasHoras.Mayor_Hora(R1FH, R1HH, R2FD, R2HD) Then
    objFechasHoras.RestaHs R1FD, R1HD, R1FH, R1HH, Tdias, Thoras, Tmin
    'HCdesde = R1HD
    'HChasta = R1HH
    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
End If

'rango1     [---------------]
'rango2         (--------)
If objFechasHoras.Menor_Hora(R1FD, R1HD, R2FD, R2HD) And objFechasHoras.Mayor_Hora(R1FH, R1HH, R2FH, R2HH) Then
    objFechasHoras.RestaHs R2FD, R2HD, R2FH, R2HH, Tdias, Thoras, Tmin
    'HCdesde = R2HD
    'HChasta = R2HH
    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
End If
   
HorasInterseccion = TotHor
    
End Function



