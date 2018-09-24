Attribute VB_Name = "Politicas"
Option Explicit


Public Sub Politica(Numero As Integer)
Dim objRs As New ADODB.Recordset 'Como esta función es recursiva el recordset lo tengo que definir en forma local
Dim StrSql As String
Dim det As Integer

    StrSql = "SELECT gti_cabpolitica.cabpolnro,gti_cabpolitica.cabpolnivel, gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro, gti_detpolitica.detpolprograma " & _
        "FROM gti_cabpolitica INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro " & _
        "WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 3 And gti_alcanpolitica.alcpolorigen = " & empleado.Ternro & " AND gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 "

    OpenRecordset StrSql, objRs
    
    If objRs.EOF Then
        ' FGZ - 04/08/2003
        ' Cambio estruc_actual por his_estructura
        StrSql = " SELECT gti_cabpolitica.cabpolnro, gti_cabpolitica.cabpolnivel,gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro,gti_detpolitica.detpolprograma,alcance_testr.alteOrden "
        StrSql = StrSql & " FROM gti_cabpolitica "
        StrSql = StrSql & " INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro "
        StrSql = StrSql & " INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro "
        StrSql = StrSql & " INNER JOIN his_estructura ON gti_alcanpolitica.alcpolorigen = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN alcance_tEstr ON his_estructura.tenro = alcance_tEstr.tenro "
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = his_estructura.ternro"
        StrSql = StrSql & " WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 2 And gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 And alcance_testr.tanro = 1 AND empleado.ternro = " & empleado.Ternro & " AND his_estructura.htethasta IS NULL ORDER BY alcance_testr.AlteOrden Asc "
        
        'StrSql = " SELECT gti_cabpolitica.cabpolnro, gti_cabpolitica.cabpolnivel, gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro, gti_detpolitica.detpolprograma,alcance_testr.alteOrden "
        'StrSql = StrSql & " FROM gti_cabpolitica "
        'StrSql = StrSql & " INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro "
        'StrSql = StrSql & " INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro "
        'StrSql = StrSql & " INNER JOIN estruc_actual ON gti_alcanpolitica.alcpolorigen = estruc_actual.estrnro "
        'StrSql = StrSql & " INNER JOIN alcance_tEstr ON estruc_actual.tenro = alcance_tEstr.tenro "
        'StrSql = StrSql & " WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 2 And gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 And alcance_testr.tanro = 1 ORDER BY alcance_testr.AlteOrden ASC"

        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            StrSql = " SELECT gti_cabpolitica.cabpolnro,gti_cabpolitica.cabpolnivel, gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro, gti_detpolitica.detpolprograma " & _
             " FROM gti_cabpolitica INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro " & _
             " WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 1 And gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 "

            OpenRecordset StrSql, objRs
        End If
    End If
    
    
    
    If Not objRs.EOF Then
    
'    If objrs.State = adStateOpen Then objrs.Close
'    objrs.Open "Politica(" & Numero & ")", objConn, adOpenStatic, adLockReadOnly, adCmdStoredProc
'    If Not IsNull(objrs.Fields(0).Value) Then
'         det = CInt(objrs.Fields(0).Value)
        
        det = objRs!detpolprograma
        
        If depurar Then GeneraTraza empleado.Ternro, p_fecha, "Nª " & Numero, Str(det)
        Select Case Numero
            Case 15
                Politica15 p_fecha, empleado.Ternro, hora_desde, fecha_desde, hora_hasta, fecha_hasta, nro_turno, Nro_dia, P_asignacion, esFeriado
            Case 20
                Politica20 p_fecha, p_fecha, depurar, empleado.Ternro, det
            Case 30
                Politica30 p_fecha, empleado.Ternro, nro_turno, depurar, det
            Case 50
                Politica50 p_fecha, empleado.Ternro, det
            Case 60
                Politica60 p_fecha, empleado.Ternro, Nro_dia, P_asignacion, det
            Case 70
                Politica70 p_fecha, empleado.Ternro, hora_desde, fecha_desde, hora_hasta, fecha_hasta, nro_turno, depurar, ok, det
            Case 80
                Politica80 empleado.Ternro, p_fecha, fecha_desde, hora_desde, fecha_hasta, hora_hasta, det
            Case 85
                Politica85 p_fecha, empleado.Ternro, nro_grupo, nro_turno, Tiene_Justif, no_trabaja_just, depurar, ok, det
            Case 90
                Politica90 p_fecha, empleado.Ternro, hora_desde, fecha_desde, hora_hasta, fecha_hasta, nro_turno, ok, det
            Case 95
                Forma_embudo = True
            Case 100
              '  Politica100 Nro_dia, p_fecha, empleado.ternro, hora_desde, fecha_desde, hora_hasta, fecha_hasta, p_asignacion, Cod_justificacion1, Cod_justificacion2
            Case 110
                Politica110 p_fecha, empleado.Ternro, nro_turno, depurar, det
            Case 130
                Politica130 det
            Case 140
                Politica140 det
            Case 150
                Politica150 det
            Case 160
                Call Politica160
            Case 170
                Call Politica170
            Case 180
                Call Politica180
            Case 190
                Politica190 empleado.Ternro, nro_turno, p_fecha, tol, acumula, depurar, toldto, acumula_dto, det
            Case 200
                Politica200 empleado.Ternro, nro_turno, p_fecha, toltemp, acumula_temp, depurar, toldto, acumula_dto, det
            Case 380
                Politica380 det
            Case 390
                Politica390
            Case 402
                'politica402
            Case 400
                Politica400 p_fecha, empleado.Ternro, hora_desde, fecha_desde, hora_hasta, fecha_hasta, nro_turno, depurar, det
            Case 410
                Politica410 p_fecha, empleado.Ternro, hora_desde, fecha_desde, hora_hasta, fecha_hasta, nro_turno, depurar, det
            Case 430
                'politica430
            Case 440
                Politica440 p_fecha, empleado.Ternro, Nro_dia, aux_TipoDia, aux_Tipohora, Total_horas, det
            Case 450
                Politica450 nro_turno, p_fecha, empleado.Ternro, nro_grupo, Nro_fpgo, Dia_libre, esFeriado, Existe_Reg, depurar, det
            Case 460
                'Politica460 empleado.ternro, p_fecha, empleado.Grupo, depurar
                Call Politica(464)
                Call Politica(465)
            Case 464
                Politica464 p_fecha, empleado.Ternro, nro_turno, Dias_trabajados, Dias_laborables, depurar, det
            Case 465
                Politica465 p_fecha, empleado.Ternro, nro_turno, Dias_trabajados, Dias_laborables, depurar, det
            Case 470
                Politica470 Sigo_Generando, det
            Case 480
                Politica480 empleado.Ternro, p_fecha, det
            Case 540
                'Politica540 tipo_hora, empleado.Ternro, p_fecha, nro_turno, Total_horas
            Case 550
                Politica550
            Case 560
                politica560
            Case 570
                politica570
            Case 575
                politica575
            Case 580
                politica580
            Case 590
                politica590
            Case 650
                Politica650
            Case 800
                politica800 fec_proc, det
            Case 810
                politica810 Usa_Conv
            Case 820
                Politica820 p_fecha, empleado.Ternro, nro_turno
            Case 850
                Politica850
            Case 900
                Politica900 NroProceso, det
            Case 2000
                Politica2000 NroProceso, det
        End Select
    End If
End Sub





Public Function ConvHora(ByVal Hora As String) As Date
Dim MiHora As String
' Hora viene como string sin :
    ConvHora = Mid(Hora, 1, 2) & ":" & Mid(Hora, 3, 2)
End Function

Public Function ConvHoraBD(ByVal Hora As Date) As String
'    ConvHoraBD = "#" & Format(hora, "hh:mm") & "#"
    ConvHoraBD = "'" & Format(Hora, "hhmm") & "'"
End Function


Sub PasarMin(ByVal Hora As String, ByVal Dias As Integer, ByRef minutos As Integer)
' Pasa una cantidad de dias y horas a minutos
' Devuelve sin :
Dim h As Integer
Dim m As String
h = (Val(Mid(Hora, 1, 2))) * 60
m = Val(Mid(Hora, 3, 2))
minutos = h + m + (1440 * Dias)
End Sub





'Public Function PasarHr(ByRef hora As String) As Boolean
'' Quita los : de la hora
'    Dim hs As String
'    Dim min As String
'
'        PasarHr = False
'        hs = Left(hora, 2)
'        If Mid(hora, 3, 1) = ":" Then min = Mid(hora, 4, 2) Else min = Mid(hora, 3, 2)
'        If IsNumeric(hs) And IsNumeric(min) Then
'            If Int(hs) <= 24 And Int(min) <= 59 And Int(hs & min) <= 2400 Then
'                hora = hs & min
'                PasarHr = True
'            End If
'        End If
'End Function




Public Sub Politica15(Fecha As Date, nroter As Long, hdesde As String, fdesde As Date, hhasta As String, fhasta As Date, nro_turno As Long, Nro_dia As Integer, P_asignacion As Boolean, esFeriado As Boolean)
' Posibles cambios de turno no informados
' fgz 25/4/2003

Dim RsRep As New ADODB.Recordset
Dim RsDia As New ADODB.Recordset
Dim RsReg As New ADODB.Recordset
Dim RsDet As New ADODB.Recordset

Dim Distancia As Long
Dim TipoHoraPCT As Long
Dim HTeorica As String ' Hora Teorica

Dim tdias As Integer
Dim thoras As Integer
Dim tmin As Integer
Dim Resultado As Single
Dim EsDiaLibre As Boolean


'Busco las horas Produccion
StrSql = "SELECT * FROM confrep WHERE repnro = 56"
OpenRecordset StrSql, RsRep
' columna 1 = distancia
' columna 2 = Tipo de Hora a Generar

If RsRep.EOF Then
    Exit Sub
End If

Do While Not RsRep.EOF
    If RsRep!confnrocol = 1 Then
        Distancia = RsRep!confval
    Else
        If RsRep!confnrocol = 2 Then
            TipoHoraPCT = RsRep!confval
        End If
    End If

    RsRep.MoveNext
Loop

EsDiaLibre = False


If P_asignacion Then

    StrSql = "SELECT * FROM gti_detturtemp WHERE (ternro =" & nroter & " ) and (" & _
             "gttempdesde <= " & ConvFecha(Fecha) & ") and (" & _
             ConvFecha(Fecha) & " <= gttemphasta)"
    OpenRecordset StrSql, RsDet
    If Not RsDet.EOF Then
       HTeorica = RsDet!ttemphdesde1
       EsDiaLibre = RsDet!ttemplibre
    End If

Else
    StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_dia
    OpenRecordset StrSql, RsDia
    
    HTeorica = RsDia!diahoradesde1
    EsDiaLibre = RsDia!Dialibre
    
End If

If EsDiaLibre Or esFeriado Then Exit Sub

StrSql = "SELECT * FROM gti_registracion WHERE ternro = " & nroter & _
    " AND ((regfecha > " & ConvFecha(fdesde) & ") OR (regfecha =" & ConvFecha(fdesde) & " AND reghora >= '" & hdesde & _
    "'))AND (( regfecha < " & ConvFecha(fhasta) & ") OR (regfecha = " & ConvFecha(fhasta) & " AND reghora <= '" & hhasta & "'))"
    
    
'    StrSql = "SELECT * FROM gti_registracion " & _
'    " WHERE ternro= " & P_TerNro & _
'    " AND (regestado = 'I')" & _
'    " AND ((regfecha > " & ConvFecha(fecha_desde) & _
'    ") OR (regfecha = " & ConvFecha(fecha_desde) & _
'    " AND reghora >= '" & hora_desde & "') )" & _
'    " AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
'    ") OR (regfecha = " & ConvFecha(fecha_hasta) & _
'    " AND reghora <= '" & hora_hasta & "') )"
    
    
    
OpenRecordset StrSql, RsReg

If Not RsReg.EOF Then
    RsReg.MoveFirst

    ' obtengo la diferencia de horas
    Call objFechasHoras.RestaHs(Fecha, HTeorica, RsReg!regfecha, RsReg!reghora, tdias, thoras, tmin)

    Resultado = (tdias * 24) + (thoras + (tmin / 60))
    
    If Resultado >= Distancia Then
                
        ' inserto el tipo de hora posible cambio de Turno
        StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
            "horvalido,ternro,thnro,turnro,empleg,horfecrep,horhoradesde,horhorahasta" & _
            ",horestado,normnro,normnro2) VALUES (" & Resultado & "," & _
            ConvFecha(fdesde) & "," & ConvFecha(fhasta) & "," & "0,-1," & _
            nroter & "," & TipoHoraPCT & "," & nro_turno & "," & empleado.Legajo & "," & _
            ConvFecha(Fecha) & ",'" & HTeorica & "','" & RsReg!reghora & "','',4,4)"
        
            objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
End If

If RsDia.State = adStateOpen Then RsDia.Close
If RsReg.State = adStateOpen Then RsReg.Close
If RsDet.State = adStateOpen Then RsDet.Close
If RsRep.State = adStateOpen Then RsRep.Close

Set RsDia = Nothing
Set RsRep = Nothing
Set RsDet = Nothing
Set RsReg = Nothing

End Sub


Public Sub Politica410(Fecha As Date, nroter As Long, ventdesde As String, ventFdesde As Date, venthasta As String, ventFhasta As Date, nro_turno As Long, deb As Boolean, subn As Integer)

'Archivo: 410auspa.p
'Descripci¢n: Genera ausencia parcial con las reg. del medio.

Dim regnro_ent As Long
Dim regnro_sal As Long
Dim rec As Integer
Dim thora As Integer
Dim hora1 As String
Dim hora2 As String
Dim fecha1 As Date
Dim fecha2 As Date
Dim canthoras As Single
Dim thoras As Integer
Dim tmin As Integer
Dim tdias As Integer
Dim tothor As Single

Dim generar As Boolean
Dim cant_reg    As Integer
Dim cant_entsal As Integer
Dim cant        As Integer
Dim i As Integer
Dim j As Integer


Dim rs As New ADODB.Recordset
Dim StrSql As String


hora1 = ""
i = 0: j = 1

StrSql = "select * from " & TTempWFTurno
OpenRecordset StrSql, objRs

cant_reg = objRs.RecordCount

StrSql = "select * from " & TTempWFDia
OpenRecordset StrSql, objRs

cant_entsal = objRs.RecordCount

If cant_reg < cant_entsal Then
    Exit Sub
End If

'StrSql = "select * from " & TTempWFTurno & " where (anornro <> 5 AND anornro <> 6) or (anornro is null)"
StrSql = "select * from " & TTempWFTurno
OpenRecordset StrSql, objRs

' Call CreateTempWFauspa
Call CreateTempTable(TTempWFAuspa)
Do While Not objRs.EOF

    StrSql = "INSERT INTO " & TTempWFAuspa & " (fecha, hora, par) VALUES (" & ConvFecha(objRs!Fecha) & ", '" & objRs!Hora & "', " & j & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    Debug.Print StrSql
    i = i + 1
    If i Mod 2 = 0 Then
        j = j + 1
    End If
    
    objRs.MoveNext
Loop

'Esto es lo nuevo
StrSql = "SELECT * from " & TTempWFAuspa
OpenRecordset StrSql, objRs, adLockOptimistic
objRs.MoveFirst
objRs.Delete
objRs.MoveLast
objRs.Delete

'Fin de lo nuevo
i = 0

StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 19" & _
" AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
rs.Open StrSql, objConn

If Not rs.EOF Then
    thora = rs!thnro
Else
    'Entrada en la traza
    GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Ausencia parcial para el Turno ", Str(nro_turno)
            
    '/* Deberia salir de todo el procedimiento */
    GoTo CierroTodo
End If

For i = 1 To j

    StrSql = "SELECT * FROM " & TTempWFAuspa & " WHERE par = " & i
    'rs.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
    OpenRecordset StrSql, objRs

    Do While Not objRs.EOF
        If hora1 = "" Then
            hora1 = objRs!Hora
            fecha1 = objRs!Fecha
        Else
            hora2 = objRs!Hora
            fecha2 = objRs!Fecha
            generar = True
        End If
    

        If generar Then

            ' RUN gtir2f01.p(fecha1, hora1, fecha2, hora2, output tdias, output thoras, output tmin).

    
            Call objFechasHoras.RestaHs(fecha1, hora1, fecha2, hora2, tdias, thoras, tmin)
            canthoras = thoras + (tmin / 60)

            If canthoras = 0 Then
                Exit Sub
            End If
        
            'Call ValidarTipoDeHora(thora, nro_turno, tipo_hora)
            ' Entrada en horario cumplido
            StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
            "horvalido,ternro,thnro,turnro,empleg,horfecrep,horhoradesde,horhorahasta" & _
            ",horestado,normnro,normnro2) VALUES (" & canthoras & "," & _
            ConvFecha(fecha1) & "," & ConvFecha(fecha2) & "," & "0,-1," & _
            nroter & "," & thora & "," & nro_turno & "," & empleado.Legajo & "," & _
            ConvFecha(Fecha) & ",'" & hora1 & "','" & hora2 & "','',7,7)"
        
            objConn.Execute StrSql, , adExecuteNoRecords

        End If
        objRs.MoveNext
    Loop
Next

CierroTodo:
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

StrSql = "DROP TABLE " & TTempWFAuspa
objConn.Execute StrSql, adExecuteNoRecords

End Sub

Sub Politica50(Fecha As Date, nroter As Long, subn As Integer)

'Genera una ventana de 18, 20, 22 o 24 hs. a partir de la primera registracion del dia.

Dim rs As New Recordset
Dim StrSql As String
Dim Hora As String
Dim fecres As Date
Dim TamVent As String
Dim horres As String
Dim ok As Boolean

' Busco la primer registracion no procesada
'Corrección: consulta corregida
StrSql = "SELECT * FROM gti_registracion where ((regestado is null) OR (regestado = 'I'))" & _
" and ternro= " & nroter & " and gti_registracion.regfecha = " & ConvFecha(Fecha) & _
" ORDER BY ternro ASC, regfecha ASC, reghora ASC"
'MsgBox (strSql)
OpenRecordset StrSql, rs
'Si no hay ni una registracion salgo.
If rs.EOF Then
    rs.Close
    Exit Sub
End If

Hora = rs!reghora
If Not objFechasHoras.ValidarHora(Hora) Then
    Exit Sub
End If
Select Case subn
Case 1
    TamVent = "1800"
Case 2
    TamVent = "2000"
Case 3
    TamVent = "2200"
Case 4
    TamVent = "2400"
Case 5
    TamVent = "1400"
Case 6
    TamVent = "1600"
Case 7
    TamVent = "1500"

End Select
fecres = rs!regfecha
horres = Hora

'Corrección: Las siguientes dos líneas son nuevas
fecha_desde = fecres
hora_desde = horres

' Sumo a la hora el tamano de la ventana
Call objFechasHoras.SumoHoras(rs!regfecha, Hora, TamVent, fecres, horres)
If Not objFechasHoras.ValidarHora(horres) Then
    rs.Close
    Exit Sub
Else
    'Agrego un extremo de 18, 20, 22 o 24 horas en el wf
    'StrSql = "INSERT INTO wf_gtitpa(tpanro,valfec,valchar) VALUES (" & "2," & ConvFecha(fecres) & ",'" & horres & "'" & ")"
    'objConn.Execute StrSql, , adExecuteNoRecords
    rs.Close
    'Corrección: Las siguientes dos líneas son nuevas
    fecha_hasta = fecres
    hora_hasta = horres
    
End If
' Cierro todo
Set rs = Nothing
End Sub

Sub Politica20(fecdesde As Date, fechasta As Date, deb As Boolean, nroter As Long, subn As Integer)
'Depura registraciones repetidas en un intervalo de 5, 10 minutos.

Dim rs As New Recordset
Dim StrSql As String
Dim DifFecha As Integer
Dim intervalo As Integer
Dim min_ant As Integer
Dim MinResta As Integer
Dim minutos As Integer
Dim fecha_ant As Date

min_ant = 0
minutos = 0
' Subnivel de la politica
Select Case subn
Case 1
    intervalo = 5
Case 2
    intervalo = 10
End Select

'Busco la primer registración dentro del intervalo especificado
StrSql = "SELECT * FROM GTI_REGISTRACION WHERE ternro= " & nroter & _
" and regfecha >= " & ConvFecha(fecdesde) & " and regfecha <= " & ConvFecha(fechasta) & _
" ORDER BY ternro ASC, regfecha ASC, reghora ASC"
rs.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
' Si no hay registraciones me voy
If rs.EOF Then
    rs.Close
    Set rs = Nothing
    Exit Sub
End If
rs.MoveFirst
'Paso a minutos la hora de la primera registracion (en min_ant)
Call PasarMin(rs!reghora, 0, min_ant)
fecha_ant = rs!regfecha

rs.MoveNext
'obtengo la siguiente registracion del empleado dentro del intervalo de interes.
Do While Not rs.EOF
    DifFecha = DateDiff("d", rs!regfecha, fecha_ant)
    'Paso a minutos la hora de la siguiente registracion
    Call PasarMin(rs!reghora, DifFecha, minutos)
    MinResta = minutos - min_ant
    'Si la diferencia entre registraciones es mínima
    If MinResta < intervalo Then
        'Entonces sobra una registracion, que borro.
        'rs.Delete
        StrSql = "DELETE FROM gti_registracion WHERE regnro = " & rs!Regnro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        If deb Then
            GeneraTraza rs!empleg, fecdesde, "Nº20 - Se Eliminó la Registraci¢n Nº", Str(rs!Regnro)
        End If
    Else
        min_ant = minutos
        fecha_ant = rs!regfecha
    End If
    'Paso a la siguiente registracion
    rs.MoveNext
Loop
' Cierro todo
rs.Close
Set rs = Nothing
End Sub

Sub Politica30(p_fecha As Date, nro_tercero As Long, nro_turno As Long, deb As Boolean, subn As Integer)

Dim rs_dia As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim i As Integer
Dim tothor As Single
Dim tdias As Integer
Dim tmin As Integer
Dim hora_ausencia As Integer
Dim canthoras As Single
Dim fec_ent As Date
Dim hor_ent As String
Dim es_turno_libre As Boolean

Dim HoraConvenio As Integer
Dim Texto As String

' Parámetros
Select Case subn
Case 1          ' 030sraus
'Genera Horario Ausente para el empleado sin registraciones.
    HoraConvenio = 2
    Texto = "Ausencia"

Case 2          ' 030srpre
' Genera presente
    HoraConvenio = 1
    Texto = "Obligatoria"

Case 3
    HoraConvenio = 2
    Call Politica30v1(p_fecha, nro_tercero, nro_turno, Nro_dia, HoraConvenio)
    Exit Sub

End Select
    

'Abro cursor wf_dia
StrSql = "SELECT * FROM " & TTempWFDia & " WHERE (hora is not null) and (entrada is not null) " & _
        " AND (hora <> '') ORDER BY codigo"
rs_dia.Open StrSql, objConn
If Not rs_dia.EOF Then rs_dia.MoveFirst

' Por cada registro no vacio en wf_dia
Do While Not rs_dia.EOF
    
    es_turno_libre = False
    
    If rs_dia!entrada Then
        fec_ent = rs_dia!Fecha
        hor_ent = rs_dia!Hora
        ' Next
        GoTo NextProgress
    
    Else
'       RUN gtir2f01.p(fec-ent, hor-ent, wf-dia.fecha, wf-dia.hora, output tdias, output thoras, output tmin).
        Call objFechasHoras.RestaHs(fec_ent, hor_ent, rs_dia!Fecha, rs_dia!Hora, tdias, thoras, tmin)
        tothor = thoras + (tmin / 60)
        If tothor = 0 Then
            ' Next
            GoTo NextProgress
        End If
        
        StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraConvenio & _
        " AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
        rs.Open StrSql, objConn
       
        If Not rs.EOF Then
            hora_ausencia = rs!thnro
        Else
            'Entrada en la traza
            GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora " & Texto & " para el Turno:", Str(nro_turno)
            
            '/* Deberia salir de todo el procedimiento */
            GoTo CierroTodo
        End If
        
        'Call ValidarTipoDeHora(hora_ausencia, nro_turno, tipo_hora)
        ' Entrada en horario cumplido
        StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
        "horvalido,ternro,thnro,turnro,empleg,horfecrep,horhoradesde,horhorahasta," & _
        "normnro,normnro2,horestado) VALUES (" & tothor & "," & _
        ConvFecha(fec_ent) & "," & ConvFecha(rs_dia!Fecha) & "," & CInt(False) & "," & CInt(True) & "," & _
        nro_tercero & "," & hora_ausencia & "," & nro_turno & "," & empleado.Legajo & "," & _
        ConvFecha(p_fecha) & ",'" & hor_ent & "','" & rs_dia!Hora & "',8,8,'')"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        rs.Close
    End If

NextProgress:
    If Not rs_dia.EOF Then
        rs_dia.MoveNext
    End If
Loop
If rs.State = adStateOpen Then rs.Close

'/* Genera Horario Ausente para Turnos Libres - La tabla wf-dia no tiene datos */

If es_turno_libre Then

'    FIND FIRST wf-dia NO-LOCK.
    rs_dia.MoveFirst

'    FIND FIRST per.gti_dias WHERE per.gti_dias.dianro = wf-dia.codigo NO-LOCK.
    StrSql = "SELECT * FROM gti_dias WHERE dianro = " & rs_dia!Codigo
    rs.Open StrSql, objConn
    canthoras = rs!diacanthoras
    
    ' Como no lo necesito mas lo cierro
    rs.Close
    
    StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraConvenio & _
    " AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
    rs.Open StrSql, objConn
       
    If Not rs.EOF Then
        hora_ausencia = rs!thnro
    Else
        'Entrada en la traza
        If deb Then GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora " & Texto & " para el Turno:", Str(nro_turno)
        
        ' Salgo del procedimiento cerrando los recordsets
        GoTo CierroTodo
    End If

    'Call ValidarTipoDeHora(hora_ausencia, nro_turno, tipo_hora)
    ' Entrada en horario cumplido
    StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
    "horvalido,ternro,thnro,turnro,empleg,horfecrep," & _
    "normnro,normnro2) VALUES (" & canthoras & "," & _
    ConvFecha(rs_dia!Fecha) & "," & ConvFecha(rs_dia!Fecha) & "," & "0,1," & _
    nro_tercero & "," & hora_ausencia & "," & nro_turno & "," & empleado.Legajo & "," & _
    ConvFecha(p_fecha) & ",8,8)"
        
    objConn.Execute StrSql, , adExecuteNoRecords

End If

CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_dia.State = adStateOpen Then rs_dia.Close
    Set rs_dia = Nothing

End Sub

Sub Politica30v1(p_fecha As Date, nro_tercero As Long, nro_turno As Long, Nro_dia As Integer, HoraConvenio As Integer)

Dim rs_dia As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim hdesde As String
Dim hhasta As String

Dim tothor As Single

Dim hora_ausencia As Integer
    
StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_dia
rs_dia.Open StrSql, objConn
If Not rs_dia.EOF Then
    tothor = rs_dia!diacanthoras
Else
    GoTo CierroTodo
End If

StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = " & HoraConvenio & _
" AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
rs.Open StrSql, objConn
       
If Not rs.EOF Then
    hora_ausencia = rs!thnro
Else
    GoTo CierroTodo
End If

StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
         "horvalido,ternro,thnro,turnro,empleg,horfecrep," & _
         "normnro,normnro2,horestado) VALUES (" & Round(tothor, 2) & "," & _
         ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & "," & CInt(False) & "," & CInt(True) & "," & _
         nro_tercero & "," & hora_ausencia & "," & nro_turno & "," & empleado.Legajo & "," & _
         ConvFecha(p_fecha) & ",8,8,'')"
       
         objConn.Execute StrSql, , adExecuteNoRecords

CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_dia.State = adStateOpen Then rs_dia.Close
    Set rs_dia = Nothing

End Sub


Sub Politica40(p_fecha As Date, fecharegistracion As Date, horaregistracion As String, nro_tercero As Long, nro_turno As Long, subn As Integer)

'Genera Horario teorico para el empleado con una registracion.

Dim rs As New ADODB.Recordset
Dim rs_dia As New ADODB.Recordset
Dim StrSql As String

Dim i As Integer
Dim tothor As Single
Dim dif_ent    As Integer
Dim dif_sal    As Integer
Dim tdias      As Integer
Dim thoras     As Integer
Dim tmin       As Integer
Dim fec_ent    As Date
Dim hor_ent    As String
Dim fec_sal    As Date
Dim hor_sal    As String
Dim hora_oblig As Integer
Dim hor_desde  As String
Dim hor_hasta  As String
Dim fec_desde  As Date
Dim fec_hasta  As Date
Dim canthoras As Single



'StrSql = "SELECT * FROM empleado WHERE ternro= " & nro_tercero
'rs.Open StrSql, objConn

'FIND FIRST per.gti_turno WHERE per.gti_turno.turnro = nro-turno
StrSql = "SELECT * FROM gti_turno WHERE turnro= " & nro_turno
rs.Open StrSql, objConn

'Abro cursor wf_dia
StrSql = "SELECT * FROM " & TTempWFDia & " ORDER BY codigo"
rs_dia.Open StrSql, objConn, adOpenKeyset

If (rs!TipoTurno = 1) Or (rs!TipoTurno = 2) Then
'* Normal o Rotativo */
    rs.Close
    rs_dia.MoveFirst
    hor_desde = rs_dia!Hora
    fec_desde = rs_dia!Fecha
    rs_dia.MoveLast
    hor_hasta = rs_dia!Hora
    fec_hasta = rs_dia!Fecha
'    /* Ver si la reg. est  m s cerca de la entrada o la salida */
    If (horaregistracion <= hor_desde) Then
'        /* Se sale ya que armar-embudo genera el presente del empleado */
        'Preguntar
        'Leave.
        GoTo CierroTodo
    Else
        If (horaregistracion >= hor_hasta) Then
        '/* Fichada posterior a la £ltima registraci¢n */
            '/* Se Genera una Entrada */
            hor_ent = hor_desde
            fec_ent = fec_desde
            hor_sal = hor_hasta
            fec_sal = fec_hasta
        Else
        '/* Fichada anterior a la £ltima registraci¢n */
        
            'RUN gtir2f01.p(fec-desde, hor-desde, fecharegistracion, horaregistracion, output tdias, output thoras, output tmin).
            Call objFechasHoras.RestaHs(fec_desde, hor_desde, fecharegistracion, horaregistracion, tdias, thoras, tmin)
            dif_ent = thoras + (tmin / 60)
            'RUN gtir2f01.p(fecharegistracion, horaregistracion, fec-hasta, hor-hasta, output tdias, output thoras, output tmin).
            Call objFechasHoras.RestaHs(fecharegistracion, horaregistracion, fec_hasta, hor_hasta, tdias, thoras, tmin)
            dif_sal = thoras + (tmin / 60)
            If dif_ent <= dif_sal Then
'           /* Más cercano a la entrada */
'                /* Se sale ya que armar-embudo genera el presente */
                'Preguntar
                'leave
                GoTo CierroTodo
            Else
'                /* Genero una Entrada */
                hor_ent = hor_desde
                fec_ent = fec_desde
                hor_sal = horaregistracion
                fec_sal = fecharegistracion
            End If
            
        End If
    End If
    
'    RUN gtir2f01.p(fec-ent, hor-ent, fec-sal, hor-sal, output tdias, output thoras, output tmin).
    Call objFechasHoras.RestaHs(fec_ent, hor_ent, fec_sal, hor_sal, tdias, thoras, tmin)
    tothor = thoras + (tmin / 60)
    If tothor = 0 Then
    'Preguntar:
'      NEXT.
        GoTo CierroTodo
    End If

'/* Horas Obligatorias */
    StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 1 " & _
    " AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
    rs.Open StrSql, objConn
           
    If Not rs.EOF Then
        hora_oblig = rs!thnro
    Else
        'Entrada en la traza
        GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Obligatoria para el Turno:", Str(nro_turno)
        
        ' Salgo del procedimiento cerrando los recordsets
        GoTo CierroTodo
    
    End If
        
    'Call ValidarTipoDeHora(hora_oblig, nro_turno, tipo_hora)
    ' Entrada en horario cumplido
    StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
    "horvalido,ternro,thnro,turnro,empleg,horhoradesde,horhorahasta" & _
    ") VALUES (" & tothor & "," & _
    ConvFecha(fec_ent) & "," & ConvFecha(fec_sal) & "," & CInt(False) & "," & CInt(True) & "," & _
    nro_tercero & "," & hora_oblig & "," & nro_turno & "," & empleado.Legajo & _
    ",'" & hor_ent & "','" & hor_sal & "')"
        
    objConn.Execute StrSql, , adExecuteNoRecords
    rs.Close
            
'/* Fin de turno Normal o Rotativo */
Else
'       /* Para turno Libre */
'    FIND FIRST wf-dia NO-LOCK.
    StrSql = "SELECT * FROM " & TTempWFDia & " ORDER BY codigo"
    rs_dia.Open StrSql, objConn

'    FIND FIRST per.gti_dias WHERE per.gti_dias.dianro = wf-dia.codigo NO-LOCK.
    StrSql = "SELECT * FROM gti_dias WHERE dianro = " & rs_dia!Codigo
    rs.Open StrSql, objConn
    canthoras = rs!diacanthoras
    
    ' Como no lo necesito mas lo cierro
    rs.Close
    
    StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 1 " & _
    " AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
    rs.Open StrSql, objConn
       
    If Not rs.EOF Then
        hora_oblig = rs!thnro
    Else
        'Entrada en la traza
        GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Obligatoria para el Turno:", Str(nro_turno)
        
        ' Salgo del procedimiento cerrando los recordsets
        GoTo CierroTodo
    End If

    'Call ValidarTipoDeHora(hora_oblig, nro_turno, tipo_hora)
    ' Entrada en horario cumplido
    StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
    "horvalido,ternro,thnro,turnro,empleg,horfecrep," & _
    "normnro,normnro2) VALUES (" & canthoras & "," & _
    ConvFecha(rs_dia!Fecha) & "," & ConvFecha(rs_dia!Fecha) & "," & CInt(False) & "," & CInt(True) & "," & _
    nro_tercero & "," & hora_oblig & "," & nro_turno & "," & empleado.Legajo & "," & _
    ConvFecha(p_fecha) & ",3,3)"
        
    objConn.Execute StrSql, , adExecuteNoRecords
        
End If

CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_dia.State = adStateOpen Then rs_dia.Close
    Set rs_dia = Nothing

End Sub

'Private Sub Escribo_gtitpa(tiempo As String)
'Dim StrSql As String
'
'StrSql = "INSERT INTO wf_gtitpa(tpanro, valchar) VALUES (" & _
'1 & ",'" & tiempo & "')"
'objConn.Execute StrSql, , adExecuteNoRecords
'
'End Sub

Public Sub Politica130(subn As Integer)
'Tolerancia general de llegada tarde.
Select Case subn
Case 1
    tol = "0005"
Case 2
    tol = "0010"
Case 3
    tol = "0015"
Case 4
    tol = "0000"
End Select
End Sub

Public Sub Politica140(subn As Integer)
'Tolerancia general de salida temprano.
Select Case subn
Case 1
    toltemp = "0005"
Case 2
    toltemp = "0010"
Case 3
    toltemp = "0015"
Case 4
    toltemp = "0000"
End Select
End Sub

Public Sub Politica150(subn As Integer)
' tolerancia general de descuento
Select Case subn
Case 1
    toldto = "0015"
Case 2
    toldto = "0030"
Case 3
    toldto = "0100"
Case 4
    toldto = "0000"
End Select
End Sub

Public Sub Politica160()
'Descripci¢n: Solamente devuelve true.
Dim StrSql As String

     acumula = True

End Sub

Public Sub Politica170()
'Descripci¢n: Devuelve true si acumula Hs de Ausencia.
Dim StrSql As String

    acumula_dto = True

End Sub

Public Sub Politica180()
'Descripci¢n: Devuelve true si acumula salidas temprano.
Dim StrSql As String

    acumula_temp = True

End Sub

Public Sub Politica380(subn As Integer)
'Descripci¢n: Devuelve true si un Feriado pesa mas que Día No laborable.
Dim StrSql As String
Select Case subn
Case 1
    diatipo = 1
Case 2
    diatipo = 2

End Select

End Sub

Public Sub Politica470(ByRef Sigo_Generando As Boolean, subn As Integer)
Select Case subn
Case 1
'Genero la Justificacion y en caso de que el empleado registre, no se genera nada.
    Sigo_Generando = False
Case 2
'Genero la Justificacion y en caso de que el empleado registre, se le genera el HC.
    Sigo_Generando = True
End Select
End Sub

Public Sub Politica440(Fecha As Date, nroter As Long, Nro_dia As Integer, nro_tdia As Integer, nro_thora As Integer, ByRef tothor As Single, subn As Integer)
Dim rs As New ADODB.Recordset

'Preguntar por {headlock.i}
Select Case subn

Case 1
'/* Saca la cantidad de horas correspondientes de las horas obligatorias del turno del empleado */
'Busco en la tabla de dias del turno de trabajo
    StrSql = "SELECT * FROM gti_dias where gti_dias.dianro = " & _
    Nro_dia & " ORDER BY dianro ASC"
    rs.Open StrSql, objConn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveFirst
        tothor = rs!diacanthoras
    End If

Case 2
'/* Saca la cantidad de horas correspondientes por el convenio del empleado, el tipo de d¡a y el tipo de hora */
'find first empleado where empleado.ternro = nroter
    
    ' Busco la cantidad de horas
    StrSql = "SELECT * FROM hsxconvenio where thnro = " & nro_thora & _
    " and tdnro = " & nro_tdia & " and convnro = " & convenio & _
    " ORDER BY convnro ASC, thnro ASC, tdnro ASC"
    rs.Open StrSql, objConn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveFirst
        tothor = rs!canths
    End If

Case 3
'/* Saca la cantidad de horas correspondientes desde lo configurado en el tipo de d¡a */
    StrSql = "SELECT tdnro,tdcanthoras FROM tipdia where tdnro = " & _
    nro_tdia & " ORDER BY tdnro ASC"
    rs.Open StrSql, objConn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveFirst
        tothor = rs!tdcanthoras
    End If
Case 5
'/* Saca la cantidad de horas correspondientes de las horas obligatorias 1er dia del turno del empleado */
    StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_subturno
    StrSql = StrSql & " ORDER BY diaorden"
    rs.Open StrSql, objConn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveFirst
        tothor = rs!diacanthoras
    End If
End Select

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub

Public Sub Politica70(Fecha As Date, nroter As Long, ventdesde As String, ventFdesde As Date, venthasta As String, ventFhasta As Date, nro_turno As Long, deb As Boolean, ByRef ok As Boolean, subn As Integer)
'Descripci¢n: Borro la primera registracion impar.
'subn = 1 : descarto primera registración
'subn = 2 : descarto la última.
'subn = 3 : genera error de tipo de hora.
Dim rs As New ADODB.Recordset
Dim cant As Integer
Dim contar As Boolean
Dim VentDConv As Date
Dim VentHConv As Date

ok = True
cant = 0
contar = True
VentDConv = ConvHora(ventdesde)
VentHConv = ConvHora(venthasta)

'/*** cuento las registraciones dentro de la ventana ***/
StrSql = "SELECT * FROM gti_registracion WHERE ternro= " & nroter & _
" and regfecha >= " & ConvFecha(ventFdesde) & " and  regfecha <= " & ConvFecha(ventFhasta) & _
" ORDER BY ternro ASC, regfecha ASC, reghora ASC"
rs.Open StrSql, objConn

Do While Not rs.EOF
  If rs!regfecha = ventFdesde Then
        ' Convierto para comparar
      If rs!reghora < ventdesde Then
        'Si es temprano no la cuento
        contar = False
      End If
  End If
  If rs!regfecha = ventFhasta Then
      If rs!reghora > venthasta Then
        'Si la registracion queda fuera de la ventana me voy
        Exit Do
      End If
  End If
  If contar Then
    cant = cant + 1
  Else
    contar = True
  End If
  rs.MoveNext
Loop
rs.Close
If (cant Mod 2) <> 0 Then ' /* las registraciones son impares */

    Select Case subn
    Case 2
'Busco la primer registración del empleado en el intervalo
        StrSql = "SELECT * FROM gti_registracion WHERE ternro= " & nroter & _
        " and regfecha >= " & ConvFecha(ventFdesde) & " AND regfecha <= " & ConvFecha(ventFhasta) & _
        " AND reghora >= " & ConvHoraBD(VentDConv) & _
        " ORDER BY ternro ASC, regfecha ASC, reghora ASC"
        rs.Open StrSql, objConn, adOpenDynamic, adLockOptimistic
        'Si encuentro una ...
        If Not rs.EOF Then
            rs.MoveFirst
            If deb Then
                GeneraTraza empleado.Ternro, Fecha, "Nro70 - Se Eliminó la (Primer) Registración Nro ", Str(rs!Regnro)
            End If
            ' La elimino
            rs.Delete
            rs.UpdateBatch
            rs.Close
        End If
    Case 1
    'Busco la ultima registración del empleado en el intervalo
        StrSql = "SELECT * FROM gti_registracion WHERE ternro= " & nroter & _
        " and regfecha >= " & ConvFecha(ventFdesde) & " and  regfecha <= " & ConvFecha(ventFhasta) & _
        " AND reghora <= " & ConvHoraBD(VentHConv) & _
        " ORDER BY ternro DESC, regfecha DESC, reghora DESC"
        rs.Open StrSql, objConn, adOpenDynamic, adLockOptimistic
        'Si encuentro una ...
        If Not rs.EOF Then
            rs.MoveFirst
            If deb Then
                GeneraTraza empleado.Ternro, Fecha, "Nro70 - Se Eliminó la (última) Registración Nro ", Str(rs!Regnro)
            End If
            ' La elimino
            rs.Delete
            rs.UpdateBatch
            rs.Close
        End If
    Case 3
        StrSql = "SELECT * FROM gti_config_tur_hor WHERE turnro= " & nro_turno & _
        " and conhornro = 22"
        rs.Open StrSql, objConn
        If rs.EOF Then
          MsgBox "No Esta Configurado el Tipo de Hora de Registracione Impares para el Turno: ", vbCritical
          rs.Close
          Exit Sub
        End If
        
        'Call ValidarTipoDeHora(rs!thnro, nro_turno, tipo_hora)
        
        StrSql = "INSERT INTO gti_horcumplido (Empleg,hordesde,horfecrep,horhasta,hormanual," & _
        "horvalido,ternro,thnro,turnro,horcant,horestado) VALUES (" & empleado.Legajo & "," & ConvFecha(Fecha) & "," & _
        ConvFecha(Fecha) & "," & ConvFecha(Fecha) & "," & "0,-1," & nroter & "," & _
        rs!thnro & "," & nro_turno & ",1,'')"
        objConn.Execute StrSql, , adExecuteNoRecords
        ok = False
        rs.Close

    End Select

End If

'Saco al rs de memoria
Set rs = Nothing
End Sub


Sub Politica110(p_fecha As Date, nro_tercero As Long, nro_turno As Long, deb As Boolean, subn As Integer)
'subn = 1 Genera Horario Ausente para el empleado que le faltan Reg. Oblig.
'subn = 2 Genera Ausente
'subn = 3 Genera error por falta de registración
'subn = 4 Genera Horario Te¢rico para el empleado que le faltan Reg. Oblig.

Dim rs_dia As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim i As Integer
Dim tothor As Single
Dim tdias As Integer
Dim tmin As Integer
Dim hora_oblig As Integer
Dim hora_generada As Integer
Dim fec_ent As Date
Dim hor_ent As String
Dim msgTraza As String
Dim es_turno_libre As Boolean


'Abro cursor
StrSql = "SELECT * FROM " & TTempWFDia & " ORDER BY codigo"
rs_dia.Open StrSql, objConn
rs_dia.MoveFirst

Do While Not rs_dia.EOF
    
    es_turno_libre = False
    
    If rs_dia!entrada Then
        fec_ent = rs_dia!Fecha
        hor_ent = rs_dia!Hora
        ' Next
        GoTo NextProgress
    
    Else

'       RUN gtir2f01.p(fec-ent, hor-ent, wf-dia.fecha, wf-dia.hora, output tdias, output thoras, output tmin).
        Call objFechasHoras.RestaHs(fec_ent, hor_ent, rs_dia!Fecha, rs_dia!Hora, tdias, thoras, tmin)
        tothor = thoras + (tmin / 60)
        If tothor = 0 Then
            ' Next
            GoTo NextProgress
        End If
               
        'Parametros
        Select Case subn
        Case 1, 3
            '/*Horas falta registracion */
            StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 10 "
            msgTraza = "No esta configurado el Tipo de Hora Falta Registracion para el Turno:"
        Case 2
'           /*Horas de Ausencia */
            StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 2 "
            msgTraza = "No esta configurado el Tipo de Hora Ausencia para el Turno:"
        Case 4
            '/*Horas Obligatorias */
            StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 1 "
            msgTraza = "No esta configurado el Tipo de Hora Obligatoria para el Turno:"
        Case 5
            Call Politica110v2(p_fecha, nro_tercero, nro_turno, deb, subn)
            Exit Sub
        End Select
        StrSql = StrSql & "AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
        OpenRecordset StrSql, rs
       
        If Not rs.EOF Then
            hora_generada = rs!thnro
        Else
            'Entrada en la traza
            If depurar Then
                GeneraTraza empleado.Ternro, p_fecha, msgTraza, Str(nro_turno)
            End If
            '/* Deberia salir de todo el procedimiento */
            GoTo CierroTodo
        End If
        
        'Call ValidarTipoDeHora(hora_generada, nro_turno, tipo_hora)
'       Creo un registro en la tabla de HC
        StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
        "horvalido,ternro,thnro,turnro,empleg,horfecrep," & _
        "normnro,normnro2,horhoradesde,horhorahasta,horestado) VALUES (" & tothor & "," & _
        ConvFecha(fec_ent) & "," & ConvFecha(rs_dia!Fecha) & "," & "0,-1," & _
        nro_tercero & "," & hora_generada & "," & nro_turno & "," & empleado.Legajo & "," & _
        ConvFecha(p_fecha) & ",3,3,'" & hor_ent & "','" & rs_dia!Hora & "','')"
        
        objConn.Execute StrSql, , adExecuteNoRecords

    End If
    rs_dia.MoveNext

NextProgress:
    If Not rs_dia.EOF Then
        rs_dia.MoveNext
    End If

Loop


'Cierro y libero todos los recordsets
CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_dia.State = adStateOpen Then rs_dia.Close
    Set rs_dia = Nothing


End Sub


Sub Politica85(p_fecha As Date, nroter As Long, nro_grupo As Long, nro_turno As Long, Tiene_Justif As Boolean, no_trabaja_just As Boolean, deb As Boolean, ok As Boolean, subn As Integer)

' Recorrida de registraciones a evaluar

Dim i, cantr, cant_reg, reg_oblig As Integer
Dim hora_generada As Integer
Dim entrada As Boolean
Dim msgTraza As String
Dim registro As Long

Dim rs_gti_detpolitica As New ADODB.Recordset
Dim rs_gti_alcanpolitica As New ADODB.Recordset
Dim rs_gti_cabpolitica As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim StrSql As String
'
i = 1
cant_reg = 0
cantr = 1
entrada = True
ok = True

'Obtengo el legajo
'StrSql = "SELECT * FROM empleado WHERE ternro= " & NroTer
'rs.Open StrSql, objConn
'Legajo = rs!Empleg

'Abro temporal wf_turno
'Preguntar por el by
'  FOR EACH wf-turno WHERE wf-turno.evenro = 2
'    By wf - turno.fecha
'      By wf - turno.Hora:
StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE evenro = 2" & _
" ORDER BY fecha ASC, hora ASC"
rs.Open StrSql, objConn, adOpenKeyset, adLockOptimistic, adCmdText
If Not rs.EOF Then rs.MoveFirst

' Por cada registro seleccionado de wf-turno
Do While Not rs.EOF
'    rs!par = i
'    rs!entrada = entrada
'    rs.UpdateBatch
    registro = rs!Regnro
    StrSql = "update " & TTempWFTurno & " set par = " & i & ", entrada = " & CInt(entrada) & " where regnro = " & registro
    objConn.Execute StrSql, , adExecuteNoRecords
    If (cantr Mod 2) = 0 Then
        i = i + 1
        entrada = True
    Else
        entrada = False
    End If
    cantr = cantr + 1
    cant_reg = cant_reg + 1
    rs.MoveNext
Loop


rs.Close

StrSql = "SELECT COUNT(codigo) as reg_oblig FROM " & TTempWFDia
OpenRecordset StrSql, rs
reg_oblig = rs!reg_oblig

'rs_wf_dia.Open "wf_dia", objConn, adOpenKeyset, adLockOptimistic
'reg_oblig = rs_wf_dia.RecordCount

If (Tiene_Justif) And (Not no_trabaja_just) Then
     reg_oblig = reg_oblig - 2
End If

'/* Hay menos reg. que en la definici¢n del turno */
If (cant_reg < reg_oblig) Then
    Call Politica(110)
    If cant_reg <= 1 Then
        ok = False
    End If
End If
        
End Sub

Public Sub Politica60(Fecha As Date, nroter As Long, Nro_dia As Integer, P_asignacion As Boolean, subn As Integer)

'Descripci¢n: Busca la primera entrada y £ltima salida y genera la ventana.
'veys 7 Ventana de 3 Horas para Entrada y 14 Para Salida.

'Nota: esto se necesita porque hace embudo con la ultima salida, debe tomar l corte final
Dim rs As New Recordset
Dim StrSql As String

Dim Hora As String
Dim ventent As String
Dim ventsal As String
Dim fecres As Date
Dim TamVent As String
Dim horres As String
Dim ok As Boolean


'Parametrización
Select Case subn

Case 6
    
    ventent = "0400"
    ventsal = "2000"

Case 3  'veys 4
    ventent = "0400"
    ventsal = "2400"

Case 7  'veys 2
    ventent = "0400"
    ventsal = "2000"
'Global
Case 8
    ventent = "0500"
    ventsal = "2200"
'individual
Case 9
    ventent = "0600"
    ventsal = "2400"
Case 10
    ventent = "0200"
    ventsal = "2400"


Case Else
    
    Call Politica60v2(Fecha, nroter, Nro_dia, P_asignacion, subn)
    Exit Sub

End Select

'Busco nro dia

If Not P_asignacion Then
    StrSql = "SELECT * FROM gti_dias WHERE dianro= " & Nro_dia
    OpenRecordset StrSql, rs
Else
    StrSql = "SELECT * FROM gti_detturtemp WHERE ternro= " & nroter & _
    " AND gttempdesde <= " & ConvFecha(Fecha) & " AND gttemphasta >= " & ConvFecha(Fecha)
    OpenRecordset StrSql, rs
End If
'rs.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
If rs.EOF Then
    Exit Sub
End If

rs.MoveFirst

If Not P_asignacion Then
    Hora = rs!diahoradesde1
Else
    Hora = rs!ttemphdesde1
End If

If Not objFechasHoras.ValidarHora(Hora) Then
    Exit Sub
End If
'run gtis1h01.p(fecha, hora, ventsal,output fecres, output horres)
Call objFechasHoras.RestaXHoras(Fecha, Hora, ventent, fecres, horres)
If Not objFechasHoras.ValidarHora(horres) Then
    Exit Sub
End If

fecha_desde = fecres
hora_desde = horres

If Not objFechasHoras.ValidarHora(Hora) Then
    Exit Sub
End If

Call objFechasHoras.SumoHoras(fecha_desde, hora_desde, ventsal, fecres, horres)
If Not objFechasHoras.ValidarHora(horres) Then
    Exit Sub
End If

fecha_hasta = fecres
hora_hasta = horres

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub

Public Sub Politica60v2(Fecha As Date, Ternro As Long, Nro_dia As Integer, P_asignacion As Boolean, subn As Integer)

Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim Hora As String
Dim ventent As String
Dim ventsal As String
Dim fecres As Date
Dim horres As String
Dim ok As Boolean


'Parametros

Select Case subn

Case 1  'veys 1
    ventent = "0300"
    ventsal = "0500"

Case 2  'veys 3
    ventent = "0400"
    ventsal = "0600"

Case 4  'veys 5
    ventent = "0300"
    ventsal = "0700"

Case 5  'veys 6
    ventent = "0300"
    ventsal = "0900"

End Select

If P_asignacion Then

    StrSql = "SELECT * FROM gti_detturtemp WHERE ternro= " & Ternro & _
    " AND gttempdesde <= " & ConvFecha(Fecha) & " AND gttemphasta >= " & ConvFecha(Fecha)
    OpenRecordset StrSql, rs
    'rs.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
    
    If rs.EOF Then
        Exit Sub
    End If
    
    rs.MoveFirst
'        /*** realiza la resta ***/
    Hora = rs!ttemphdesde1
    If Not objFechasHoras.ValidarHora(Hora) Then
        Exit Sub
    End If
    Call objFechasHoras.RestaXHoras(Fecha, Hora, ventent, fecres, horres)
    If Not objFechasHoras.ValidarHora(horres) Then
        Exit Sub
    End If
    
    fecha_desde = fecres
    hora_desde = horres
    
    If (rs!ttemphdesde3 <> "0000") And (rs!ttemphhasta3 <> "0000") Then
        If rs!diasiguiente = 1 Or rs!diasiguiente = 2 Or rs!diasiguiente = 3 Then
            Fecha = Fecha + 1
        End If
        Hora = rs!ttemphhasta3
    Else
        If (rs!ttemphdesde2 <> "0000") And (rs!ttemphhasta2 <> "0000") Then
            If rs!diasiguiente = 1 Or rs!diasiguiente = 2 Then
                Fecha = Fecha + 1
            End If
            Hora = rs!ttemphhasta2
        Else
            If rs!diasiguiente = 1 Then
                Fecha = Fecha + 1
            End If
            Hora = rs!ttemphhasta1
        End If
    End If

    If Not objFechasHoras.ValidarHora(Hora) Then
        Exit Sub
    End If
    
    Call objFechasHoras.SumoHoras(Fecha, Hora, ventsal, fecres, horres)
    If Not objFechasHoras.ValidarHora(horres) Then
        Exit Sub
    End If

    fecha_hasta = fecres
    hora_hasta = horres
Else
    StrSql = "SELECT * FROM gti_dias WHERE dianro= " & Nro_dia
    rs.Open StrSql, objConn, adOpenForwardOnly, adLockReadOnly
    
    If rs.EOF Then
        Exit Sub
    End If
    
    rs.MoveFirst

    Hora = rs!diahoradesde1
    
    If Not objFechasHoras.ValidarHora(Hora) Then
        Exit Sub
    End If
    
    Call objFechasHoras.RestaXHoras(Fecha, Hora, ventent, fecres, horres)
    
    If Not objFechasHoras.ValidarHora(horres) Then
        Exit Sub
    End If
    
    fecha_desde = fecres
    hora_desde = horres
    
    If rs!diahoradesde3 <> "0000" And rs!diahorahasta3 <> "0000" Then
        If rs!diasiguiente = 1 Or rs!diasiguiente = 2 Or rs!diasiguiente = 3 Then
            Fecha = Fecha + 1
        End If
        Hora = rs!diahorahasta3
    Else
        If rs!diahoradesde2 <> "0000" And rs!diahorahasta2 <> "0000" Then
            If rs!diasiguiente = 1 Or rs!diasiguiente = 2 Then
                Fecha = Fecha + 1
            End If
            Hora = rs!diahorahasta2
        Else
            If rs!diasiguiente = 1 Then
                Fecha = Fecha + 1
            End If
            Hora = rs!diahorahasta1
        End If
    End If

    If Not objFechasHoras.ValidarHora(Hora) Then
        Exit Sub
    End If
    
    Call objFechasHoras.SumoHoras(Fecha, Hora, ventsal, fecres, horres)
    If Not objFechasHoras.ValidarHora(horres) Then
        Exit Sub
    End If
    
    fecha_hasta = fecres
    hora_hasta = horres
    
End If

If rs.State = adStateOpen Then rs.Close

Set rs = Nothing

End Sub

Public Sub Politica650()

    usaControlDias = True

End Sub


Public Sub Politica450(nro_turno As Long, p_fecha As Date, nroter As Long, nro_grupo As Long, Nro_fpgo As Integer, Dia_libre As Boolean, esFeriado As Boolean, Existe_Reg As Boolean, deb As Boolean, subn As Integer)
' Imputa Horas sin control de presencia

'/*------------------------------------------------------------------------------
'    Busca el d¡a que le corresponde al empleado, de acuerdo al d¡a que se esta
'    procesando y al turno al que pertenece.
'------------------------------------------------------------------------------*/

Dim weekdia As Integer
Dim diatipo As Integer
Dim j As Integer
Dim dia As Integer

Dim rs_wf_gtitpa As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

diatipo = 0

If esFeriado And Dia_libre Then

    Call Politica(380)
    
    If Dia_libre Then
        diatipo = 1
    Else
        diatipo = 2
    End If
    
End If

weekdia = Weekday(p_fecha)


StrSql = "SELECT * FROM gti_desgdia  WHERE fpgonro= " & Nro_fpgo & _
" AND desgdtipo = " & diatipo & " AND desgdnrodia = " & weekdia & _
" ORDER BY desgdnrodia ASC, desgdtipo ASC"
rs.Open StrSql, objConn

If Not rs.EOF Then
    rs.MoveFirst
    dia = rs!desgdnro
    ' Cierro gti_desgdia
    rs.Close
'    FOR EACH per.gti_desghora OF per.gti_desgdia where (PER.gti_desghora.desghplus = 2) or
'                                                       (PER.gti_desghora.desghplus = 3) NO-LOCK:
    ' Para el dia en cuestion busco el desgloce de la hora
    StrSql = "SELECT * FROM gti_desghora  WHERE desgdnro = " & dia & _
    " AND ((desghplus = 2) OR (desghplus = 3))"
    rs.Open StrSql, objConn
    ' Para todas las horas del dia...
    Do While Not rs.EOF
        If (rs!desghplus = 3) And Existe_Reg Then
            rs.MoveNext
        Else
            'Call ValidarTipoDeHora(rs!thnro, nro_turno, tipo_hora)
        ' Entrada en horario cumplido
            StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
            "horvalido,ternro,thnro,turnro,empleg,horfecrep,horestado) VALUES (" & rs!desgcantmin & "," & _
            ConvFecha(p_fecha) & "," & _
            ConvFecha(p_fecha) & "," & "0,-1," & nroter & "," & rs!thnro & "," & nro_turno & _
            "," & empleado.Legajo & "," & ConvFecha(p_fecha) & ",'')"
            objConn.Execute StrSql, , adExecuteNoRecords
            rs.MoveNext
        End If
    Loop
    rs.Close
    
Else
    ' Genero una entrada con el error
    GeneraTraza empleado.Ternro, p_fecha, "No existe el desgloce del día"
End If

' Libero los recordsets
If rs_wf_gtitpa.State = adStateOpen Then rs_wf_gtitpa.Close
Set rs_wf_gtitpa = Nothing
Set rs = Nothing


End Sub


Public Sub Politica80(p_ternro As Long, Fecha As Date, fecha_desde As Date, hora_desde As String, fecha_hasta As Date, hora_hasta As String, subn As Integer)

' Registracion a evaluar

Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim p_horario As Boolean
Dim salir As Boolean
Dim temp As String
Dim cantidad As Integer

Select Case subn
Case 2, 4
    Call Politica80v1(p_ternro, Fecha, fecha_desde, hora_desde, fecha_hasta, hora_hasta, subn)
    Exit Sub
Case 5
    Call Politica80v2(p_ternro, Fecha, fecha_desde, hora_desde, fecha_hasta, hora_hasta, subn)
    Exit Sub
Case 6
    Call Politica80v3(p_ternro, Fecha, fecha_desde, hora_desde, fecha_hasta, hora_hasta, subn)
    Exit Sub
End Select

salir = False

' Busco en el Detalle del parte de autorizacion de Horarios
' una entrada para p_ternro

'FIND FIRST gti_regdet where PER.gti_regdet.ternro = p-ternro AND
'                            PER.gti_regdet.gprdfecdesde <= fecha AND
'                            fecha <= PER.gti_regdet.gprdfechasta no-lock no-error.
StrSql = "SELECT * FROM gti_regdet WHERE ternro = " & p_ternro & _
" and gprdfecdesde <= " & ConvFecha(Fecha) & " and gprdfechasta >= " & ConvFecha(Fecha) & _
" ORDER BY ternro ASC, gprdfecdesde ASC, gprdfechasta ASC"
rs.Open StrSql, objConn

If Not rs.EOF Then
    rs.MoveFirst
' MUY IMPORTANTE
' Chequear en la BD que los campos corresp a horas
' sean de tipo Fecha o con formato de hora
' Cambiar ConvHoraBD segun corresponda.
    
    ' Registracion - entrada

    StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,entrada,evenro)" & _
    " VALUES(" & p_ternro & "," & ConvFecha(rs!gprdfecdesde) & ",'" & _
    rs!grpdhoradesde & "',1,2)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Registracion - salida
    StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,entrada,evenro)" & _
    " VALUES(" & p_ternro & "," & ConvFecha(rs!gprdfechasta) & ",'" & _
    rs!grpdhorahasta & "',0,2)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    p_horario = True
    
End If
rs.Close

'  /* Generar el wf con cada registracion dentro de la ventana */

' Revisar: completar la consulta

StrSql = "SELECT gti_registracion.* " & _
"FROM gti_reloj INNER JOIN gti_registracion ON gti_reloj.relnro = gti_registracion.relnro" & _
" WHERE ternro= " & p_ternro & _
" AND ((regestado ='I') OR (regestado = 'S'))" & _
" AND regfecha >= " & ConvFecha(fecha_desde) & " AND regfecha <= " & _
ConvFecha(fecha_hasta) & _
" ORDER BY ternro ASC, regfecha ASC, reghora ASC"
' Permito modificaciones
rs.Open StrSql, objConn

Do While Not rs.EOF
'Revisar: fn de formateo de hora
    If rs!regfecha = fecha_desde Then
        ' Convierto para comparar
        If rs!reghora < hora_desde Then
        ' Proceso el proximo
        ' Con esto reemplazo al next
            salir = True
        End If
    End If
    If rs!regfecha = fecha_hasta And Not salir Then
        If rs!reghora > hora_hasta Then
        'Si la registracion queda fuera de la ventana termino de procesar
            Exit Do
        End If
    End If
 
    If Not salir Then
    
        If p_horario Then
    
            StrSql = "UPDATE gti_registracion SET regestado = 'S' WHERE regnro = " & rs!Regnro
            objConn.Execute StrSql, , adExecuteNoRecords
        
        Else
            
            StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
            " VALUES(" & rs!Ternro & "," & ConvFecha(rs!regfecha) & ",'" & _
            rs!reghora & "'," & rs!Regnro & _
            "," & IIf(rs!regentsal = "E", -1, 0) & ",2)"
            objConn.Execute StrSql, , adExecuteNoRecords
        
        End If
        
    Else
        salir = False
    
    End If
    rs.MoveNext
Loop

If rs.State = adStateOpen Then rs.Close

Set rs = Nothing

End Sub


Public Sub Politica200(nroter As Long, nro_turno As Long, p_fecha As Date, tol As String, acumula As Boolean, deb As Boolean, toldto As String, acumula_dto As Boolean, subn As Integer)
'Generacion de las salidas temprano

Dim rs_dia As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String
Dim th_tol As Integer
Dim horas_acum As Integer
Dim hora_red   As String
Dim hora_a_red As String
Dim tothoras   As Single
Dim tdias      As Integer
Dim thoras     As Integer
Dim tmin       As Integer
Dim horas_tol As Single
Dim fecha_tol As Date
Dim hora_tol  As String
Dim ok As Boolean
Dim conhredtipo As Integer
Dim i As Integer

Dim fecres As Date
Dim horres As String

Dim fecha_toldto As Date
Dim hora_toldto As String


StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 5 " & _
" AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
rs.Open StrSql, objConn

If Not rs.EOF Then
    th_tol = rs!thnro
    conhredtipo = rs!conhredondeo
Else
    If depurar Then GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Llegada Tarde para el Turno:", Str(nro_turno)
    rs.Close
    Set rs = Nothing
    Exit Sub
End If

StrSql = "SELECT * FROM " & TTempWFDia & " WHERE entrada = " & CInt(False) & " ORDER BY codigo ASC"

rs_dia.Open StrSql, objConn

Do While Not rs_dia.EOF
    
    StrSql = "SELECT * FROM " & TTempWFEmbudo & " WHERE codigo = " & rs_dia!Codigo
    
    rs.Close
    rs.Open StrSql, objConn

    StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE entrada = " & CInt(False) & " AND " & _
    " (( fecha > " & ConvFecha(rs!fecha_desde) & " ) OR " & _
    "  ( fecha = " & ConvFecha(rs!fecha_desde) & " and hora >= '" & rs!hora_desde & "')) and" & _
    " (( fecha < " & ConvFecha(rs!fecha_hasta) & " ) OR " & _
    "  ( fecha = " & ConvFecha(rs!fecha_hasta) & " and hora <= '" & rs!hora_hasta & "'))" & _
    " ORDER BY fecha ASC, hora ASC"
    
    rs.Close
    
    rs.Open StrSql, objConn, adOpenDynamic, adLockOptimistic
    
    If Not rs.EOF Then rs.MoveFirst
       
    Do While Not rs.EOF

        StrSql = " UPDATE " & TTempWFTurno & " SET anornro = 6 WHERE regnro = " & rs!Regnro
        objConn.Execute StrSql, , adExecuteNoRecords

        Call objFechasHoras.RestaXHoras(rs_dia!Fecha, rs_dia!Hora, tol, fecha_tol, hora_tol)
        ok = objFechasHoras.ValidarHora(hora_tol)
        If (rs!Fecha = fecha_tol) And (rs!Hora < hora_tol) And _
        (rs!Hora < rs_dia!Hora) Then

            If acumula = True Then
                Call objFechasHoras.RestaHs(rs!Fecha, rs!Hora, rs_dia!Fecha, rs_dia!Hora, tdias, thoras, tmin)
                tothoras = (tdias * 24) + (thoras + (tmin / 60))
                
                Call objFechasHoras.Convertir_A_Hora(tothoras * 60, hora_a_red)

                Call objFechasHoras.Redondeo_Horas_Tipo(hora_a_red, conhredtipo, horas_tol)
                
                If horas_tol <= 0 Then
                    GoTo Salida
                End If
                
                'Call ValidarTipoDeHora(th_tol, nro_turno, tipo_hora)
                
                StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
                "horvalido,ternro,thnro,turnro,empleg,horfecrep,horhoradesde,horhorahasta," & _
                "normnro,normnro2,horestado) VALUES (" & horas_tol & "," & _
                ConvFecha(fecha_tol) & "," & ConvFecha(fecha_tol) & "," & "0,-1," & _
                nroter & "," & th_tol & "," & nro_turno & "," & empleado.Legajo & "," & _
                ConvFecha(p_fecha) & ",'" & rs!Hora & "','" & rs_dia!Hora & "',6,6,'')"
                objConn.Execute StrSql, , adExecuteNoRecords
                    
            End If
            fecha_toldto = rs_dia!Fecha: hora_toldto = rs_dia!Hora
            Call objFechasHoras.SumoHoras(fecha_toldto, hora_toldto, toldto, fecres, horres)

            ok = objFechasHoras.ValidarHora(horres)
            
            If (rs!Fecha = fecres And rs!Hora > horres) And (acumula_dto) Then
                
                Call objFechasHoras.RestaHs(fecres, horres, rs!Fecha, rs!Hora, tdias, thoras, tmin)
                tothoras = (tdias * 24) + (thoras + (tmin / 60))
                
                If tothoras = 0 Then
                    GoTo Salida
                Else
                    Call DescontarHoras(rs!Fecha, tothoras, nro_turno, p_fecha)
                End If
                            
            End If
        Else
            
            If (rs!Fecha = fecha_tol) And (rs!Hora < rs_dia!Hora) Then

                StrSql = " UPDATE " & TTempWFTurno & " SET fecha = " & _
                ConvFecha(rs!Fecha) & ", hora = '" & rs_dia!Hora & _
                "' WHERE regnro = " & rs!Regnro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    
        rs.MoveNext
        
    Loop
Salida:
        rs_dia.MoveNext
Loop

rs.Close
Set rs = Nothing
rs_dia.Close
Set rs_dia = Nothing

End Sub


Private Sub DescontarHoras(fecha_dto As Date, horas_dto As Single, nro_turno As Long, p_fecha As Date)

Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim th_dto As Integer
Dim horas_acum As Single
Dim hora_red   As String
Dim hora_a_red As String
Dim limite As Integer



horas_acum = 0

'If horas_dto = 0 Then
    'Next
'End If

'  FIND FIRST per.gti_config-hora WHERE per.gti_config-hora.conhornro = 3 NO-LOCK NO-ERROR. /* Horas de Dto. */
'  FIND FIRST per.gti_config-tur-hora OF per.gti_config-hora WHERE per.gti_config-tur-hora.turnro = nro-turno NO-LOCK NO-ERROR.
' /* Horas de Dto. */
StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 3 " & _
" AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
rs.Open StrSql, objConn

If Not rs.EOF Then
    th_dto = rs!thnro
Else
    GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Descuento para el Turno:", Str(nro_turno)

'/* Deberia salir de todo el procedimiento */
    rs.Close
    Set rs = Nothing
    Exit Sub
End If

limite = rs!conhfracc / 60
Do While limite < horas_dto
    horas_acum = horas_acum + limite
    horas_dto = horas_dto - limite
Loop

horas_acum = horas_acum + limite
Call objFechasHoras.Convertir_A_Hora(horas_acum * 60, hora_a_red)

'  RUN gtired01.p(hora-a-red,per.gti_config-tur-hora.conhredtipo,OUTPUT horas-acum).
Call objFechasHoras.Redondeo_Horas_Tipo(hora_a_red, rs!conhredondeo, horas_acum)

'Call ValidarTipoDeHora(th_dto, nro_turno, tipo_hora)

StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
"horvalido,ternro,thnro,turnro,empleg,horfecrep" & _
") VALUES (" & horas_acum & "," & _
ConvFecha(fecha_dto) & "," & ConvFecha(fecha_dto) & "," & "0,-1," & _
empleado.Ternro & "," & th_dto & "," & nro_turno & "," & empleado.Legajo & "," & _
ConvFecha(p_fecha) & ")"

objConn.Execute StrSql, , adExecuteNoRecords

'  ASSIGN per.gti_horcumplido.horcant = horas - acum
'         per.gti_horcumplido.hordesde = Fecha - dto
'         per.gti_horcumplido.horestado = ""
'         per.gti_horcumplido.horhasta = Fecha - dto
'         per.gti_horcumplido.hormanual = False
'         per.gti_horcumplido.horvalido = True
'         per.gti_horcumplido.ternro = empleado.ternro
'         per.gti_horcumplido.thnro = th - dto
'         per.gti_horcumplido.turnro = nro - turno
'         per.gti_horcumplido.Empleg = empleado.Empleg
'         per.gti_horcumplido.horfecrep    = p-fecha.


End Sub

Public Sub Politica190(nroter As Long, nro_turno As Long, p_fecha As Date, tol As String, acumula As Boolean, deb As Boolean, toldto As String, acumula_dto As Boolean, subn As Integer)

'Generacion de llegadas tardes

Dim th_tol As Integer
Dim horas_acum As Integer
Dim hora_red   As String
Dim hora_a_red As String
Dim tothoras   As Single
Dim tdias      As Integer
Dim thoras     As Integer
Dim tmin       As Integer
Dim horas_tol As Single
Dim fecha_tol As Date
Dim hora_tol  As String
Dim ok As Boolean
Dim conhredtipo As Integer
Dim i As Integer

Dim fecha_toldto As Date
Dim hora_toldto As String

Dim fecres As Date
Dim horres As String

Dim rs_dia As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

i = -1

''Flog.writeline Now & " Dentro de la 190"

StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 4 " & _
" AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
OpenRecordset StrSql, rs

If Not rs.EOF Then
    th_tol = rs!thnro
    conhredtipo = rs!conhredondeo
Else
    If depurar Then GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Llegada Tarde para el Turno:", Str(nro_turno)
    Exit Sub
End If

'StrSql = "SELECT * FROM " & TTempWFdia & " WHERE entrada = " & CInt(True) & " ORDER BY codigo ASC"

StrSql = "SELECT " & TTempWFDia & ".Codigo,fecha,hora,fecha_desde,fecha_hasta,hora_desde,hora_hasta FROM " & TTempWFDia & " "
StrSql = StrSql & " INNER JOIN " & TTempWFEmbudo & " ON " & TTempWFDia & ".Codigo = " & TTempWFEmbudo & ".Codigo "
StrSql = StrSql & " WHERE " & TTempWFDia & ".entrada = " & CInt(True) & " ORDER BY " & TTempWFDia & ".codigo ASC"

''Flog.writeline Now & " Antes del open wf_dia"
OpenRecordset StrSql, rs_dia, adLockOptimistic

Do While Not rs_dia.EOF



    StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE entrada = " & CInt(True) & " AND " & _
    " (( fecha > " & ConvFecha(rs_dia!fecha_desde) & " ) OR " & _
    "  ( fecha = " & ConvFecha(rs_dia!fecha_desde) & " and hora >= '" & rs_dia!hora_desde & "')) and" & _
    " (( fecha < " & ConvFecha(rs_dia!fecha_hasta) & " ) OR " & _
    "  ( fecha = " & ConvFecha(rs_dia!fecha_hasta) & " and hora <= '" & rs_dia!hora_hasta & "'))" & _
    " ORDER BY fecha ASC, hora ASC"
    
    OpenRecordset StrSql, rs, adLockOptimistic
    
    If Not rs.EOF Then rs.MoveFirst
        
    Do While Not rs.EOF
        
        Call objFechasHoras.SumoHoras(rs_dia!Fecha, rs_dia!Hora, tol, fecha_tol, hora_tol)
        ok = objFechasHoras.ValidarHora(hora_tol)
        If (rs!Fecha = fecha_tol) And (rs!Hora > hora_tol) And _
        (rs!Hora > rs_dia!Hora) Then
        
            StrSql = " UPDATE " & TTempWFTurno & " SET anornro = 5 WHERE regnro = " & rs!Regnro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            If acumula = True Then
                
                Call objFechasHoras.RestaHs(rs_dia!Fecha, rs_dia!Hora, rs!Fecha, rs!Hora, tdias, thoras, tmin)
                tothoras = (tdias * 24) + (thoras + (tmin / 60))
                
                Call objFechasHoras.Convertir_A_Hora(tothoras * 60, hora_a_red)

                Call objFechasHoras.Redondeo_Horas_Tipo(hora_a_red, conhredtipo, horas_tol)
                
                If horas_tol <= 0 Then
                    GoTo Salida
                End If
                
                'Call ValidarTipoDeHora(th_tol, nro_turno, tipo_hora)
                
                StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
                "horvalido,ternro,thnro,turnro,empleg,horfecrep,horhoradesde,horhorahasta," & _
                "normnro,normnro2,horestado) VALUES (" & horas_tol & "," & _
                ConvFecha(fecha_tol) & "," & ConvFecha(fecha_tol) & "," & CInt(False) & "," & CInt(True) & "," & _
                nroter & "," & th_tol & "," & nro_turno & "," & empleado.Legajo & "," & _
                ConvFecha(p_fecha) & ",'" & rs_dia!Hora & "','" & rs!Hora & "',5,5,'')"
                objConn.Execute StrSql, , adExecuteNoRecords

            End If
            fecha_toldto = rs_dia!Fecha
            hora_toldto = rs_dia!Hora
            Call objFechasHoras.SumoHoras(fecha_toldto, hora_toldto, toldto, fecres, horres)

            ok = objFechasHoras.ValidarHora(horres)
            
            If (rs!Fecha = fecres And rs!Hora > horres) And (acumula_dto) Then
                Call objFechasHoras.RestaHs(fecres, horres, rs!Fecha, rs!Hora, tdias, thoras, tmin)
                tothoras = (tdias * 24) + (thoras + (tmin / 60))
                
                If tothoras = 0 Then
                    GoTo Salida
                Else
                    Call DescontarHoras(rs!Fecha, tothoras, nro_turno, p_fecha)
                End If
                            
            End If
        Else

            
            If (rs!Fecha = fecha_tol) And (rs!Hora > rs_dia!Hora) Then

                StrSql = " UPDATE " & TTempWFTurno & " SET fecha  = " & ConvFecha(rs_dia!Fecha) & ", hora = '" & rs_dia!Hora & "' WHERE regnro = " & rs!Regnro
                objConn.Execute StrSql, , adExecuteNoRecords
            
            End If
        End If
    
        rs.MoveNext
    
    Loop

Salida:
    rs_dia.MoveNext
    
Loop

' Cierro todo antes de salir
rs.Close
Set rs = Nothing
rs_dia.Close
Set rs_dia = Nothing


End Sub


Public Sub Politica480_RHPRO(nroter As Long, p_fecha As Date, subn As Integer)

' Generacion de justificaciones parciales variables

Dim aux_Tipohora As Integer

Dim aux_hora As Integer
Dim Hora As String
Dim fecres As Date
Dim horres As String
Dim ok As Boolean

Dim vhora  As String
Dim vfecha As Date
Dim vcant  As Single

Dim tothor As Single

Dim ActHora As Boolean

'Buffers
Dim wf_hcp As New ADODB.Recordset
Dim wf_hcp2 As New ADODB.Recordset

Dim rs_justif As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim Cantidad_a_Justificar As Single

'Call CreateTempWFJustif
Call CreateTempTable(TTempWFJustif)

StrSql = "SELECT * FROM GTI_JUSTIFICACION WHERE ternro = " & nroter & _
" AND juseltipo = 3 AND jusdesde <= " & ConvFecha(p_fecha) & " AND " & _
" jushasta >= " & ConvFecha(p_fecha) & " ORDER BY juselorden"

If rs_justif.State = adStateOpen Then rs_justif.Close
rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic

'OpenRecordset StrSql, rs_justif

Do While Not rs_justif.EOF

    Select Case rs_justif!jussigla
    
    Case "NOV"

        StrSql = "SELECT gti_tiponovedad.thnro FROM gti_novedad " & _
                 " INNER JOIN gti_tiponovedad ON gti_tiponovedad.gtnovnro = gti_novedad.gtnovnro " & _
                 " WHERE gti_novedad.gnovotoa = " & nroter & _
                 " AND gti_novedad.gnovnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If

        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords


    Case "LIC"
    
        StrSql = "SELECT tipdia.thnro FROM emp_lic " & _
                 " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
                 " WHERE emp_lic.empleado = " & nroter & _
        " AND emp_lic.emp_licnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If
        
        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End Select
    
    rs_justif.MoveNext

Loop

rs_justif.UpdateBatch
If rs_justif.State = adStateOpen Then rs_justif.Close

If rs.State = adStateOpen Then rs.Close

StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
" AND horfecrep = " & ConvFecha(p_fecha) & " AND ((normnro IS NOT NULL)" & _
" OR (normnro2 IS NOT NULL))"
 '" ORDER BY juselorden"
OpenRecordset StrSql, rs


Do While Not (rs.EOF)

    If Cantidad_a_Justificar = 0 Then
        Cantidad_a_Justificar = rs!horcant
    End If

    StrSql = "SELECT * FROM " & TTempWFJustif & " ORDER BY orden"
 
    If rs_justif.State = adStateOpen Then rs_justif.Close
    rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
   
    Do While Not rs_justif.EOF
        
        If Cantidad_a_Justificar > rs_justif!maxhoras Then
            
            StrSql = "SELECT * FROM GTI_HORCUMPLIDO WHERE hornro = " & rs!hornro
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            wf_hcp.Open StrSql, objConn
            
            If Not wf_hcp.EOF Then
                vcant = wf_hcp!horcant - rs_justif!maxhoras

                If (wf_hcp!horhorahasta = Null) Then
                    Call Calcula_Hora(vcant, wf_hcp!hordesde, vhora, vfecha)
                    Call objFechasHoras.RestaXHoras(wf_hcp!horfecrep, wf_hcp!horhorahasta, vhora, fecres, horres)
    
                    ok = objFechasHoras.ValidarHora(horres)
                    
                    Cantidad_a_Justificar = vcant
                    
                    StrSql = "update gti_horcumplido set horcant = " & vcant & _
                    ", horhoradesde = '" & horres & _
                    "' where hornro = " & rs!hornro
                Else
                    Cantidad_a_Justificar = vcant
                
                    StrSql = "update gti_horcumplido set horcant = " & vcant & _
                    " where hornro = " & rs!hornro
                End If
                objConn.Execute StrSql, adExecuteNoRecords
                
            End If
            
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn
            
            If wf_hcp2.EOF Then
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                Round(rs_justif!maxhoras, 2) & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & Round(wf_hcp2!horcant + rs_justif!maxhoras, 2) & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords

            GoTo NextProgress

        End If
        
        If Cantidad_a_Justificar = rs_justif!maxhoras Then
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE hornro = " & rs!hornro
            
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            
            wf_hcp.Open StrSql, objConn
            StrSql = "delete from gti_horcumplido where hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
                " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
                rs_justif!thnro
            
            
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn

            If wf_hcp2.EOF Then
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                Round(rs_justif!maxhoras, 2) & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & Round(wf_hcp2!horcant + rs_justif!maxhoras, 2) & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
            GoTo NextProgress

        End If
        
   If Cantidad_a_Justificar < rs_justif!maxhoras Then
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
        
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn

                        
            If wf_hcp2.EOF Then
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                Round(Cantidad_a_Justificar, 2) & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & Round(wf_hcp2!horcant + Cantidad_a_Justificar, 2) & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            
            StrSql = "UPDATE " & TTempWFJustif & " SET maxhoras = " & rs_justif!maxhoras - rs!horcant & _
            " WHERE nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
                       
                       
            StrSql = "DELETE FROM gti_horcumplido WHERE hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords

            GoTo NextProgress

        End If
               
            
        
NextProgress:

        rs_justif.MoveNext
    
    Loop
    
    rs.MoveNext
    
    Cantidad_a_Justificar = 0

Loop


CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    If rs_justif.State = adStateOpen Then
        rs_justif.UpdateBatch
        rs_justif.Close
    End If
    Set rs_justif = Nothing
    
    
StrSql = "DROP TABLE " & TTempWFJustif
objConn.Execute StrSql, adExecuteNoRecords
    
    
    
End Sub

Public Sub Politica480_vieja(nroter As Long, p_fecha As Date, subn As Integer)

' Generacion de justificaciones parciales variables

Dim aux_Tipohora As Integer

Dim aux_hora As Integer
Dim Hora As String
Dim fecres As Date
Dim horres As String
Dim ok As Boolean

Dim vhora  As String
Dim vfecha As Date
Dim vcant  As Single

Dim tothor As Single

Dim ActHora As Boolean

'Buffers
Dim wf_hcp As New ADODB.Recordset
Dim wf_hcp2 As New ADODB.Recordset

Dim rs_justif As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

'Call CreateTempWFJustif
Call CreateTempTable(TTempWFJustif)

StrSql = "SELECT * FROM GTI_JUSTIFICACION WHERE ternro = " & nroter & _
" AND juseltipo = 3 AND jusdesde <= " & ConvFecha(p_fecha) & " AND " & _
" jushasta >= " & ConvFecha(p_fecha) & " ORDER BY juselorden"

If rs_justif.State = adStateOpen Then rs_justif.Close
rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic

'OpenRecordset StrSql, rs_justif

Do While Not rs_justif.EOF

    Select Case rs_justif!jussigla
    
    Case "NOV"

        StrSql = "SELECT gti_tiponovedad.thnro FROM gti_novedad " & _
                 " INNER JOIN gti_tiponovedad ON gti_tiponovedad.gtnovnro = gti_novedad.gtnovnro " & _
                 " WHERE gti_novedad.gnovotoa = " & nroter & _
                 " AND gti_novedad.gnovnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If

        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords


    Case "LIC"
    
        StrSql = "SELECT tipdia.thnro FROM emp_lic " & _
                 " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
                 " WHERE emp_lic.empleado = " & nroter & _
        " AND emp_lic.emp_licnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If
        
        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End Select
    
    rs_justif.MoveNext

Loop

StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
" AND horfecrep = " & ConvFecha(p_fecha) & " AND ((normnro IS NOT NULL)" & _
" OR (normnro2 IS NOT NULL))"
 '" ORDER BY juselorden"
OpenRecordset StrSql, rs


rs_justif.UpdateBatch
If rs_justif.State = adStateOpen Then rs_justif.Close


Do While Not (rs.EOF)

    StrSql = "SELECT * FROM " & TTempWFJustif & " ORDER BY orden"
 
    If rs_justif.State = adStateOpen Then rs_justif.Close
    rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
   
    Do While Not rs_justif.EOF
        
        If rs!horcant > rs_justif!maxhoras Then
            
            StrSql = "SELECT * FROM GTI_HORCUMPLIDO WHERE hornro = " & rs!hornro
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            wf_hcp.Open StrSql, objConn
            
            If Not wf_hcp.EOF Then
                vcant = wf_hcp!horcant - rs_justif!maxhoras

                If (wf_hcp!horhorahasta = Null) Then
                    Call Calcula_Hora(vcant, wf_hcp!hordesde, vhora, vfecha)
                    Call objFechasHoras.RestaXHoras(wf_hcp!horfecrep, wf_hcp!horhorahasta, vhora, fecres, horres)
    
                    ok = objFechasHoras.ValidarHora(horres)
                    
                    StrSql = "update gti_horcumplido set horcant = " & vcant & _
                    ", horhoradesde = '" & horres & _
                    "' where hornro = " & rs!hornro
                Else
                    StrSql = "update gti_horcumplido set horcant = " & vcant & _
                    " where hornro = " & rs!hornro
                End If
                objConn.Execute StrSql, adExecuteNoRecords
                
            End If
            
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn
            
            If wf_hcp2.EOF Then
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                Round(rs_justif!maxhoras, 2) & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & Round(wf_hcp2!horcant + rs_justif!maxhoras, 2) & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords

            GoTo NextProgress

        End If
        
        If rs!horcant = rs_justif!maxhoras Then
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE hornro = " & rs!hornro
            
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            
            wf_hcp.Open StrSql, objConn
            StrSql = "delete from gti_horcumplido where hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
                " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
                rs_justif!thnro
            
            
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn

            If wf_hcp2.EOF Then
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                Round(rs_justif!maxhoras, 2) & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & Round(wf_hcp2!horcant + rs_justif!maxhoras, 2) & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
            GoTo NextProgress

        End If
        
   If rs!horcant < rs_justif!maxhoras Then
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
        
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn

                        
            If wf_hcp2.EOF Then
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                Round(rs!horcant, 2) & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & Round(wf_hcp2!horcant + rs!horcant, 2) & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            
            StrSql = "UPDATE " & TTempWFJustif & " SET maxhoras = " & rs_justif!maxhoras - rs!horcant & _
            " WHERE nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
                       
                       
            StrSql = "DELETE FROM gti_horcumplido WHERE hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords

            GoTo NextProgress

        End If
               
               
        
NextProgress:
        rs_justif.MoveNext
    Loop
    
    rs.MoveNext

Loop


CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    If rs_justif.State = adStateOpen Then
        rs_justif.UpdateBatch
        rs_justif.Close
    End If
    Set rs_justif = Nothing
    
    
StrSql = "DROP TABLE " & TTempWFJustif
objConn.Execute StrSql, adExecuteNoRecords
    
    
    
End Sub



Public Sub Politica480_nue(nroter As Long, p_fecha As Date, subn As Integer)

' Generacion de justificaciones parciales variables

Dim aux_Tipohora As Integer

Dim aux_hora As Integer
Dim Hora As String
Dim fecres As Date
Dim horres As String
Dim ok As Boolean

Dim vhora  As String
Dim vfecha As Date
Dim vcant  As Single

'Buffers
Dim wf_hcp As New ADODB.Recordset
Dim wf_hcp2 As New ADODB.Recordset

Dim rs_justif As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

'Call CreateTempWFJustif
Call CreateTempTable(TTempWFJustif)

StrSql = "SELECT * FROM GTI_JUSTIFICACION WHERE ternro = " & nroter & _
" AND juseltipo = 3 AND jusdesde <= " & ConvFecha(p_fecha) & " AND " & _
" jushasta >= " & ConvFecha(p_fecha) & " ORDER BY juselorden"

If rs_justif.State = adStateOpen Then rs_justif.Close
rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic

'OpenRecordset StrSql, rs_justif

Do While Not rs_justif.EOF

    Select Case rs_justif!jussigla
    
    Case "NOV"

        StrSql = "SELECT gti_tiponovedad.thnro FROM gti_novedad " & _
                 " INNER JOIN gti_tiponovedad ON gti_tiponovedad.gtnovnro = gti_novedad.gtnovnro " & _
                 " WHERE gti_novedad.gnovotoa = " & nroter & _
                 " AND gti_novedad.gnovnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If

        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords


    Case "LIC"
    
        StrSql = "SELECT tipdia.thnro FROM emp_lic " & _
                 " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
                 " WHERE emp_lic.empleado = " & nroter & _
        " AND emp_lic.emp_licnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If
        
        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End Select
    
    rs_justif.MoveNext

Loop

StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
" AND horfecrep = " & ConvFecha(p_fecha) & " AND (normnro IS NOT NULL)" & _
" AND (normnro2 IS NOT NULL)"
'" ORDER BY juselorden"
OpenRecordset StrSql, rs


rs_justif.UpdateBatch
If rs_justif.State = adStateOpen Then rs_justif.Close


Do While Not (rs.EOF)

    StrSql = "SELECT * FROM " & TTempWFJustif & " ORDER BY orden"
 
    If rs_justif.State = adStateOpen Then rs_justif.Close
    rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
   
'    OpenRecordset StrSql, rs_justif
    
    Do While Not rs_justif.EOF
        
        If rs_justif!thnro = 0 Then GoTo NextProgress
        
        If rs!horcant > rs_justif!maxhoras Then
            ' Abro el buffer hcp
            StrSql = "SELECT * FROM GTI_HORCUMPLIDO WHERE hornro = " & rs!hornro
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            wf_hcp.Open StrSql, objConn
            
            If Not wf_hcp.EOF Then
                vcant = CSng("0" & wf_hcp!horcant) - rs_justif!maxhoras
                'vcant = CSng("0" & 0.75) - rs_justif!maxhoras
                Call Calcula_Hora(vcant, wf_hcp!hordesde, vhora, vfecha)
                Call objFechasHoras.RestaXHoras(wf_hcp!horfecrep, wf_hcp!horhorahasta, vhora, fecres, horres)

                'run pasarHr.p(horres, output horres, output ok).
                ok = objFechasHoras.ValidarHora(horres)
                'wf_hcp!horcant = CSng("0" & wf_hcp!horcant) - rs_justif!maxhoras
                'wf_hcp!horhoradesde = horres
                
                StrSql = "update gti_horcumplido set horcant = " & vcant & _
                ", horhoradesde = '" & horres & _
                "' where hornro = " & rs!hornro
                objConn.Execute StrSql, adExecuteNoRecords
                
            End If
            'wf_hcp.UpdateBatch
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            
'     FIND FIRST wf-hcp2 WHERE wf-hcp2.ternro    = nroter AND
'                              wf-hcp2.horfecrep = p-fecha AND
'                              wf-hcp2.thnro     = wf-justif.thnro EXCLUSIVE-LOCK NO-ERROR.
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            'wf_hcp2.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
            wf_hcp2.Open StrSql, objConn
            ' Si no hay uno que cumpla con las condiciones lo creo
            'If wf_hcp2.EOF Then
            '    wf_hcp2.AddNew
            'End If
            
            'wf_hcp2!horcant = CSng("0" & wf_hcp2!horcant) + rs_justif!maxhoras
            'wf_hcp2!hordesde = p_fecha
            'wf_hcp2!horhasta = p_fecha
            'wf_hcp2!horhoradesde = rs_justif!horadesde
            'wf_hcp2!horhorahasta = rs_justif!horahasta
            'wf_hcp2!hormanual = 0
            'wf_hcp2!horvalido = -1
            'wf_hcp2!Ternro = empleado.Ternro
            'wf_hcp2!thnro = rs_justif!thnro
            'wf_hcp2!Empleg = empleado.Legajo
            'wf_hcp2!horfecrep = p_fecha
            'wf_hcp2!horestado = ""
        
            'Grabo y cierro el recordset
            'wf_hcp2.UpdateBatch
            'wf_hcp2.Close
            
            'DELETE wf-justif.
            'rs_justif.Delete
            'StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            'objConn.Execute StrSql, , adExecuteNoRecords
            
            If wf_hcp2.EOF Then
                'wf_hcp2.AddNew
                'Call ValidarTipoDeHora(rs_justif!thnro, nro_turno, tipo_hora)
                
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                rs_justif!maxhoras & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & CSng("0" & wf_hcp2!horcant) + rs_justif!maxhoras & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Grabo y cierro el recordset
            wf_hcp2.Close
            
            'DELETE wf-justif.
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords

            'Next.
            GoTo NextProgress

        End If
        
        If rs!horcant = rs_justif!maxhoras Then
'     FIND FIRST wf-hcp WHERE RECID(wf-hcp) = RECID(gti_horcumplido) EXCLUSIVE-LOCK NO-ERROR.
            StrSql = "SELECT * FROM gti_horcumplido WHERE hornro = " & rs!hornro
            
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            'wf_hcp.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
            
            wf_hcp.Open StrSql, objConn
            StrSql = "delete from gti_horcumplido where hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords
            
'     FIND FIRST wf-hcp2 WHERE wf-hcp2.ternro    = nroter AND
'                              wf-hcp2.horfecrep = p-fecha AND
'                              wf-hcp2.thnro     = wf-justif.thnro EXCLUSIVE-LOCK NO-ERROR.
            'StrSql = "SELECT * FROM GTI_HORCUMPLIDO WHERE ternro = " & nroter & _
            '" AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            'rs_justif!thnro
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
                " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
                rs_justif!thnro
            
            
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn

            ' Si no hay uno que cumpla con las condiciones lo creo
            If wf_hcp2.EOF Then
                'wf_hcp2.AddNew
                'Call ValidarTipoDeHora(rs_justif!thnro, nro_turno, tipo_hora)
                
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                rs_justif!maxhoras & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & CSng("0" & wf_hcp2!horcant) + rs_justif!maxhoras & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            'DELETE wf-justif.
            'rs_justif.Delete
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
            'Next.
            GoTo NextProgress

        End If
        
   If rs!horcant < rs_justif!maxhoras Then
'     FIND FIRST wf-hcp2 WHERE wf-hcp2.ternro    = nroter AND
'                              wf-hcp2.horfecrep = p-fecha AND
'                              wf-hcp2.thnro     = wf-justif.thnro EXCLUSIVE-LOCK NO-ERROR.
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
        
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            'wf_hcp2.Open StrSql, objConn, adOpenKeyset, adLockBatchOptimistic
            wf_hcp2.Open StrSql, objConn
                        
            ' Si no hay uno que cumpla con las condiciones lo creo
            If wf_hcp2.EOF Then
                'wf_hcp2.AddNew
                'Call ValidarTipoDeHora(rs_justif!thnro, nro_turno, tipo_hora)
                
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                rs!horcant & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & CSng("0" & wf_hcp2!horcant) + rs!horcant & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'wf_hcp2!horcant = CSng("0" & wf_hcp2!horcant) + rs!horcant
            'wf_hcp2!hordesde = Date
            'wf_hcp2!horhasta = Date
            'wf_hcp2!horhoradesde = rs_justif!horadesde
            'wf_hcp2!horhorahasta = rs_justif!horahasta
            'wf_hcp2!hormanual = CInt(False)
            'wf_hcp2!horvalido = CInt(True)
            'wf_hcp2!Ternro = empleado.Ternro
            'wf_hcp2!thnro = rs_justif!thnro
            'wf_hcp2!Empleg = empleado.Legajo
            'wf_hcp2!horfecrep = Date
            'wf_hcp2!horestado = ""
        
            'Grabo y cierro el recordset
            'wf_hcp2.UpdateBatch
            wf_hcp2.Close
            
            'rs_justif!maxhoras = CSng("0" & rs_justif!maxhoras) - rs!horcant
            
            StrSql = "UPDATE " & TTempWFJustif & " SET maxhoras = " & rs_justif!maxhoras - rs!horcant & _
            " WHERE nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
                       
                       
            'StrSql = "SELECT * FROM gti_horcumplido WHERE (hornro = " & rs!hornro & ")"
        

            'If wf_hcp.State = adStateOpen Then
            '        wf_hcp.UpdateBatch
            '        wf_hcp.Close
            'End If
            'wf_hcp.Open StrSql, objConn, adOpenDynamic, adLockOptimistic
            
            'Borro y cierro el recordset
           ' wf_hcp.Delete
            'wf_hcp.UpdateBatch
            
            StrSql = "DELETE FROM gti_horcumplido WHERE hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords

            'wf_hcp.Close
                
            'Next.
            GoTo NextProgress

        End If
               
               
        
NextProgress:
        rs_justif.MoveNext
    Loop
    
    rs.MoveNext

Loop


CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    If rs_justif.State = adStateOpen Then
        rs_justif.UpdateBatch
        rs_justif.Close
    End If
    Set rs_justif = Nothing
    
    
StrSql = "DROP TABLE " & TTempWFJustif
objConn.Execute StrSql, adExecuteNoRecords
    
    
    
End Sub


Public Sub Politica480(nroter As Long, p_fecha As Date, subn As Integer)
' ---------------------------------------------------------------------
' Descripcion: Generacion de justificaciones parciales variables
' Autor: ?
' Fecha:
' Ultima Mod: No Andaba Bien, tenia problemas de avance de recordsets y estaba muy enquilombada.
'             Muchas manos ...
' Autor: FGZ
' Fecha Ultima Mod: 01/09/2003
' ---------------------------------------------------------------------

Dim aux_Tipohora As Integer
Dim aux_hora As Integer
Dim Hora As String
Dim fecres As Date
Dim horres As String
Dim ok As Boolean
Dim vhora  As String
Dim vfecha As Date
Dim vcant  As Single
'Buffers
Dim wf_hcp As New ADODB.Recordset
Dim wf_hcp2 As New ADODB.Recordset
Dim rs_justif As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String
Dim Sigue As Boolean



'Call CreateTempWFJustif
Call CreateTempTable(TTempWFJustif)

StrSql = "SELECT * FROM GTI_JUSTIFICACION WHERE ternro = " & nroter & _
" AND juseltipo = 3 AND jusdesde <= " & ConvFecha(p_fecha) & " AND " & _
" jushasta >= " & ConvFecha(p_fecha) & " ORDER BY juselorden"

If rs_justif.State = adStateOpen Then rs_justif.Close
rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic

'OpenRecordset StrSql, rs_justif

Do While Not rs_justif.EOF

    Select Case rs_justif!jussigla
    
    Case "NOV"

        StrSql = "SELECT gti_tiponovedad.thnro FROM gti_novedad " & _
                 " INNER JOIN gti_tiponovedad ON gti_tiponovedad.gtnovnro = gti_novedad.gtnovnro " & _
                 " WHERE gti_novedad.gnovotoa = " & nroter & _
                 " AND gti_novedad.gnovnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If

        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords


    Case "LIC"
    
        StrSql = "SELECT tipdia.thnro FROM emp_lic " & _
                 " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
                 " WHERE emp_lic.empleado = " & nroter & _
        " AND emp_lic.emp_licnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If
        
        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End Select
    
    rs_justif.MoveNext

Loop

StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
" AND horfecrep = " & ConvFecha(p_fecha) & " AND (normnro IS NOT NULL)" & _
" AND (normnro2 IS NOT NULL)"
'" ORDER BY juselorden"
OpenRecordset StrSql, rs


rs_justif.UpdateBatch
If rs_justif.State = adStateOpen Then rs_justif.Close


'Do While Not (rs.EOF)

Sigue = True
' FGZ -  01/09/2003
' Si No Hay horas a Justificar ==> Termino
If Not rs.EOF Then
    StrSql = "SELECT * FROM " & TTempWFJustif & " ORDER BY orden"
 
    If rs_justif.State = adStateOpen Then rs_justif.Close
    rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
    ' OpenRecordset StrSql, rs_justif
    
    ' Miemtras tenga horas para justificar y tenga Justificaciones
    Do While Not rs_justif.EOF And Not rs.EOF And Sigue
        
        If rs_justif!thnro <> 0 Then
        
            ' 1_ Justificacion Menor que las Horas generadas
            If rs!horcant > rs_justif!maxhoras Then
                ' Abro el buffer hcp
                StrSql = "SELECT * FROM GTI_HORCUMPLIDO WHERE hornro = " & rs!hornro
                If wf_hcp.State = adStateOpen Then wf_hcp.Close
                wf_hcp.Open StrSql, objConn
                
                If Not wf_hcp.EOF Then
                    vcant = CSng("0" & wf_hcp!horcant) - rs_justif!maxhoras
                    'vcant = CSng("0" & 0.75) - rs_justif!maxhoras
                    Call Calcula_Hora(vcant, wf_hcp!hordesde, vhora, vfecha)
                    Call objFechasHoras.RestaXHoras(wf_hcp!horfecrep, wf_hcp!horhorahasta, vhora, fecres, horres)
    
                    'run pasarHr.p(horres, output horres, output ok).
                    ok = objFechasHoras.ValidarHora(horres)
                    'wf_hcp!horcant = CSng("0" & wf_hcp!horcant) - rs_justif!maxhoras
                    'wf_hcp!horhoradesde = horres
                    
                    StrSql = "update gti_horcumplido set horcant = " & vcant & _
                    ", horhoradesde = '" & horres & _
                    "' where hornro = " & rs!hornro
                    objConn.Execute StrSql, adExecuteNoRecords
                    
                End If
                'wf_hcp.UpdateBatch
                If wf_hcp.State = adStateOpen Then wf_hcp.Close
                
                StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
                " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
                rs_justif!thnro
                If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
                'wf_hcp2.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
                wf_hcp2.Open StrSql, objConn
                ' Si no hay uno que cumpla con las condiciones lo creo
                
                If wf_hcp2.EOF Then
                    'wf_hcp2.AddNew
                    'Call ValidarTipoDeHora(rs_justif!thnro, nro_turno, tipo_hora)
                    
                    StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                    "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                    "horfecrep,horestado) values (" & _
                    rs_justif!maxhoras & "," & _
                    ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                    rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                    empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                    "," & ConvFecha(p_fecha) & ",'')"
                Else
                    StrSql = "UPDATE gti_horcumplido set horcant = " & CSng("0" & wf_hcp2!horcant) + rs_justif!maxhoras & _
                    ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                    ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                    "',horhorahasta='" & rs_justif!horahasta & _
                    "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                    ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                    ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                    wf_hcp2!hornro
                    
                End If
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'Grabo y cierro el recordset
                wf_hcp2.Close
                
                'DELETE wf-justif.
                StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
                objConn.Execute StrSql, , adExecuteNoRecords
    
                ' FGZ  - 01/09/2003
                ' NO Avanzo el rs porque me quedó saldo
                ' rs.MoveNext
                
                ' FGZ  - 01/09/2003
                ' Avanzo el wfJustif porque Eliminé de wf_Justif
                rs_justif.MoveNext
                
            Else 'If rs!horcant > rs_justif!maxhoras Then
                ' 2_ Justificacion = que las Horas generadas
                If rs!horcant = rs_justif!maxhoras Then
                    'FIND FIRST wf-hcp WHERE RECID(wf-hcp) = RECID(gti_horcumplido) EXCLUSIVE-LOCK NO-ERROR.
                    StrSql = "SELECT * FROM gti_horcumplido WHERE hornro = " & rs!hornro
                    
                    If wf_hcp.State = adStateOpen Then wf_hcp.Close
                    'wf_hcp.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
                    
                    wf_hcp.Open StrSql, objConn
                    StrSql = "DELETE FROM gti_horcumplido WHERE hornro = " & rs!hornro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
                        " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
                        rs_justif!thnro
                    
                    If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
                    wf_hcp2.Open StrSql, objConn
        
                    ' Si no hay uno que cumpla con las condiciones lo creo
                    If wf_hcp2.EOF Then
                        'wf_hcp2.AddNew
                        'Call ValidarTipoDeHora(rs_justif!thnro, nro_turno, tipo_hora)
                        
                        StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                        "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                        "horfecrep,horestado) values (" & _
                        rs_justif!maxhoras & "," & _
                        ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                        rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                        empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                        "," & ConvFecha(p_fecha) & ",'')"
                    Else
                        StrSql = "UPDATE gti_horcumplido set horcant = " & CSng("0" & wf_hcp2!horcant) + rs_justif!maxhoras & _
                        ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                        ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                        "',horhorahasta='" & rs_justif!horahasta & _
                        "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                        ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                        ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                        wf_hcp2!hornro
                        
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    wf_hcp2.Close
                    
                    StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'Next.
                    
                    ' FGZ  - 01/09/2003
                    ' Avanzo el rs porque eliminé el registro de GTI_HORCUMPLIDO
                    rs.MoveNext
                    
                    ' FGZ  - 01/09/2003
                    ' Avanzo el wfJustif porque Eliminé de wf_Justif
                    rs_justif.MoveNext
                    
                Else ' If rs!horcant = rs_justif!maxhoras Then
                    ' 3_ Justificacion Mayor que las Horas generadas
                    If rs!horcant < rs_justif!maxhoras Then
                            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
                            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
                            rs_justif!thnro
                        
                            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
                            wf_hcp2.Open StrSql, objConn
                                        
                            ' Si no hay uno que cumpla con las condiciones lo creo
                            If wf_hcp2.EOF Then
                                'wf_hcp2.AddNew
                                'Call ValidarTipoDeHora(rs_justif!thnro, nro_turno, tipo_hora)
                                
                                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                                "horfecrep,horestado) values (" & _
                                rs!horcant & "," & _
                                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                                "," & ConvFecha(p_fecha) & ",'')"
                            Else
                                StrSql = "UPDATE gti_horcumplido set horcant = " & CSng("0" & wf_hcp2!horcant) + rs!horcant & _
                                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                                "',horhorahasta='" & rs_justif!horahasta & _
                                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                                wf_hcp2!hornro
                            End If
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            StrSql = "UPDATE " & TTempWFJustif & " SET maxhoras = " & rs_justif!maxhoras - rs!horcant & _
                            " WHERE nro = " & rs_justif!nro
                            objConn.Execute StrSql, , adExecuteNoRecords
                                       
                            StrSql = "DELETE FROM gti_horcumplido WHERE hornro = " & rs!hornro
                            objConn.Execute StrSql, , adExecuteNoRecords
                
                            ' FGZ  - 01/09/2003
                            ' Avanzo el rs porque eliminé el registro de GTI_HORCUMPLIDO
                            rs.MoveNext
                            
                            ' FGZ  - 01/09/2003
                            ' No Avanzo el wfJustif porque me quedó saldo a justificar
                            ' rs_justif.MoveNext
                            
                    End If ' If rs!horcant < rs_justif!maxhoras Then
                End If ' If rs!horcant = rs_justif!maxhoras Then
            End If 'If rs!horcant > rs_justif!maxhoras Then
        
        Else 'If rs_justif!thnro <> 0 Then
            rs_justif.MoveNext
        End If
    
        ' Si no tengo mas Justificaciones o No tengo mas horas a Justificar ==> NO Continuo
        If rs.EOF Or rs_justif.EOF Then
            Sigue = False
        End If
    
    Loop 'wfJustif
    
End If
'Loop 'rs (GTI_HORCUMPLIDO)


CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    If rs_justif.State = adStateOpen Then
        rs_justif.UpdateBatch
        rs_justif.Close
    End If
    Set rs_justif = Nothing
    
    
StrSql = "DROP TABLE " & TTempWFJustif
objConn.Execute StrSql, adExecuteNoRecords
    
    
    
End Sub


Public Sub Politica480_2(nroter As Long, p_fecha As Date, subn As Integer)

' Generacion de justificaciones parciales variables

Dim aux_Tipohora As Integer

Dim aux_hora As Integer
Dim Hora As String
Dim fecres As Date
Dim horres As String
Dim ok As Boolean

Dim vhora  As String
Dim vfecha As Date
Dim vcant  As Single

Dim tothor As Single

Dim ActHora As Boolean

'Buffers
Dim wf_hcp As New ADODB.Recordset
Dim wf_hcp2 As New ADODB.Recordset

Dim rs_justif As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim Cantidad_a_Justificar As Single

'Call CreateTempWFJustif
Call CreateTempTable(TTempWFJustif)

StrSql = "SELECT * FROM GTI_JUSTIFICACION WHERE ternro = " & nroter & _
" AND juseltipo = 3 AND jusdesde <= " & ConvFecha(p_fecha) & " AND " & _
" jushasta >= " & ConvFecha(p_fecha) & " ORDER BY juselorden"

If rs_justif.State = adStateOpen Then rs_justif.Close
rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic

'OpenRecordset StrSql, rs_justif

Do While Not rs_justif.EOF

    Select Case rs_justif!jussigla
    
    Case "NOV"

        StrSql = "SELECT gti_tiponovedad.thnro FROM gti_novedad " & _
                 " INNER JOIN gti_tiponovedad ON gti_tiponovedad.gtnovnro = gti_novedad.gtnovnro " & _
                 " WHERE gti_novedad.gnovotoa = " & nroter & _
                 " AND gti_novedad.gnovnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If

        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords


    Case "LIC"
    
        StrSql = "SELECT tipdia.thnro FROM emp_lic " & _
                 " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
                 " WHERE emp_lic.empleado = " & nroter & _
        " AND emp_lic.emp_licnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If
        
        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End Select
    
    rs_justif.MoveNext

Loop

rs_justif.UpdateBatch
If rs_justif.State = adStateOpen Then rs_justif.Close

If rs.State = adStateOpen Then rs.Close

StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
" AND horfecrep = " & ConvFecha(p_fecha) & " AND ((normnro IS NOT NULL)" & _
" OR (normnro2 IS NOT NULL))"
 '" ORDER BY juselorden"
OpenRecordset StrSql, rs


Do While Not (rs.EOF)

    If Cantidad_a_Justificar = 0 Then
        Cantidad_a_Justificar = rs!horcant
    End If

    StrSql = "SELECT * FROM " & TTempWFJustif & " ORDER BY orden"
 
    If rs_justif.State = adStateOpen Then rs_justif.Close
    rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
   
    Do While Not rs_justif.EOF
        
        If Cantidad_a_Justificar > rs_justif!maxhoras Then
            
            StrSql = "SELECT * FROM GTI_HORCUMPLIDO WHERE hornro = " & rs!hornro
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            wf_hcp.Open StrSql, objConn
            
            If Not wf_hcp.EOF Then
                vcant = wf_hcp!horcant - rs_justif!maxhoras

                If (wf_hcp!horhorahasta = Null) Then
                    Call Calcula_Hora(vcant, wf_hcp!hordesde, vhora, vfecha)
                    Call objFechasHoras.RestaXHoras(wf_hcp!horfecrep, wf_hcp!horhorahasta, vhora, fecres, horres)
    
                    ok = objFechasHoras.ValidarHora(horres)
                    
                    Cantidad_a_Justificar = vcant
                    
                    StrSql = "update gti_horcumplido set horcant = " & vcant & _
                    ", horhoradesde = '" & horres & _
                    "' where hornro = " & rs!hornro
                Else
                    Cantidad_a_Justificar = vcant
                
                    StrSql = "update gti_horcumplido set horcant = " & vcant & _
                    " where hornro = " & rs!hornro
                End If
                objConn.Execute StrSql, adExecuteNoRecords
                
            End If
            
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn
            
            If wf_hcp2.EOF Then
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                Round(rs_justif!maxhoras, 2) & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & Round(wf_hcp2!horcant + rs_justif!maxhoras, 2) & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords

            GoTo NextProgress

        End If
        
        If Cantidad_a_Justificar = rs_justif!maxhoras Then
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE hornro = " & rs!hornro
            
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            
            wf_hcp.Open StrSql, objConn
            StrSql = "delete from gti_horcumplido where hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
                " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
                rs_justif!thnro
            
            
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn

            If wf_hcp2.EOF Then
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                Round(rs_justif!maxhoras, 2) & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & Round(wf_hcp2!horcant + rs_justif!maxhoras, 2) & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
            GoTo NextProgress

        End If
        
   If Cantidad_a_Justificar < rs_justif!maxhoras Then
            
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
        
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn

                        
            If wf_hcp2.EOF Then
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                Round(Cantidad_a_Justificar, 2) & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & Round(wf_hcp2!horcant + Cantidad_a_Justificar, 2) & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            
            StrSql = "UPDATE " & TTempWFJustif & " SET maxhoras = " & rs_justif!maxhoras - rs!horcant & _
            " WHERE nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
                       
                       
            StrSql = "DELETE FROM gti_horcumplido WHERE hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords

            GoTo NextProgress

        End If
               
            
        
NextProgress:

        rs_justif.MoveNext
    
    Loop
    
    rs.MoveNext
    
    Cantidad_a_Justificar = 0

Loop


CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    If rs_justif.State = adStateOpen Then
        rs_justif.UpdateBatch
        rs_justif.Close
    End If
    Set rs_justif = Nothing
    
    
StrSql = "DROP TABLE " & TTempWFJustif
objConn.Execute StrSql, adExecuteNoRecords
    
    
    
End Sub


Public Sub Politica480_Expo(nroter As Long, p_fecha As Date, subn As Integer)

' Generacion de justificaciones parciales variables

Dim aux_Tipohora As Integer

Dim aux_hora As Integer
Dim Hora As String
Dim fecres As Date
Dim horres As String
Dim ok As Boolean

Dim vhora  As String
Dim vfecha As Date
Dim vcant  As Single

'Buffers
Dim wf_hcp As New ADODB.Recordset
Dim wf_hcp2 As New ADODB.Recordset

Dim rs_justif As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

'Call CreateTempWFJustif
Call CreateTempTable(TTempWFJustif)

StrSql = "SELECT * FROM GTI_JUSTIFICACION WHERE ternro = " & nroter & _
" AND juseltipo = 3 AND jusdesde <= " & ConvFecha(p_fecha) & " AND " & _
" jushasta >= " & ConvFecha(p_fecha) & " ORDER BY juselorden"

If rs_justif.State = adStateOpen Then rs_justif.Close
rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic

'OpenRecordset StrSql, rs_justif

Do While Not rs_justif.EOF

    Select Case rs_justif!jussigla
    
    Case "NOV"

        StrSql = "SELECT gti_tiponovedad.thnro FROM gti_novedad " & _
                 " INNER JOIN gti_tiponovedad ON gti_tiponovedad.gtnovnro = gti_novedad.gtnovnro " & _
                 " WHERE gti_novedad.gnovotoa = " & nroter & _
                 " AND gti_novedad.gnovnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Novedad, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If

        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords


    Case "LIC"
    
        StrSql = "SELECT tipdia.thnro FROM emp_lic " & _
                 " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
                 " WHERE emp_lic.empleado = " & nroter & _
        " AND emp_lic.emp_licnro = " & rs_justif!juscodext
        
        OpenRecordset StrSql, rs
        If rs.EOF Then
'            MsgBox "Hay problemas con la Licencia, avise al soporte tecnico de HEIDT & ASOC. (gtiprc30)", vbCritical
            GoTo CierroTodo
        Else
            aux_Tipohora = rs!thnro
        End If
        
        StrSql = "INSERT INTO " & TTempWFJustif & " (codext,desde,hasta,maxhoras,orden," & _
        "nro,sigla,ternro,tjusnro,horadesde,horahasta,thnro) " & _
        " VALUES(" & rs_justif!juscodext & "," & ConvFecha(rs_justif!jusdesde) & _
        "," & ConvFecha(rs_justif!jushasta) & _
        "," & rs_justif!juselmaxhoras & "," & rs_justif!juselorden & "," & rs_justif!jusnro & ",'" & rs_justif!jussigla & _
        "'," & rs_justif!Ternro & "," & rs_justif!tjusnro & ",'" & _
        rs_justif!jushoradesde & "','" & _
        rs_justif!jushorahasta & "'," & aux_Tipohora & ")"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End Select
    
    rs_justif.MoveNext

Loop

StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
" AND horfecrep = " & ConvFecha(p_fecha) & " AND (normnro IS NOT NULL)" & _
" AND (normnro2 IS NOT NULL)"
'" ORDER BY juselorden"
OpenRecordset StrSql, rs


rs_justif.UpdateBatch
If rs_justif.State = adStateOpen Then rs_justif.Close


Do While Not (rs.EOF)

    StrSql = "SELECT * FROM " & TTempWFJustif & " ORDER BY orden"
 
    If rs_justif.State = adStateOpen Then rs_justif.Close
    rs_justif.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
   
'    OpenRecordset StrSql, rs_justif
    
    Do While Not rs_justif.EOF
        ' FGZ
        ' ojo -- modifique esta linea 17/07/2003
        'If rs_justif!thnro = 0 Then GoTo NextProgress
        
        If rs!horcant > rs_justif!maxhoras Then
            ' Abro el buffer hcp
            StrSql = "SELECT * FROM GTI_HORCUMPLIDO WHERE hornro = " & rs!hornro
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            wf_hcp.Open StrSql, objConn
            
            If Not wf_hcp.EOF Then
                vcant = CSng("0" & wf_hcp!horcant) - rs_justif!maxhoras
                'vcant = CSng("0" & 0.75) - rs_justif!maxhoras
                Call Calcula_Hora(vcant, wf_hcp!hordesde, vhora, vfecha)
                Call objFechasHoras.RestaXHoras(wf_hcp!horfecrep, wf_hcp!horhorahasta, vhora, fecres, horres)

                'run pasarHr.p(horres, output horres, output ok).
                ok = objFechasHoras.ValidarHora(horres)
                'wf_hcp!horcant = CSng("0" & wf_hcp!horcant) - rs_justif!maxhoras
                'wf_hcp!horhoradesde = horres
                
                StrSql = "update gti_horcumplido set horcant = " & vcant & _
                ", horhoradesde = '" & horres & _
                "' where hornro = " & rs!hornro
                objConn.Execute StrSql, adExecuteNoRecords
                
            End If
            'wf_hcp.UpdateBatch
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            
'     FIND FIRST wf-hcp2 WHERE wf-hcp2.ternro    = nroter AND
'                              wf-hcp2.horfecrep = p-fecha AND
'                              wf-hcp2.thnro     = wf-justif.thnro EXCLUSIVE-LOCK NO-ERROR.
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            'wf_hcp2.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
            wf_hcp2.Open StrSql, objConn
            ' Si no hay uno que cumpla con las condiciones lo creo
            'If wf_hcp2.EOF Then
            '    wf_hcp2.AddNew
            'End If
            
            'wf_hcp2!horcant = CSng("0" & wf_hcp2!horcant) + rs_justif!maxhoras
            'wf_hcp2!hordesde = p_fecha
            'wf_hcp2!horhasta = p_fecha
            'wf_hcp2!horhoradesde = rs_justif!horadesde
            'wf_hcp2!horhorahasta = rs_justif!horahasta
            'wf_hcp2!hormanual = 0
            'wf_hcp2!horvalido = -1
            'wf_hcp2!Ternro = empleado.Ternro
            'wf_hcp2!thnro = rs_justif!thnro
            'wf_hcp2!Empleg = empleado.Legajo
            'wf_hcp2!horfecrep = p_fecha
            'wf_hcp2!horestado = ""
        
            'Grabo y cierro el recordset
            'wf_hcp2.UpdateBatch
            'wf_hcp2.Close
            
            'DELETE wf-justif.
            'rs_justif.Delete
            'StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            'objConn.Execute StrSql, , adExecuteNoRecords
            
            If wf_hcp2.EOF Then
                'wf_hcp2.AddNew
                'Call ValidarTipoDeHora(rs_justif!thnro, nro_turno, tipo_hora)
                
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                rs_justif!maxhoras & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & CSng("0" & wf_hcp2!horcant) + rs_justif!maxhoras & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Grabo y cierro el recordset
            wf_hcp2.Close
            
            'DELETE wf-justif.
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords

            'Next.
            GoTo NextProgress

        End If
        
        If rs!horcant = rs_justif!maxhoras Then
'     FIND FIRST wf-hcp WHERE RECID(wf-hcp) = RECID(gti_horcumplido) EXCLUSIVE-LOCK NO-ERROR.
            StrSql = "SELECT * FROM gti_horcumplido WHERE hornro = " & rs!hornro
            
            If wf_hcp.State = adStateOpen Then wf_hcp.Close
            'wf_hcp.Open StrSql, objConn, adOpenKeyset, adLockOptimistic
            
            wf_hcp.Open StrSql, objConn
            StrSql = "delete from gti_horcumplido where hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords
            
'     FIND FIRST wf-hcp2 WHERE wf-hcp2.ternro    = nroter AND
'                              wf-hcp2.horfecrep = p-fecha AND
'                              wf-hcp2.thnro     = wf-justif.thnro EXCLUSIVE-LOCK NO-ERROR.
            'StrSql = "SELECT * FROM GTI_HORCUMPLIDO WHERE ternro = " & nroter & _
            '" AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            'rs_justif!thnro
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
                " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
                rs_justif!thnro
            
            
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            wf_hcp2.Open StrSql, objConn

            ' Si no hay uno que cumpla con las condiciones lo creo
            If wf_hcp2.EOF Then
                'wf_hcp2.AddNew
                'Call ValidarTipoDeHora(rs_justif!thnro, nro_turno, tipo_hora)
                
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                rs_justif!maxhoras & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & CSng("0" & wf_hcp2!horcant) + rs_justif!maxhoras & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
                
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            wf_hcp2.Close
            
            'DELETE wf-justif.
            'rs_justif.Delete
            StrSql = "DELETE FROM " & TTempWFJustif & " where nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
            'Next.
            GoTo NextProgress

        End If
        
   If rs!horcant < rs_justif!maxhoras Then
'     FIND FIRST wf-hcp2 WHERE wf-hcp2.ternro    = nroter AND
'                              wf-hcp2.horfecrep = p-fecha AND
'                              wf-hcp2.thnro     = wf-justif.thnro EXCLUSIVE-LOCK NO-ERROR.
            StrSql = "SELECT * FROM gti_horcumplido WHERE ternro = " & nroter & _
            " AND horfecrep = " & ConvFecha(p_fecha) & " AND thnro = " & _
            rs_justif!thnro
        
            If wf_hcp2.State = adStateOpen Then wf_hcp2.Close
            'wf_hcp2.Open StrSql, objConn, adOpenKeyset, adLockBatchOptimistic
            wf_hcp2.Open StrSql, objConn
                        
            ' Si no hay uno que cumpla con las condiciones lo creo
            If wf_hcp2.EOF Then
                'wf_hcp2.AddNew
                'Call ValidarTipoDeHora(rs_justif!thnro, nro_turno, tipo_hora)
                
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horhasta," & _
                "horhoradesde,horhorahasta,hormanual,horvalido,ternro,thnro,empleg," & _
                "horfecrep,horestado) values (" & _
                rs!horcant & "," & _
                ConvFecha(p_fecha) & "," & ConvFecha(p_fecha) & ",'" & _
                rs_justif!horadesde & "','" & rs_justif!horahasta & "',0,-1," & _
                empleado.Ternro & "," & rs_justif!thnro & "," & empleado.Legajo & _
                "," & ConvFecha(p_fecha) & ",'')"
            Else
                StrSql = "UPDATE gti_horcumplido set horcant = " & CSng("0" & wf_hcp2!horcant) + rs!horcant & _
                ",hordesde = " & ConvFecha(p_fecha) & ", horhasta = " & _
                ConvFecha(p_fecha) & ",horhoradesde='" & rs_justif!horadesde & _
                "',horhorahasta='" & rs_justif!horahasta & _
                "',hormanual=0,horvalido=-1,ternro=" & empleado.Ternro & _
                ",thnro=" & rs_justif!thnro & ",empleg = " & empleado.Legajo & _
                ",horfecrep = " & ConvFecha(p_fecha) & ",horestado='' where hornro=" & _
                wf_hcp2!hornro
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'wf_hcp2!horcant = CSng("0" & wf_hcp2!horcant) + rs!horcant
            'wf_hcp2!hordesde = Date
            'wf_hcp2!horhasta = Date
            'wf_hcp2!horhoradesde = rs_justif!horadesde
            'wf_hcp2!horhorahasta = rs_justif!horahasta
            'wf_hcp2!hormanual = CInt(False)
            'wf_hcp2!horvalido = CInt(True)
            'wf_hcp2!Ternro = empleado.Ternro
            'wf_hcp2!thnro = rs_justif!thnro
            'wf_hcp2!Empleg = empleado.Legajo
            'wf_hcp2!horfecrep = Date
            'wf_hcp2!horestado = ""
        
            'Grabo y cierro el recordset
            'wf_hcp2.UpdateBatch
            wf_hcp2.Close
            
            'rs_justif!maxhoras = CSng("0" & rs_justif!maxhoras) - rs!horcant
            
            StrSql = "UPDATE " & TTempWFJustif & " SET maxhoras = " & rs_justif!maxhoras - rs!horcant & _
            " WHERE nro = " & rs_justif!nro
            objConn.Execute StrSql, , adExecuteNoRecords
                       
                       
            'StrSql = "SELECT * FROM gti_horcumplido WHERE (hornro = " & rs!hornro & ")"
        

            'If wf_hcp.State = adStateOpen Then
            '        wf_hcp.UpdateBatch
            '        wf_hcp.Close
            'End If
            'wf_hcp.Open StrSql, objConn, adOpenDynamic, adLockOptimistic
            
            'Borro y cierro el recordset
           ' wf_hcp.Delete
            'wf_hcp.UpdateBatch
            
            StrSql = "DELETE FROM gti_horcumplido WHERE hornro = " & rs!hornro
            objConn.Execute StrSql, , adExecuteNoRecords

            'wf_hcp.Close
                
            'Next.
            GoTo NextProgress

        End If
               
               
        
NextProgress:
        rs_justif.MoveNext
    Loop
    
    rs.MoveNext

Loop


CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    If rs_justif.State = adStateOpen Then
        rs_justif.UpdateBatch
        rs_justif.Close
    End If
    Set rs_justif = Nothing
    
    
StrSql = "DROP TABLE " & TTempWFJustif
objConn.Execute StrSql, adExecuteNoRecords
    
    
    
End Sub



Private Sub Calcula_Hora(cant As Single, Fecha As Date, vhora As String, vfecha As Date)

Dim cantminutos As Integer
'/* cantidad de minutos en una hora */
Dim Hora As Integer
Dim choras As Integer

Hora = 60
choras = 0
cantminutos = cant * Hora

Do While cantminutos >= Hora
    choras = choras + 1
    cantminutos = cantminutos - 60
Loop

vhora = Format(choras, "00") + Format(cantminutos, "00")
vfecha = IIf(choras > 24, Fecha + 1, Fecha)

End Sub


Public Sub Politica80v1(p_ternro As Long, Fecha As Date, fecha_desde As Date, hora_desde As String, fecha_hasta As Date, hora_hasta As String, subn As Integer)

' Registracion a evaluar
' subn = 2 repun  - Proceso las puntas, borro el resto.
' subn = 4 repes  - Proceso las puntas segun E o S, mantengo el resto.

Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim p_horario As Boolean
Dim salir As Boolean
Dim temp As String
Dim cantidad As Integer

salir = False

' Busco en el Detalle del parte de autorizacion de Horarios
' una entrada para p_ternro
StrSql = "SELECT * FROM gti_regdet WHERE ternro= " & p_ternro & _
" and gprdfecdesde <= " & ConvFecha(Fecha) & " and gprdfechasta >= " & ConvFecha(Fecha) & _
" ORDER BY ternro ASC, gprdfecdesde ASC, gprdfechasta ASC"

' Abro sin bloqueo
rs.Open StrSql, objConn
If Not rs.EOF Then
    
    ' Registracion - entrada

    StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,entrada,evenro)" & _
    " VALUES(" & p_ternro & "," & ConvFecha(rs!gprdfecdesde) & ",'" & _
    rs!grpdhoradesde & "',-1,2)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Registracion - salida
    StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,entrada,evenro)" & _
    " VALUES(" & p_ternro & "," & ConvFecha(rs!gprdfechasta) & ",'" & _
    rs!grpdhorahasta & "',0,2)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    p_horario = True
End If
rs.Close

'  /* Generar el wf con cada registracion dentro de la ventana */

If p_horario Then

    StrSql = "SELECT * FROM gti_registracion " & _
    " WHERE ternro= " & p_ternro & _
    " AND (regestado = 'I')" & _
    " AND ( (regfecha > " & ConvFecha(fecha_desde) & _
    ") OR (regfecha = " & ConvFecha(fecha_desde) & _
    " AND reghora >= '" & hora_desde & "') )" & _
    " AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
    ") OR (regfecha = " & ConvFecha(fecha_hasta) & _
    " AND reghora <= '" & hora_hasta & "') )"
    StrSql = StrSql & " ORDER BY ternro ASC, regfecha ASC, reghora ASC"

    OpenRecordset StrSql, rs
    
    Do While Not rs.EOF
        StrSql = "UPDATE gti_registracion SET regestado = 'S' " & _
        " WHERE regnro = " & rs!Regnro
        rs.MoveNext
    Loop
    rs.Close
Else

    'Proceso las puntas
    StrSql = "SELECT * FROM gti_registracion " & _
    " WHERE ternro= " & p_ternro & _
    " AND ((regestado = 'I') OR (regestado = 'S'))" & _
    " AND ( (regfecha > " & ConvFecha(fecha_desde) & _
    ") OR (regfecha = " & ConvFecha(fecha_desde) & _
    " AND reghora >= '" & hora_desde & "') )" & _
    " AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
    ") OR (regfecha = " & ConvFecha(fecha_hasta) & _
    " AND reghora <= '" & hora_hasta & "') )"
    
    If subn = 4 Then
        StrSql = StrSql & " AND regentsal = 'E'"
    End If
    
    StrSql = StrSql & " ORDER BY ternro ASC, regfecha ASC, reghora ASC"
       
    ' Solo lectura con keyset (para poder contar)
    rs.Open StrSql, objConn, adOpenKeyset, adLockReadOnly
    
    cantidad = rs.RecordCount
    ' assign rec-ent = recid(gti_registracion).

    ' Si hay un registro disponible
    If Not rs.EOF Then
 
        rs.MoveFirst
    
        ' genero el wf con la Registracion
        StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
        " VALUES(" & rs!Ternro & "," & ConvFecha(rs!regfecha) & ",'" & _
        rs!reghora & "'," & rs!Regnro & _
        "," & IIf(rs!regentsal = "E", -1, 0) & ",2)"
        objConn.Execute StrSql, , adExecuteNoRecords
      
    End If
    
    
    'Proceso las puntas
    rs.Close

    StrSql = "SELECT * FROM gti_registracion " & _
    " WHERE ternro= " & p_ternro & _
    " AND ((regestado = 'I') OR (regestado = 'S'))" & _
    " AND ( (regfecha > " & ConvFecha(fecha_desde) & _
    ") OR (regfecha = " & ConvFecha(fecha_desde) & _
    " AND reghora >= '" & hora_desde & "') )" & _
    " AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
    ") OR (regfecha = " & ConvFecha(fecha_hasta) & _
    " AND reghora <= '" & hora_hasta & "') )"
    
    If subn = 4 Then
        StrSql = StrSql & " AND regentsal = 'S'"
    End If
    
    StrSql = StrSql & " ORDER BY ternro ASC, regfecha ASC, reghora ASC"
              
    ' Solo lectura
    rs.Open StrSql, objConn
    
    If Not rs.EOF Then
    
        rs.MoveLast
    
        ' genero el wf con la Registracion
        
        If cantidad = 1 Then
            StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
            " VALUES(" & rs!Ternro & "," & ConvFecha(fecha_hasta) & ",'" & _
            hora_hasta & "'," & rs!Regnro & _
            "," & IIf(rs!regentsal = "E", -1, 0) & ",2)"
    
        Else
            StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
            " VALUES(" & rs!Ternro & "," & ConvFecha(rs!regfecha) & ",'" & _
            rs!reghora & "'," & rs!Regnro & _
            "," & IIf(rs!regentsal = "E", -1, 0) & ",2)"
    
        End If
    
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If

End If


If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub

Public Sub Politica80v3(p_ternro As Long, Fecha As Date, fecha_desde As Date, hora_desde As String, fecha_hasta As Date, hora_hasta As String, subn As Integer)

' Registracion a evaluar

Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim p_horario As Boolean
Dim salir As Boolean
Dim temp As String
Dim cantidad As Integer

Dim encontrado As Boolean
Dim fecres As Date
Dim horres As String
Dim ok As Boolean

'Este es similar al retod pero
'cuando al terminar tiene registraciones impares
'amplía la ventana de búsqueda hasta encontrar la registracion de
'salida faltante.


salir = False

' Busco en el Detalle del parte de autorizacion de Horarios
' una entrada para p_ternro

StrSql = "SELECT * FROM gti_regdet WHERE ternro = " & p_ternro & _
" and gprdfecdesde <= " & ConvFecha(Fecha) & " and gprdfechasta >= " & ConvFecha(Fecha) & _
" ORDER BY ternro ASC, gprdfecdesde ASC, gprdfechasta ASC"
rs.Open StrSql, objConn

If Not rs.EOF Then
    rs.MoveFirst
' MUY IMPORTANTE
' Chequear en la BD que los campos corresp a horas
' sean de tipo Fecha o con formato de hora
' Cambiar ConvHoraBD segun corresponda.
    
    ' Registracion - entrada

    StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,entrada,evenro)" & _
    " VALUES(" & p_ternro & "," & ConvFecha(rs!gprdfecdesde) & ",'" & _
    rs!grpdhoradesde & "',1,2)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Registracion - salida
    StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,entrada,evenro)" & _
    " VALUES(" & p_ternro & "," & ConvFecha(rs!gprdfechasta) & ",'" & _
    rs!grpdhorahasta & "',0,2)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    p_horario = True
    
End If
rs.Close

'  /* Generar el wf con cada registracion dentro de la ventana */

StrSql = "SELECT gti_registracion.* " & _
"FROM gti_reloj INNER JOIN gti_registracion ON gti_reloj.relnro = gti_registracion.relnro" & _
" WHERE ternro= " & p_ternro & _
" AND ((regestado ='I') OR (regestado = 'S'))" & _
" AND regfecha >= " & ConvFecha(fecha_desde) & " AND regfecha <= " & _
ConvFecha(fecha_hasta) & _
" ORDER BY ternro ASC, regfecha ASC, reghora ASC"

rs.Open StrSql, objConn
If rs.EOF Then Exit Sub

Do While Not rs.EOF
'Revisar: fn de formateo de hora
    If rs!regfecha = fecha_desde Then
        ' Convierto para comparar
        If rs!reghora < hora_desde Then
        ' Proceso el proximo
        ' Con esto reemplazo al next
            salir = True
        End If
    End If
    If rs!regfecha = fecha_hasta And Not salir Then
        If rs!reghora > hora_hasta Then
        'Si la registracion queda fuera de la ventana termino de procesar
            Exit Do
        End If
    End If
 
    If Not salir Then
    
        If p_horario Then
    
            StrSql = "UPDATE gti_registracion SET regestado = 'S' WHERE regnro = " & rs!Regnro
            objConn.Execute StrSql, , adExecuteNoRecords
        
        Else
            
            StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
            " VALUES(" & rs!Ternro & "," & ConvFecha(rs!regfecha) & ",'" & _
            rs!reghora & "'," & rs!Regnro & _
            "," & IIf(rs!regentsal = "E", -1, 0) & ",2)"
            objConn.Execute StrSql, , adExecuteNoRecords
        
        End If
        
    Else
        salir = False
    
    End If
    rs.MoveNext
Loop

' Busco las registraciones del intervalo en cuestión

StrSql = "SELECT * FROM gti_registracion " & _
" WHERE ternro= " & p_ternro & _
" AND ((regestado = 'I') OR (regestado = 'S'))" & _
" AND ( (regfecha > " & ConvFecha(fecha_desde) & _
") OR (regfecha = " & ConvFecha(fecha_desde) & _
" AND reghora >= '" & hora_desde & "') )" & _
" AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
") OR (regfecha = " & ConvFecha(fecha_hasta) & _
" AND reghora <= '" & hora_hasta & "') )" & _
" ORDER BY regfecha asc, reghora asc"

OpenRecordset StrSql, rs
cantidad = rs.RecordCount
rs.MoveLast

'Si son impares, entonces me falta una.
'Tengo que encontrar una de salida para cerrar.
'La ultima encontrada tiene que ser de entrada
If ((cantidad Mod 2) <> 0) And rs!regentsal = "E" Then

    encontrado = False
       
    Do While Not encontrado
    'Armo una ventana de búsqueda de tres horas para encontrar
    'la registracion faltante
                
        fecha_desde = fecha_hasta
        hora_desde = hora_hasta
        objFechasHoras.SumoHoras fecha_desde, hora_desde, "0300", fecres, horres
        ok = objFechasHoras.ValidarHora(horres)
        If Not ok Then Exit Sub
        hora_hasta = horres
        fecha_hasta = fecres
       
        'Consulto con la nueva ventana
        StrSql = "SELECT * FROM gti_registracion " & _
        " WHERE ternro= " & p_ternro & _
        " AND ((regestado = 'I') OR (regestado = 'S'))" & _
        " AND ( (regfecha > " & ConvFecha(fecha_desde) & _
        ") OR (regfecha = " & ConvFecha(fecha_desde) & _
        " AND reghora >= '" & hora_desde & "') )" & _
        " AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
        ") OR (regfecha = " & ConvFecha(fecha_hasta) & _
        " AND reghora <= '" & hora_hasta & "') )"
        
        OpenRecordset StrSql, rs
        
        'Si encuentro registraciones en la ventana
        If Not rs.EOF Then
        
            If rs!regentsal = "S" Then
            'Encontre la registracion de salida que me faltaba
                StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
                " VALUES(" & rs!Ternro & "," & ConvFecha(rs!regfecha) & ",'" & _
                rs!reghora & "'," & rs!Regnro & _
                "," & IIf(rs!regentsal = "E", -1, 0) & ",2)"
                objConn.Execute StrSql, , adExecuteNoRecords
            
                encontrado = True
            Else
            'Encontre una de entrada. No hay solución
                Exit Sub
            End If
        ' FGZ 15/07/2003
        Else
            ' No tiene solucion
            Exit Sub
        End If
    Loop
End If

If rs.State = adStateOpen Then rs.Close

Set rs = Nothing

End Sub


Public Sub politica810(ByRef Usa_Conv As Boolean)
'  Descripci¢n: Politica que indica si se usa conversi¢n de horas
'               true  = Usa Conversi¢n
'               false = No Usa Conversi¢n

    Usa_Conv = True

End Sub


Public Sub Politica390()
'Descripci¢n: Devuelve TRUE si se paga horas extras el almuerzo por no tomarlo
'                      FALSE si no se paga.

StrSql = "INSERT INTO wf_gtitpa(tpanro, vallog) VALUES (1, 0)"
objConn.Execute StrSql, , adExecuteNoRecords

End Sub


Public Sub politica800(ByRef fecha_proc As Integer, subn As Integer)
'  Descripci¢n: Pol¡tica que indica que fecha se debe tomar del HC para
'               calcular el Acum. Diario.
'               1 - Fecha Desde del HC
'               2 - Fecha de Procesamiento del HC
'               3 - Fecha Hasta del HC

Select Case subn
Case 1
    ' Fecha de primer registracion
    fecha_proc = 1
    
Case 2
    ' Fecha de la registracion
    fecha_proc = 2

Case 2
    ' Fecha de la ultima registracion
    fecha_proc = 3

End Select

End Sub


Public Sub Politica90(Fecha As Date, nroter As Long, ventdesde As String, ventFdesde As Date, venthasta As String, ventFhasta As Date, nro_turno As Long, ByRef ok As Boolean, subn As Integer)

'Descripci¢n: Devuelve TRUE si el empleado tiene MENOS o IGUAL de 6 registraciones
'                      FALSE si el empleado tiene MAS de 6 registraciones.

Dim rs As New ADODB.Recordset
Dim StrSql As String
Dim cant As Integer
Dim contar As Boolean


cant = 0
'/*** cuento las registraciones dentro de la ventana ***/
StrSql = "SELECT * FROM gti_registracion WHERE ternro= " & nroter & _
" and regfecha >= " & ConvFecha(ventFdesde) & " and  regfecha <= " & ConvFecha(ventFhasta) & _
" ORDER BY ternro ASC, regfecha ASC, reghora ASC"

OpenRecordset StrSql, rs

Do While Not rs.EOF
  If rs!regfecha = ventFdesde Then
      If rs!reghora < ventdesde Then
        'Si es temprano no la cuento
        contar = False
      End If
  End If
  If rs!regfecha = ventFhasta Then
      If rs!reghora > venthasta Then
        'Si la registracion queda fuera de la ventana me voy
        Exit Do
      End If
  End If
  If contar Then
    cant = cant + 1
  Else
    contar = True
  End If
  rs.MoveNext
Loop
ok = cant <= 6

End Sub


Public Sub Politica464(p_fecha As Date, nro_tercero As Long, nro_turno As Long, Dias_trabajados As Integer, Dias_laborables As Integer, deb As Boolean, subn As Integer)

' Genera los francos trabajados

' fratr
Dim tothor As Single

Dim rs As New ADODB.Recordset
Dim StrSql As String


If Dias_trabajados > Dias_laborables Then

        StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 29" & _
        " AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
        rs.Open StrSql, objConn
       
        If rs.EOF Then
            'Entrada en la traza
            GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Franco Trabajado para el Turno ", Str(nro_turno)
            
            '/* Deberia salir de todo el procedimiento */
            GoTo CierroTodo
        End If
        
        tothor = Dias_trabajados - Dias_laborables
        
        'Call ValidarTipoDeHora(rs!thnro, nro_turno, tipo_hora)
        ' Entrada en horario cumplido
        StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
        "horvalido,ternro,thnro,turnro,empleg,horfecrep,horhoradesde,horhorahasta" & _
        ") VALUES (" & tothor & "," & _
        ConvFecha(DateAdd("d", -1, p_fecha)) & "," & ConvFecha(DateAdd("d", -1, p_fecha)) & "," & "0,-1," & _
        nro_tercero & "," & rs!thnro & "," & nro_turno & "," & empleado.Legajo & "," & _
        ConvFecha(DateAdd("d", -1, p_fecha)) & ",'0000','0000')"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        rs.Close

End If

CierroTodo:
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub


Public Sub Politica465(p_fecha As Date, nro_tercero As Long, nro_turno As Long, Dias_trabajados As Integer, Dias_laborables As Integer, deb As Boolean, subn As Integer)

' Genera los francos compensatorios

    Call Politica465(p_fecha, nro_tercero, nro_turno, Dias_trabajados, Dias_laborables, deb, subn)

End Sub


Public Sub Politica400(Fecha As Date, nroter As Long, ventdesde As String, ventFdesde As Date, venthasta As String, ventFhasta As Date, nro_turno As Long, deb As Boolean, subn As Integer)

' Descripci¢n: Pol¡tica de almuerzo variable.
' Lo genera segun las 2 registraciones del medio.

Dim regnro_ent As Long
Dim regnro_sal As Long
Dim rec As Integer
Dim thora As Integer
Dim hora1 As String
Dim hora2 As String
Dim fecha1 As Date
Dim fecha2 As Date
Dim canthoras As Single
Dim thoras As Integer
Dim tmin As Integer
Dim tdias As Integer
Dim tothor As Single

Dim rs As New ADODB.Recordset
Dim StrSql As String


hora1 = ""

StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 18" & _
" AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
rs.Open StrSql, objConn
       
If Not rs.EOF Then
    thora = rs!thnro
Else
    'Entrada en la traza
    GeneraTraza empleado.Ternro, p_fecha, "No esta configurado el Tipo de Hora Almuerzo para el Turno ", Str(nro_turno)
            
    '/* Deberia salir de todo el procedimiento */
    GoTo CierroTodo
End If

rs.Close
' Esto supongo que es una busqueda sobre gti_registracion (dentro de una ventana)
' Devuelve el recid en gti_registracion

'  /*** busco la primer registracion de la ventana ***/
'run gtibureg.p(nroter, true, ventfdesde, ventdesde, ventfhasta, venthasta,
'               output recent).
StrSql = "SELECT * FROM gti_registracion WHERE ternro= " & nroter & _
" and regfecha >= " & ConvFecha(ventFdesde) & " AND regfecha <= " & ConvFecha(ventFhasta) & _
" AND reghora >= '" & ventdesde & "' AND reghora <= '" & venthasta & _
"' ORDER BY ternro ASC, regfecha ASC, reghora ASC"
' Abro para lectura
rs.Open StrSql, objConn, adOpenKeyset

' Busco el primero en la ventana
rs.MoveFirst
regnro_ent = rs!Regnro

'  /*** busco la ultima registracion de la ventana ***/
'run gtibureg.p(nroter, false, ventfdesde, ventdesde, ventfhasta, venthasta,
'               output recsal).

' Busco el ultimo en la ventana
rs.MoveLast
regnro_sal = rs!Regnro

' Cierro pues ya no lo necesito
rs.Close

If regnro_ent = regnro_sal Then
    '/* una sola registracion en el embudo */
    GoTo CierroTodo
End If

StrSql = "SELECT * FROM gti_registracion WHERE ternro= " & nroter & _
" ORDER BY ternro ASC, regfecha ASC, reghora ASC"
' Abro para lectura/escritura
rs.Open StrSql, objConn, adOpenKeyset, adLockOptimistic

Do While Not rs.EOF
'    if not available per.gti_registracion or
'       recid(per.gti_registracion) = recsal
'      then leave.
    If (rs.EOF) Or (rs!Regnro = regnro_sal) Then
        Exit Do
    End If
    If hora1 = "" Then
        hora1 = rs!reghora
        fecha1 = rs!regfecha
    Else
        hora2 = rs!reghora
        fecha2 = rs!regfecha
    End If
    
    rs!regestado = "P"
    rs.Update
    rs.MoveNext
Loop

rs.Close
' RUN gtir2f01.p(fecha1, hora1, fecha2, hora2, output tdias, output thoras, output tmin).

'Call objFechasHoras.RestaEntreHoras(fecha1, hora1, fecha2, hora2, tdias, thoras, tmin)
Call objFechasHoras.RestaHs(fecha1, hora1, fecha2, hora2, tdias, thoras, tmin)
canthoras = thoras + (tmin / 60)

'Call ValidarTipoDeHora(thora, nro_turno, tipo_hora)
' Entrada en horario cumplido
StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
"horvalido,ternro,thnro,turnro,empleg,horfecrep,horhoradesde,horhorahasta" & _
") VALUES (" & canthoras & "," & _
ConvFecha(Fecha) & "," & ConvFecha(Fecha) & "," & "0,-1," & _
nroter & "," & thora & "," & nro_turno & "," & empleado.Legajo & "," & _
ConvFecha(Fecha) & ",'" & hora1 & "','" & hora2 & "')"
        
objConn.Execute StrSql, , adExecuteNoRecords


CierroTodo:
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub


Public Sub Politica460(nroter As Long, p_fecha As Date, nro_grupo As Integer, deb As Boolean, Dias_trabajados As Integer, Dias_laborables As Integer, subn As Integer)

'' Habilita las politicas de francos
'
'Dim Tiene_Turno     As Boolean
'Dim Nro_dia         As Integer
'Dim nro_turno       As Integer
'Dim Nro_subturno    As Integer
'Dim Nro_fpgo        As Integer
'Dim diatipo         As Integer   '/*Indica si es laborable, no laborable o feriado */
'Dim Tiene_Justif    As Boolean
'Dim nro_justif      As Integer
'Dim Justif_turno    As Boolean
'Dim turcomp         As Boolean
'Dim Trabaja         As Boolean
'Dim fecha_desde     As Date
'Dim hora_desde      As String
'Dim fecha_hasta     As Date
'Dim hora_hasta      As String
'Dim fecha_tol       As Date
'Dim hora_tol        As String
'Dim fecha_toldto    As Date
'Dim hora_toldto     As String
'Dim hora_frac       As String
'Dim Orden_dia       As Integer
'Dim no_trabaja_just As Boolean
'Dim nro_jus_ent     As Integer
'Dim nro_jus_sal     As Integer
'Dim Existe_Reg      As Boolean
'Dim Fecha_inicio    As Date '/* fecha de inicio del turno */
'Dim codjus1         As Integer
'Dim codjus2         As Integer
'Dim Dia_libre       As Boolean
'Dim esFeriado       As Boolean
'Dim p_asignacion As Boolean
''DEF STREAM salida.
'Dim descripcion As String
'Dim Fecha As Date
'Dim i As Integer
''Dim recent as recid.
'Dim hab_464 As Boolean
'Dim hab_465 As Boolean
'
'Dim rs As New ADODB.Recordset
'Dim StrSql As String
'
'
'Fecha = DateAdd("d", -1, p_fecha)
''  FIND FIRST per.gti_cabpolitica where (per.gti_cabpolitica.cabpolnivel = 464) and
''                                       (per.gti_cabpolitica.cabpolestado) NO-LOCK NO-ERROR.
''  FIND FIRST per.gti_alcanpolitica of per.gti_cabpolitica
''        WHERE (per.gti_alcanpolitica.alcpolestado) AND
''              ((per.gti_alcanpolitica.alcpolnivel = 3 AND per.gti_alcanpolitica.alcpolorigen = per.empleado.empleg) OR
''               (per.gti_alcanpolitica.alcpolnivel = 2 AND per.gti_alcanpolitica.alcpolorigen = nro-grupo) OR
''               (per.gti_alcanpolitica.alcpolnivel = 1)) USE-INDEX cab-est-nivel NO-LOCK NO-ERROR.
''   If AVAILABLE(per.gti_alcanpolitica) Then
'    hab_464 = True
''  FIND FIRST per.gti_cabpolitica where (per.gti_cabpolitica.cabpolnivel = 465) and
''                                       (per.gti_cabpolitica.cabpolestado) NO-LOCK NO-ERROR.
''  FIND FIRST per.gti_alcanpolitica of per.gti_cabpolitica
''        WHERE (per.gti_alcanpolitica.alcpolestado) AND
''              ((per.gti_alcanpolitica.alcpolnivel = 3 AND per.gti_alcanpolitica.alcpolorigen = per.empleado.empleg) OR
''               (per.gti_alcanpolitica.alcpolnivel = 2 AND per.gti_alcanpolitica.alcpolorigen = nro-grupo) OR
''               (per.gti_alcanpolitica.alcpolnivel = 1)) USE-INDEX cab-est-nivel NO-LOCK NO-ERROR.
''   IF AVAILABLE(per.gti_alcanpolitica)
'    hab_465 = True
'
'If (Not hab_464) And (Not hab_465) Then
'    Exit Sub
'End If
'
'For i = 1 To 7
''   RUN esferiad.p(fecha, nroter, false ,output esferiado)
'
''   RUN bustur00.p(fecha ,nroter, false, OUTPUT tiene-turno,
''                  OUTPUT nro-turno, OUTPUT tiene-justif, OUTPUT nro-justif,
''                  OUTPUT justif-turno, OUTPUT turcomp, OUTPUT nro-grupo,
''                  OUTPUT nro-fpgo, OUTPUT fecha-inicio ,OUTPUT p-asignacion).
'
''   RUN busdia00.p (INPUT  nro-turno,
''                   INPUT  fecha,
''                   INPUT  fecha-inicio,
''                   INPUT  nroter,
''                   INPUT  p-asignacion,
''                   OUTPUT trabaja,
''                   OUTPUT orden-dia,
''                   OUTPUT nro-dia,
''                   OUTPUT nro-subturno,
''                   OUTPUT dia-libre).
'
'    If (Not Dia_libre) And (Not esFeriado) Then
'        Dias_laborables = Dias_laborables + 1
'    End If
'
''   find first gti_turno where gti_turno.turnro = nro-turno no-lock no-error.
'    StrSql = "SELECT * FROM gti_turno WHERE turnro= " & nro_turno
'    rs.Open StrSql, objConn
'
'    'run gti\busreg00.p(true, fecha, empleado.ternro ,output recent)
'
'    If ((recent <> "?") Or ((recent = "?") And (Tiene_Justif) And (Not Dia_libre))) And (Not esFeriado) Then
'        Dias_trabajados = Dias_trabajados + 1
'    End If
'
'    Fecha = DateAdd("d", -1, Fecha)
'
'Next

End Sub


Public Sub Politica110v2(p_fecha As Date, nro_tercero As Long, nro_turno As Long, deb As Boolean, subn As Integer)

'Descripci¢n: Genera Horario Ausente para el empleado que le faltan Reg. Oblig.

' corresponde a 110auspa
Dim rs_turno As New ADODB.Recordset
Dim rs_dia As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim i As Integer
Dim tothor As Single
Dim tdias As Integer
Dim tmin As Integer
Dim hora_oblig As Integer
Dim hora_generada As Integer
Dim fec_ent As Date
Dim hor_ent As String
Dim msgTraza As String
Dim es_turno_libre As Boolean

Dim cod_ent  As Integer

Dim ent_fic  As Boolean

Dim fec_sal  As Date
Dim hor_sal  As String
Dim cod_sal  As Integer

Dim sal_fic  As Boolean

Dim hor_ent_aux  As String
Dim cod_ent_aux  As Integer
Dim fec_ent_aux  As Date

Dim hor_sal_aux  As String
Dim cod_sal_aux  As Integer
Dim fec_sal_aux  As Date
 
Dim hora_ausencia As Integer

Dim Cant_emb(15) As Integer


'Abro cursor
StrSql = "SELECT * FROM " & TTempWFEmbudo
rs.Open StrSql, objConn

Do While Not rs.EOF
'      IF CAN-FIND(FIRST wf-turno WHERE (((Convfecha(fecha) > Convfecha(rs!fecha_desde)) or
'                                         (Convfecha(fecha) = Convfecha(rs!fecha_desde) and
'                                          hora >= rs!hora_desde)) and
'                                        ((Convfecha(fecha) < Convfecha(rs!fecha_hasta)) or
'                                         (Convfecha(fecha) = Convfecha(rs!fecha_hasta) and
'                                          hora <= rs!hora_hasta))))
'
    StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE" & _
    "(((fecha > " & ConvFecha(rs!fecha_desde) & ") or " & _
    "(fecha = " & ConvFecha(rs!fecha_desde) & " and " & _
    "hora >= '" & rs!hora_desde & "')) and " & _
    "((fecha < " & ConvFecha(rs!fecha_hasta) & ") or " & _
    "(fecha  = " & ConvFecha(rs!fecha_hasta) & " and hora <= '" & rs!hora_hasta & "')))"
    
'    " fecha >= " & ConvFecha(rs!fecha_desde) & _
'    " AND hora >= '" & rs!hora_desde & "' AND fecha <= " & _
'    ConvFecha(rs!fecha_hasta) & " AND hora <= '" & rs!hora_hasta & _
'    "' ORDER BY fecha ASC, hora ASC"
    OpenRecordset StrSql, rs_turno
    'rs_turno.Open StrSql, objConn
    
    If Not rs_turno.EOF Then
        Cant_emb(rs!Codigo) = Cant_emb(rs!Codigo) + 1
        rs_turno.Close
    End If
    rs.MoveNext
Loop
    
'Cierro el wf del embudo
rs.Close

'Abro cursor
StrSql = "SELECT * FROM " & TTempWFDia
rs_dia.Open StrSql, objConn
rs_dia.MoveFirst

ent_fic = True
sal_fic = True

Do While Not rs_dia.EOF
    If rs_dia!entrada Then
        fec_ent_aux = rs_dia!Fecha
        hor_ent_aux = rs_dia!Hora
        cod_ent_aux = rs_dia!Codigo

        StrSql = "SELECT * FROM " & TTempWFEmbudo & " WHERE codigo = " & rs_dia!Codigo
        rs.Open StrSql, objConn
    
        StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE" & _
        "(((fecha > " & ConvFecha(rs!fecha_desde) & ") or " & _
        "(fecha = " & ConvFecha(rs!fecha_desde) & " and " & _
        "hora >= '" & rs!hora_desde & "')) and " & _
        "((fecha < " & ConvFecha(rs!fecha_hasta) & ") or " & _
        "(fecha  = " & ConvFecha(rs!fecha_hasta) & " and hora <= '" & rs!hora_hasta & "')))"

        OpenRecordset StrSql, rs_turno
    
        If rs_turno.EOF Then
            ent_fic = False
        End If
        
        rs.Close
        rs_turno.Close
        
    Else
        fec_sal_aux = rs_dia!Fecha
        hor_sal_aux = rs_dia!Hora
        cod_sal_aux = rs_dia!Codigo
    
        StrSql = "SELECT * FROM " & TTempWFEmbudo & " WHERE codigo = " & rs_dia!Codigo
        rs.Open StrSql, objConn
    
        StrSql = "SELECT * FROM " & TTempWFTurno & " WHERE" & _
        "(((fecha > " & ConvFecha(rs!fecha_desde) & ") or " & _
        "(fecha = " & ConvFecha(rs!fecha_desde) & " and " & _
        "hora >= '" & rs!hora_desde & "')) and " & _
        "((fecha < " & ConvFecha(rs!fecha_hasta) & ") or " & _
        "(fecha  = " & ConvFecha(rs!fecha_hasta) & " and hora <= '" & rs!hora_hasta & "')))"
        
        OpenRecordset StrSql, rs_turno
        
        ' Cierro por innecesario
        rs.Close
        
        If rs_turno.EOF Then
            sal_fic = False
        End If

        If (ent_fic = False Or sal_fic = False) Then
            fec_ent = fec_ent_aux
            hor_ent = hor_ent_aux
            cod_ent = cod_ent_aux
            fec_sal = fec_sal_aux
            hor_sal = hor_sal_aux
            cod_sal = cod_sal_aux
        End If
    
        If Not ((hor_ent <> "" And hor_sal <> "") And (cod_sal = cod_ent + 1)) Then
            GoTo NextProgress
        End If
    
        'RUN gtir2f01.p(fec-ent, hor-ent, wf-dia.fecha, wf-dia.hora, output tdias, output thoras, output tmin).
        Call objFechasHoras.RestaHs(fec_ent, hor_ent, rs_dia!Fecha, rs_dia!Hora, tdias, thoras, tmin)
        tothor = thoras + (tmin / 60)
        If tothor = 0 Then
            ' Next
            GoTo NextProgress
        End If
        
        'Si no Encuentro una registracion en el embudo de entrada y tampoco encuentro registracion
        'en el embudo de salida, entonces se debe generar Ausencia Parcial.
        
        'Caso contrario, debo generar una Falta de Registracion Obligatoria
        
        If (ent_fic = False And sal_fic = False) Then
            StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 19 " & _
            "AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
            rs.Open StrSql, objConn
        Else
            StrSql = "SELECT * FROM gti_config_tur_hor WHERE conhornro = 10 " & _
            "AND turnro = " & nro_turno & " ORDER BY conhornro ASC, turnro ASC"
            rs.Open StrSql, objConn
        End If
    
        msgTraza = "No esta configurado el Tipo de Hora Falta Registracion para el Turno:"
   
        If Not rs.EOF Then
            hora_ausencia = rs!thnro
        Else
            'Entrada en la traza
            If depurar Then
                GeneraTraza empleado.Ternro, p_fecha, msgTraza, Str(nro_turno)
            End If
            
            '/* Deberia salir de todo el procedimiento */
            GoTo CierroTodo
        End If

        'Call ValidarTipoDeHora(hora_ausencia, nro_turno, tipo_hora)
'       Creo un registro en la tabla de HC
        StrSql = "INSERT INTO gti_horcumplido (horcant,hordesde,horhasta,hormanual," & _
        "horvalido,ternro,thnro,turnro,empleg,horfecrep," & _
        "normnro,normnro2,horhoradesde,horhorahasta,horestado) VALUES (" & tothor & "," & _
        ConvFecha(fec_ent) & "," & ConvFecha(fec_sal) & "," & "0,-1," & _
        nro_tercero & "," & hora_ausencia & "," & nro_turno & "," & empleado.Legajo & "," & _
        ConvFecha(p_fecha) & ",3,3,'" & hor_ent & "','" & hor_sal & "','')"
        
        objConn.Execute StrSql, , adExecuteNoRecords
        
        ent_fic = True
        sal_fic = True
        hor_ent = ""
        hor_sal = ""
        
        rs_turno.Close
        rs.Close
        
    End If
    
NextProgress:
    rs_dia.MoveNext

Loop
   

'Cierro y libero todos los recordsets
CierroTodo:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    If rs_dia.State = adStateOpen Then rs_dia.Close
    Set rs_dia = Nothing
    If rs_turno.State = adStateOpen Then rs_turno.Close
    Set rs_turno = Nothing


End Sub


Public Sub Politica80v2(p_ternro As Long, Fecha As Date, fecha_desde As Date, hora_desde As String, fecha_hasta As Date, hora_hasta As String, subn As Integer)

' Registracion a evaluar
' subn = 5 - esm26  - Proceso las puntas, maximo 26 horas.

Dim rs As New ADODB.Recordset
Dim StrSql As String

Dim p_horario As Boolean
Dim salir As Boolean
Dim temp As String
Dim cantidad As Integer

Dim tt_crpnnro As Integer
Dim tt_fechaproc As Date
Dim tt_hornro As Integer
Dim tt_regentsal As String
Dim tt_regestado As String
Dim tt_regfecha As Date
Dim tt_reghora As String
Dim tt_regmanual As Integer
Dim tt_regnro As Long
Dim tt_relnro As Integer
Dim tt_ternro As Integer

Dim tope As String

Dim fecha_aux As Date
Dim hora_aux  As String
Dim Hora As String
Dim fecres As Date
Dim horres As String
Dim ok As Boolean
Dim rec_ent As Long
Dim rec_sal As Long


salir = False
p_horario = False

' Busco en el Detalle del parte de autorizacion de Horarios
' una entrada para p_ternro

'FIND FIRST gti_regdet where PER.gti_regdet.ternro = p-ternro AND
'                            PER.gti_regdet.gprdfecdesde <= fecha AND
'                            fecha <= PER.gti_regdet.gprdfechasta no-lock no-error.
StrSql = "SELECT * FROM gti_regdet WHERE ternro= " & p_ternro & _
" and gprdfecdesde <= " & ConvFecha(Fecha) & " and gprdfechasta >= " & ConvFecha(Fecha) & _
" ORDER BY ternro ASC, gprdfecdesde ASC, gprdfechasta ASC"
' Abro sin bloqueo
rs.Open StrSql, objConn
If Not rs.EOF Then
    
    ' Registracion - entrada

    StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,entrada,evenro)" & _
    " VALUES(" & p_ternro & "," & ConvFecha(rs!gprdfecdesde) & ",'" & _
    rs!grpdhoradesde & "',-1,2)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Registracion - salida
    StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,entrada,evenro)" & _
    " VALUES(" & p_ternro & "," & ConvFecha(rs!gprdfechasta) & ",'" & _
    rs!grpdhorahasta & "',0,2)"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    p_horario = True
End If
rs.Close

'  /* Generar el wf con cada registracion dentro de la ventana */


If p_horario Then
    StrSql = "SELECT * FROM gti_registracion " & _
    " WHERE ternro= " & p_ternro & _
    " AND (regestado = 'I')" & _
    " AND ( (regfecha > " & ConvFecha(fecha_desde) & _
    ") OR (regfecha = " & ConvFecha(fecha_desde) & _
    " AND reghora >= '" & hora_desde & "') )" & _
    " AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
    ") OR (regfecha = " & ConvFecha(fecha_hasta) & _
    " AND reghora <= '" & hora_hasta & "') )"
    StrSql = StrSql & " ORDER BY ternro ASC, regfecha ASC, reghora ASC"

    OpenRecordset StrSql, rs
    
    Do While Not rs.EOF
        StrSql = "UPDATE gti_registracion SET regestado = 'S' " & _
        " WHERE regnro = " & rs!Regnro
        rs.MoveNext
    Loop
    rs.Close
    
Else

    'Proceso las puntas
    StrSql = "SELECT * FROM gti_registracion " & _
    " WHERE ternro= " & p_ternro & _
    " AND ((regestado = 'I') OR (regestado = 'S'))" & _
    " AND ( (regfecha > " & ConvFecha(fecha_desde) & _
    ") OR (regfecha = " & ConvFecha(fecha_desde) & _
    " AND reghora >= '" & hora_desde & "') )" & _
    " AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
    ") OR (regfecha = " & ConvFecha(fecha_hasta) & _
    " AND reghora <= '" & hora_hasta & "') )"
    StrSql = StrSql & " AND regentsal = 'E'"
    StrSql = StrSql & " ORDER BY ternro ASC, regfecha ASC, reghora ASC"
    
    ' Solo lectura
    OpenRecordset StrSql, rs
    
    cantidad = rs.RecordCount

    ' Si hay un registro disponible
    If Not rs.EOF Then
 
        rec_ent = rs!Regnro
        rs.MoveFirst
    
        ' genero el wf con la Registracion
        StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
        " VALUES(" & rs!Ternro & "," & ConvFecha(rs!regfecha) & ",'" & _
        rs!reghora & "'," & rs!Regnro & ",-1,2)"
        objConn.Execute StrSql, , adExecuteNoRecords
      
    End If
    
    rs.Close
    StrSql = "SELECT * FROM gti_registracion " & _
    " WHERE ternro= " & p_ternro & _
    " AND ((regestado = 'I') OR (regestado = 'S'))" & _
    " AND ( (regfecha > " & ConvFecha(fecha_desde) & _
    ") OR (regfecha = " & ConvFecha(fecha_desde) & _
    " AND reghora >= '" & hora_desde & "') )" & _
    " AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
    ") OR (regfecha = " & ConvFecha(fecha_hasta) & _
    " AND reghora <= '" & hora_hasta & "') )"
    StrSql = StrSql & " AND regentsal = 'S'"
    StrSql = StrSql & " ORDER BY ternro ASC, regfecha ASC, reghora ASC"
              
    ' Solo lectura
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
    
        rs.MoveLast
    
        ' genero el wf con la Registracion
        tt_crpnnro = rs!crpnnro
        If Not IsNull(rs!fechaproc) Then
            tt_fechaproc = rs!fechaproc
        End If
        If Not IsNull(rs!hornro) Then
            tt_hornro = rs!hornro
        End If
        tt_regentsal = rs!regentsal
        tt_regestado = rs!regestado
        tt_regfecha = rs!regfecha
        tt_reghora = rs!reghora
        tt_regmanual = rs!regmanual
        tt_regnro = rs!Regnro
        tt_relnro = rs!relnro
        tt_ternro = rs!Ternro
        
        rs.Close
        
        'rec_sal = recid(gti_registracion)
        
        StrSql = "SELECT * FROM gti_registracion " & _
        " WHERE ternro= " & p_ternro & _
        " AND ((regestado = 'I') OR (regestado = 'S'))" & _
        " AND ( (regfecha > " & ConvFecha(fecha_desde) & _
        ") OR (regfecha = " & ConvFecha(fecha_desde) & _
        " AND reghora >= '" & hora_desde & "') )" & _
        " AND ( (regfecha < " & ConvFecha(fecha_hasta) & _
        ") OR (regfecha = " & ConvFecha(fecha_hasta) & _
        " AND reghora <= '" & hora_hasta & "') )"
        StrSql = StrSql & " AND regentsal = 'E'"
        StrSql = StrSql & " ORDER BY ternro ASC, regfecha ASC, reghora ASC"
    
        ' Solo lectura
        OpenRecordset StrSql, rs
        
        If Not rs.EOF Then
            Call buscar_max26(p_ternro, rec_ent, rs!regfecha, rs!reghora, fecha_desde)
            'return
            GoTo CierroTodo
        Else
        
            If cantidad = 1 Then
                StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,entrada,evenro)" & _
                " VALUES(" & tt_ternro & "," & ConvFecha(tt_regfecha) & ",'" & _
                tt_reghora & "',0,2)"
    
            Else
                StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
                " VALUES(" & rs!Ternro & "," & ConvFecha(rs!regfecha) & ",'" & _
                rs!reghora & "'," & rs!Regnro & _
                ",0,2)"
    
            End If
    
            objConn.Execute StrSql, , adExecuteNoRecords
                    
        End If
    Else
        Call buscar_max26(p_ternro, rec_ent, rs!regfecha, rs!reghora, fecha_desde)
        
    End If

End If

CierroTodo:
If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub


Private Sub buscar_max26(p_ternro As Long, rec_ent As Long, fecres As Date, horres As String, fecha_desde As Date)
Dim fecha_aux As Date
Dim hora_aux  As String

Dim rs As New ADODB.Recordset
Dim StrSql As String

fecha_aux = fecha_hasta
hora_aux = hora_hasta
Dim Hora As String
Dim tope

           
'          run gtis1h01.p(per.gti_registracion.regfecha, hora, tope,
'                         output fecres, output horres).
'          /*** realiza la suma de 18 hs. ***/
    'Call objFechasHoras.SumoHoras(fecres, horres, "2600")
'          run pasarHr.p(horres, output horres, output ok).

    If Not objFechasHoras.ValidarHora(horres) Then
        Exit Sub
    End If

    fecha_hasta = fecres
    hora_hasta = horres

    StrSql = "SELECT * FROM gti_registracion " & _
    " WHERE ternro= " & p_ternro & _
    " AND ((regestado is null) OR (regestado = 'S'))" & _
    " AND regfecha >= " & ConvFecha(fecha_desde) & " AND regfecha <= " & _
    ConvFecha(fecha_hasta) & " AND regentsal = 'S'" & _
    " ORDER BY ternro ASC, regfecha ASC, reghora ASC"
              
    ' Solo lectura
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.MoveLast
        If rec_ent = rs!Regnro Then
                StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
                " VALUES(" & rs!Ternro & "," & ConvFecha(fecha_hasta) & ",'" & _
                hora_hasta & "'," & rs!Regnro & _
                ",0,2)"
        Else
                StrSql = "INSERT INTO " & TTempWFTurno & " (ternro,fecha,hora,regnro,entrada,evenro)" & _
                " VALUES(" & rs!Ternro & "," & ConvFecha(rs!regfecha) & ",'" & _
                rs!reghora & "'," & rs!Regnro & _
                ",0,2)"
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

End Sub

'Private Sub Politica900(nroProc As Integer)
'Dim desde As Date
'dim As Date
'Dim horas As Single
'
'    horas = 0
'
'    strsql = "SELECT * FROM gti_procacum WHERE gti_procacum.gpanro =" & nroProc
'    openrecordset strsql, objrs
'    If objrs.EOF Then Exit Sub
'
'    strsql = "SELECT * FROM gti_cab WHERE ternro =" & empleado.ternro & " AND gpanro = " & nroProc
'
'  'Depura Novedades de GTI
'  strsql = "DELETE FROM acu_nov WHERE (ternro = " & empleado.ternro & ") AND (gpanro =" & nroProc & ")"
'  objconn.execute strsql, , adexecutenorecords
'
'  find first empleado where empleado.ternro = nroter no-lock no-error.
'
'  assign desde = gti_procacum.gpadesde
'         hasta = gti_procacum.gpahasta.
'
'  for each per.gti_det where (gti_det.cgtinro = gti_cab.cgtinro) no-lock,
'      each per.tiph_con where (per.tiph_con.thnro = per.gti_det.thnro) and
'                              (per.tiph_con.folinro = empleado.folinro or
'                               per.tiph_con.folinro = ?) no-lock,
'      each per.tipopar where per.tipopar.tpanro = per.tiph_con.tpanro no-lock
'        break by per.tiph_con.concnro
'              By per.tiph_con.tpanro:
'
'        assign horas = horas + if tiph_con.thsuma
'                                  then PER.gti_det.dgticant
'                                  else (- PER.gti_det.dgticant).
'
'        if (last-of(per.tiph_con.concnro) or
'            last-of(per.tiph_con.tpanro)) and
'           (horas <> 0)
'        then
'          Do:
'           {findexlk.i &tabla = acu_nov
'                    &condicion = " (per.acu_nov.ternro = nroter) and
'                                   (per.acu_nov.concnro = per.tiph_con.concnro) and
'                                   (per.acu_nov.tpanro = per.tipopar.tpanro) and
'                                   (per.acu_nov.gpanro = nroproc) "}.
'
'            if not avail(per.acu_nov)
'            then
'              Do:
'                create per.acu_nov.
'                assign per.acu_nov.ternro = nroter
'                       per.acu_nov.concnro = per.tiph_con.concnro
'                       per.acu_nov.tpanro = per.tipopar.tpanro
'                       per.acu_nov.gpanro = nroproc.
'              end.
'            assign per.acu_nov.acnovvalor = horas
'                   horas = 0
'
'         end.
'
'  end. /* for each */
'
'
'End Sub

Private Sub Politica820(Fecha As Date, Ternro As Long, nro_turno As Long)
Dim saldo_actual As Single
Dim monto As Single
Dim objrsCC As New ADODB.Recordset


    monto = 0
    
    StrSql = "SELECT * FROM gti_tipo_ctacte "
    StrSql = StrSql & " INNER JOIN gti_config_ctacte ON gti_config_ctacte.tipccnro = gti_tipo_ctacte.tipccnro "
    StrSql = StrSql & " WHERE (gti_tipo_ctacte.tipccdesde >= " & ConvFecha(Fecha) & " or gti_tipo_ctacte.tipccdesde is null) and"
    StrSql = StrSql & " (gti_tipo_ctacte.tipcchasta <= " & ConvFecha(Fecha) & " or gti_tipo_ctacte.tipcchasta is null) and (gti_config_ctacte.turnro = " & nro_turno & ")"
    
    OpenRecordset StrSql, objrsCC
    Do While Not objrsCC.EOF
        StrSql = "SELECT * FROM gti_acumdiario WHERE (adfecha = " & ConvFecha(Fecha) & ") AND (ternro = " & Ternro & ") AND (thnro = " & objrsCC!thnro & ")"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
             If objrsCC!ccsigno Then
                monto = objRs!adcanthoras * objrsCC!ccpje / 100
             Else
                monto = ((objRs!adcanthoras * objrsCC!ccpje / 100) * (-1))
             End If
        Else
            'Si no está el AD, puede que haya sido borrado, entonces se busca el
            ' detalle para ese acum. diario y si está se borra y se actualiza el saldo */
            Politica820BorraDet Ternro, objrsCC!tipccnro, Fecha, objrsCC!thnro
            
            GoTo continuar
        End If
        StrSql = "SELECT * FROM gti_ctacte_saldo WHERE (ternro =" & Ternro & ") AND (tipccnro = " & objrsCC!tipccnro & ")"
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            StrSql = "INSERT INTO gti_ctacte_saldo(ternro,tipccnro,saldoccvalor) VALUES ("
            StrSql = StrSql & Ternro & "," & objrsCC!tipccnro & "," & ValorNulo & ")"
            
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        saldo_actual = IIf(IsNull(objRs!saldoccvalor), 0, objRs!saldoccvalor)
        If (saldo_actual + monto) < 0 Then
            If objrsCC!tipccneg Then
                Politica820ActualizaDet Ternro, objrsCC!tipccnro, monto, Fecha, objrsCC!thnro
            Else
              ' El Saldo del Cta. es < 0 y no se permiten saldos (-) */
              'message "Saldo (-) No Permitido, Cta. Cte. : " + string(per.gti_tipo_ctacte.tipccnro, ">>>>9").
            End If
        Else
            Politica820ActualizaDet Ternro, objrsCC!tipccnro, monto, Fecha, objrsCC!thnro
        End If
        
continuar:
        objrsCC.MoveNext
    Loop
End Sub

Private Sub Politica820ActualizaDet(Ternro As Long, nrocc As Long, monto As Single, Fecha As Date, thnro As Integer)
Dim objRs As New ADODB.Recordset
Dim objrsBuf As New ADODB.Recordset

    StrSql = "SELECT * FROM gti_ctacte_det WHERE (ternro = " & Ternro & ") AND (tipccnro = " & nrocc & ") AND " & _
             "(detccfecha = " & ConvFecha(Fecha) & ") AND (thnro = " & thnro & ")"
    
    OpenRecordset StrSql, objRs, adLockOptimistic
    
    If objRs.EOF Then
        StrSql = "INSERT INTO gri_ctacte_det(ternro,tipccnro,detccfecha,thnro) VALUES (" & _
                 Ternro & "," & nrocc & "," & ConvFecha(Fecha) & "," & thnro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    StrSql = "SELECT * FROM gti_ctacte_det WHERE (ternro = " & Ternro & ") AND (tipccnro = " & nrocc & ") AND " & _
             "(detccfecha = " & ConvFecha(Fecha) & ") AND (thnro = " & thnro & ")"
    
    OpenRecordset StrSql, objRs
        
    StrSql = "UPDATE gti_ctacte_det SET detccvalor = " & monto & " WHERE (ternro = " & Ternro & ") AND (tipccnro = " & nrocc & ") AND " & _
             "(detccfecha = " & ConvFecha(Fecha) & ") AND (thnro = " & thnro & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Busca si estaba compensado con alguno
    If objRs!detcccomp Then
        StrSql = "SELECT * FROM gti_ctacte_det WHERE detccidcomp = " & objRs!detccnro
        OpenRecordset StrSql, objrsBuf
        If Not objrsBuf.EOF Then
            objrsBuf!detccidcomp = 0
            objRs!detcccomp = CInt(False)
        End If
    End If
    
    StrSql = "SELECT * FROM gti_ctacte_det WHERE detccnro= " & objRs!detccidcomp
    OpenRecordset StrSql, objrsBuf
    If Not objrsBuf.EOF Then
        objrsBuf!detcccomp = CInt(False)
        objRs!detccidcomp = 0
    End If
    
    StrSql = "SELECT * FROM gti_ctacte_det WHERE (detcccomp = " & CInt(False) & ") AND " & _
             "(detccvalor = ( " & objRs!detccvalor * -1 & ")) AND (detccfecha  <= " & ConvFecha(objRs!detccfecha) & " )"
    OpenRecordset StrSql, objrsBuf
    If Not objrsBuf.EOF Then
        objRs!detccidcomp = objrsBuf!detccnro
        objrsBuf!detcccomp = CInt(True)
    End If
    
    
End Sub

Private Sub Politica820BorraDet(Ternro As Long, nrocc As Long, Fecha As Date, thnro As Integer)
Dim monto As Single
Dim objRs As New ADODB.Recordset
Dim objrsBuf As New ADODB.Recordset
  
  
    StrSql = "SELECT * FROM gti_ctacte_det WHERE (ternro = " & Ternro & ") AND " & _
             "(tipccnro = " & nrocc & ") AND (detccfecha = " & ConvFecha(Fecha) & ") AND " & _
             "(thnro =" & thnro & ")"
    OpenRecordset StrSql, objRs
    If objRs.EOF Then Exit Sub
    monto = IIf(IsNull(objRs!detccvalor), 0, objRs!detccvalor)
    
    ' Busca si estaba compensado con alguno
    StrSql = "UPDATE gti_ctacte_det SET detccidcomp = 0 WHERE detccidcomp = " & objRs!detccnro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Busca si compensaba a alguien
    StrSql = "UPDATE gti_ctacte_det SET detcccomp = " & CInt(False) & " WHERE detccnro = " & objRs!detccidcomp
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Actualiza el Saldo
    StrSql = "UPDATE gti_ctacte_saldo SET saldoccvalor = " & objRs!saldoccvalor - monto & " WHERE (ternro =" & Ternro & ") AND (tipccnro = " & nrocc & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
    
    ' Borra el Detalle
    StrSql = "DELETE FROM gti_ctacte_det WHERE (ternro = " & Ternro & ") AND " & _
             "(tipccnro = " & nrocc & ") AND (detccfecha = " & ConvFecha(Fecha) & ") AND " & _
             "(thnro =" & thnro & ")"
    objConn.Execute StrSql, , adExecuteNoRecords
      
End Sub

Private Sub Politica540(thnro As Integer, Ternro As Long, Fecha As Date, nro_turno As Long, ByRef tothor As Single)
' Controla la cuenta corriente de horas extras y de acuerdo al limite mensual hace la conversion
Dim th50 As Integer
Dim th100 As Integer
Dim objRs As New ADODB.Recordset
Dim objrsCC As New ADODB.Recordset

    StrSql = "SELECT gti_config_tur_hora.thnro FROM gti_config_hora " & _
            " INNER JOIN gti_config_tur_hora ON gti_config_hora.conhornro = gti_config_tur_hora.conhornro " & _
            " WHERE gti_config_hora.conhornro = 1 AND gti_config_tur_hora.turnro =" & nro_turno
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then th50 = objRs!thnro
    
    StrSql = "SELECT gti_config_tur_hora.thnro FROM gti_config_hora " & _
            " INNER JOIN gti_config_tur_hora ON gti_config_hora.conhornro = gti_config_tur_hora.conhornro " & _
            " WHERE gti_config_hora.conhornro = 1 AND gti_config_tur_hora.turnro =" & nro_turno
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then th100 = objRs!thnro
    
    If (thnro = th50) Or (thnro = th100) Then

        StrSql = " SELECT  * FROM gti_ctacte_saldo WHERE tipccnro = 1 AND ternro = " & Ternro
        OpenRecordset StrSql, objrsCC
        If Not objrsCC.EOF Then
            StrSql = " SELECT * FROM tiphora WHERE tiphora.thnro = " & thnro
            OpenRecordset StrSql, objRs
            If objrsCC!saldoccvalor = 30 Then
                'Call ValidarTipoDeHora(objRs!thexcmes, nro_turno, tipo_hora)
                StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horestado,horhasta,horhoradesde," & _
                         "horhorahasta,hormanual,horvalido,jusnro,jusnro2,normnro,normnro2,Ternro," & _
                         "thnro,turnro,regent,horrealent,regsal,horrealsal,Empleg,horfecrep) VALUES (" & _
                         tothor & "," & ConvFecha(Fecha) & ",''," & ConvFecha(Fecha) & "," & ValorNulo & "," & _
                         ValorNulo & "," & CInt(False) & "," & CInt(True) & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                         ValorNulo & "," & empleado.Ternro & "," & objRs!thexcmes & "," & nro_turno & "," & _
                         ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                         ValorNulo & "," & empleado.Legajo & "," & ConvFecha(p_fecha) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
                
                tothor = 0
                
            ElseIf objrsCC!saldoccvalor < 30 Then
            
                If (objrsCC!saldoccvalor + tothor) > 30 Then
                    'Call ValidarTipoDeHora(objRs!thexcmes, nro_turno, tipo_hora)
                    StrSql = "INSERT INTO gti_horcumplido(horcant,hordesde,horestado,horhasta,horhoradesde," & _
                         "horhorahasta,hormanual,horvalido,jusnro,jusnro2,normnro,normnro2,Ternro," & _
                         "thnro,turnro,regent,horrealent,regsal,horrealsal,Empleg,horfecrep) VALUES (" & _
                         (objrsCC!saldoccvalor + tothor - 30) & "," & ConvFecha(Fecha) & ",''," & ConvFecha(Fecha) & "," & ValorNulo & "," & _
                         ValorNulo & "," & CInt(False) & "," & CInt(True) & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                         ValorNulo & "," & empleado.Ternro & "," & objRs!thexcmes & "," & nro_turno & "," & _
                         ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                         ValorNulo & "," & empleado.Legajo & "," & ConvFecha(p_fecha) & ")"
                         
                    objConn.Execute StrSql, , adExecuteNoRecords
                    tothor = 30
                End If
            End If
        End If
    End If

End Sub

Public Sub Politica900(NroProceso As Long, subn As Integer)
Dim objRs As New ADODB.Recordset
Dim Progreso As Single
Dim CEmpleadosAProc As Integer
Dim IncPorc As Single
    
    StrSql = " SELECT bprcfecdesde,bprcfechasta,gti_cab.gpanro,gti_cab.cgtinro,batch_proceso.bpronro,gti_procacum.gtprocnro,gti_cab.ternro FROM batch_proceso " & _
             " INNER JOIN batch_procacum ON batch_procacum.bpronro = batch_proceso.bpronro " & _
             " INNER JOIN gti_cab ON  gti_cab.gpanro = batch_procacum.gpanro " & _
             " INNER JOIN gti_procacum ON batch_procacum.gpanro = gti_procacum.gpanro " & _
             " INNER JOIN batch_empleado ON batch_empleado.ternro = gti_cab.ternro AND batch_empleado.bpronro = batch_proceso.bpronro" & _
             " WHERE batch_proceso.bpronro = " & NroProceso
             
    OpenRecordset StrSql, objRs
    
    ' Seteo el incremento de progreso
    CEmpleadosAProc = objRs.RecordCount
    IncPorc = (100 / CEmpleadosAProc)
    Progreso = 0
    
    Do While Not objRs.EOF
        
        Select Case subn
        Case 1:
            Flog.writeline "Novedades_GTI_Expo"
            Call Novedades_GTI_Expo(objRs!Ternro, objRs!gpanro)
        Case 2:
            Flog.writeline "Novedades_GTI_Expo"
            Call Novedades_GTI_Expo(objRs!Ternro, objRs!gpanro)
        Case 3:
            Flog.writeline "Novedades_GTI_MAZUL"
            Call Novedades_GTI_MAZUL(objRs!Ternro, objRs!gpanro)
        End Select
        
        ' Actualizo el progreso
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        
        objRs.MoveNext
    Loop

End Sub

Private Sub Novedades_GTI_Expo(nroter As Long, nroproc As Long)

Dim desde As Date
Dim hasta As Date
Dim horas As Single
Dim nrocab As Long
Dim objRs As New ADODB.Recordset
Dim objrsemp As New ADODB.Recordset
Dim objrsanov As New ADODB.Recordset

    StrSql = " SELECT * FROM gti_procacum WHERE gpanro = " & nroproc
    OpenRecordset StrSql, objRs
    
    If objRs.EOF Then
        Exit Sub
    Else
    
        StrSql = " SELECT * FROM gti_cab WHERE ternro = " & nroter & " and gpanro = " & nroproc
        OpenRecordset StrSql, objRs
        
        nrocab = objRs!cgtinro
        
        StrSql = " DELETE FROM gti_acunov WHERE ternro = " & nroter & " and gpanro = " & nroproc
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " SELECT * FROM empleado WHERE ternro = " & nroter
        OpenRecordset StrSql, objrsemp
        
        
' Hay que filtrar por la estructura del empleado, ya que la relacion entre
' los tipos de horas y los conceptos, puede ser por estructura

'AIB - 26/3/2003
'La relación no necesariamente es por estructura

'        StrSql = " SELECT gti_det.dgticant,tiph_con.thsuma,tiph_con.concnro,tiph_con.tpanro " & _
                 " FROM gti_det " & _
                 " INNER JOIN tiph_con ON tiph_con.thnro = gti_det.thnro " & _
                 " INNER JOIN estruc_actual ON estruc_actual.ternro = " & nroter & _
                 " AND estruc_actual.tenro = tiph_con.tenro AND estruc_actual.estrnro = tiph_con.estrnro " & _
                 " WHERE cgtinro = " & nrocab & _
                 " ORDER BY tiph_con.concnro "
        ' FGZ 04/08/2003
        ' cambio estruc_actual por his_estructura
        StrSql = " SELECT gti_det.dgticant,tiph_con.thsuma,tiph_con.concnro,tiph_con.tpanro " & _
            " FROM gti_det " & _
            " INNER JOIN tiph_con ON tiph_con.thnro = gti_det.thnro " & _
            " INNER JOIN his_estructura ON his_estructura.ternro = " & nroter & _
            " AND his_estructura.tenro = tiph_con.tenro AND his_estructura.estrnro = tiph_con.estrnro " & _
            " WHERE cgtinro = " & nrocab & " AND his_estructura.htethasta IS NULL " & _
            " UNION " & _
            " SELECT gti_det.dgticant,tiph_con.thsuma,tiph_con.concnro,tiph_con.tpanro " & _
            " FROM gti_det " & _
            " INNER JOIN tiph_con ON tiph_con.thnro = gti_det.thnro " & _
            " WHERE cgtinro = " & nrocab & _
            " AND tiph_con.tenro is null "
' Hay que filtrar por la estructura del empleado, ya que la relacion entre
' los tipos de horas y los conceptos, puede ser por estructura
        
        OpenRecordset StrSql, objRs
        
        Do While Not objRs.EOF
        
            If objRs!thsuma Then
        
                horas = horas + objRs!dgticant
            
            Else
                
                horas = horas + (-objRs!dgticant)
            
            End If
            
            'Informacion en Cantidad de Días
            
            If objRs!tpanro = 161 Then horas = Round(horas / 8, 2)
            
            
            ' 30 Días Trabajados como Máximo
            
            
            'Premio Reduccion al Ausentismo
            
            'If (objRs!conccod = "911") And (objRs!dgticant < 5) Then
                
            '    horas = objRs!dgticant * 20
                
            'Else
            
            '    horas = 100
                
            'End If
            
            StrSql = "SELECT * FROM gti_acunov WHERE gti_acunov.ternro = " & nroter & _
                     "AND gti_acunov.concnro = " & objRs!concnro & _
                     "AND gti_acunov.tpanro = " & objRs!tpanro & _
                     "AND gti_acunov.gpanro = " & nroproc
            
            OpenRecordset StrSql, objrsanov, adLockOptimistic
            
            If objrsanov.EOF Then
            
                StrSql = " INSERT INTO  gti_acunov " & _
                         " (ternro,concnro,tpanro,gpanro,acnovvalor) VALUES " & _
                         " ( " & nroter & "," & objRs!concnro & "," & objRs!tpanro & "," & nroproc & "," & horas & ")"
            
                objConn.Execute StrSql, , adExecuteNoRecords
            
            Else
            
               objrsanov.Update "acnovvalor", objrsanov!acnovvalor + horas
                        
            End If
        
            objRs.MoveNext
            
            horas = 0
            
        Loop
        
    
    End If
    
    
End Sub



Private Sub Novedades_GTI_Standar(nroter As Long, nroproc As Long)

Dim desde As Date
Dim hasta As Date
Dim horas As Single
Dim nrocab As Long
Dim objRs As New ADODB.Recordset
Dim objrsemp As New ADODB.Recordset
Dim objrsanov As New ADODB.Recordset

    StrSql = " SELECT * FROM gti_procacum WHERE gpanro = " & nroproc
    OpenRecordset StrSql, objRs
    
    If objRs.EOF Then
        Exit Sub
    Else
    
        StrSql = " SELECT * FROM gti_cab WHERE ternro = " & nroter & " and gpanro = " & nroproc
        OpenRecordset StrSql, objRs
        
        nrocab = objRs!cgtinro
        
        StrSql = " DELETE FROM gti_acunov WHERE ternro = " & nroter & " and gpanro = " & nroproc
        objConn.Execute StrSql, , adExecuteNoRecords
        
        StrSql = " SELECT * FROM empleado WHERE ternro = " & nroter
        OpenRecordset StrSql, objrsemp
        
        
' Hay que filtrar por la estructura del empleado, ya que la relacion entre
' los tipos de horas y los conceptos, puede ser por estructura

'AIB - 26/3/2003
'La relación no necesariamente es por estructura

        ' FGZ 04/08/2003
        ' cambio estruc_actual por his_estructura
        StrSql = " SELECT gti_det.dgticant,tiph_con.thsuma,tiph_con.concnro,tiph_con.tpanro " & _
            " FROM gti_det " & _
            " INNER JOIN tiph_con ON tiph_con.thnro = gti_det.thnro " & _
            " INNER JOIN his_estructura ON his_estructura.ternro = " & nroter & _
            " AND his_estructura.tenro = tiph_con.tenro AND his_estructura.estrnro = tiph_con.estrnro " & _
            " WHERE cgtinro = " & nrocab & " AND his_estructura.htethasta IS NULL " & _
            " UNION " & _
            " SELECT gti_det.dgticant,tiph_con.thsuma,tiph_con.concnro,tiph_con.tpanro " & _
            " FROM gti_det " & _
            " INNER JOIN tiph_con ON tiph_con.thnro = gti_det.thnro " & _
            " WHERE cgtinro = " & nrocab & _
            " AND tiph_con.tenro is null "
' Hay que filtrar por la estructura del empleado, ya que la relacion entre
' los tipos de horas y los conceptos, puede ser por estructura
        
        OpenRecordset StrSql, objRs
        
        Do While Not objRs.EOF
            If objRs!thsuma Then
                horas = horas + objRs!dgticant
            Else
                horas = horas + (-objRs!dgticant)
            End If
            
            'Informacion en Cantidad de Días
            If objRs!tpanro = 161 Then horas = Round(horas / 8, 2)
            
            StrSql = "SELECT * FROM gti_acunov WHERE gti_acunov.ternro = " & nroter & _
                     "AND gti_acunov.concnro = " & objRs!concnro & _
                     "AND gti_acunov.tpanro = " & objRs!tpanro & _
                     "AND gti_acunov.gpanro = " & nroproc
            OpenRecordset StrSql, objrsanov, adLockOptimistic
            
            If objrsanov.EOF Then
                StrSql = " INSERT INTO  gti_acunov " & _
                         " (ternro,concnro,tpanro,gpanro,acnovvalor) VALUES " & _
                         " ( " & nroter & "," & objRs!concnro & "," & objRs!tpanro & "," & nroproc & "," & horas & ")"
            
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
               objrsanov.Update "acnovvalor", objrsanov!acnovvalor + horas
            End If
        
            objRs.MoveNext
            horas = 0
        Loop
    End If
End Sub


Private Sub Novedades_GTI_MAZUL(nroter As Long, nroproc As Long)

Dim desde As Date
Dim hasta As Date
Dim horas As Single
Dim nrocab As Long
Dim objRs As New ADODB.Recordset
Dim objrsemp As New ADODB.Recordset
Dim objrsanov As New ADODB.Recordset

    StrSql = " SELECT * FROM gti_procacum WHERE gpanro = " & nroproc
    Flog.writeline "SQL :" & StrSql
    OpenRecordset StrSql, objRs
    
    
    If objRs.EOF Then
        Flog.writeline "EOF en gti_procacum :" & Now
        Exit Sub
    Else
    
        StrSql = " SELECT * FROM gti_cab WHERE ternro = " & nroter & " and gpanro = " & nroproc
        OpenRecordset StrSql, objRs
        
        nrocab = objRs!cgtinro
        
        Flog.writeline "cabecera :" & objRs!cgtinro
        
        StrSql = " DELETE FROM gti_acunov WHERE ternro = " & nroter & " and gpanro = " & nroproc
        objConn.Execute StrSql, , adExecuteNoRecords
        
        Flog.writeline "Borró gti_acunov :" & StrSql
        
        StrSql = " SELECT * FROM empleado WHERE ternro = " & nroter
        OpenRecordset StrSql, objrsemp
        
        
' Hay que filtrar por la estructura del empleado, ya que la relacion entre
' los tipos de horas y los conceptos, puede ser por estructura

'AIB - 26/3/2003
'La relación no necesariamente es por estructura
        StrSql = " SELECT gti_det.dgticant,tiph_con.thsuma,tiph_con.concnro,tiph_con.tpanro,concepto.conccod, tiph_con.thnro" & _
        " FROM gti_det " & _
        " INNER JOIN tiph_con ON tiph_con.thnro = gti_det.thnro " & _
        " INNER JOIN his_estructura ON his_estructura.ternro = " & nroter & _
        " AND his_estructura.tenro = tiph_con.tenro AND his_estructura.estrnro = tiph_con.estrnro " & _
        " INNER JOIN concepto ON tiph_con.concnro = concepto.concnro" & _
        " WHERE cgtinro = " & nrocab & " AND his_estructura.htethasta IS NULL " & _
        " UNION" & _
        " (SELECT gti_det.dgticant,tiph_con.thsuma,tiph_con.concnro,tiph_con.tpanro,concepto.conccod, tiph_con.thnro " & _
        " FROM gti_det " & _
        " INNER JOIN tiph_con ON tiph_con.thnro = gti_det.thnro " & _
        " INNER JOIN concepto ON tiph_con.concnro = concepto.concnro" & _
        " WHERE cgtinro = " & nrocab & " AND tiph_con.tenro IS NULL)"

        ' cambio estruc_actual por his_estructura
'        StrSql = " SELECT gti_det.dgticant,tiph_con.thsuma,tiph_con.concnro,tiph_con.tpanro " & _
'            " FROM gti_det " & _
'            " INNER JOIN tiph_con ON tiph_con.thnro = gti_det.thnro " & _
'            " INNER JOIN his_estructura ON his_estructura.ternro = " & nroter & _
'            " AND his_estructura.tenro = tiph_con.tenro AND his_estructura.estrnro = tiph_con.estrnro " & _
'            " WHERE cgtinro = " & nrocab & " AND his_estructura.htethasta IS NULL " & _
'            " UNION " & _
'            " SELECT gti_det.dgticant,tiph_con.thsuma,tiph_con.concnro,tiph_con.tpanro " & _
'            " FROM gti_det " & _
'            " INNER JOIN tiph_con ON tiph_con.thnro = gti_det.thnro " & _
'            " WHERE cgtinro = " & nrocab & _
'            " AND tiph_con.tenro is null "
' Hay que filtrar por la estructura del empleado, ya que la relacion entre
' los tipos de horas y los conceptos, puede ser por estructura
        
        OpenRecordset StrSql, objRs
        
        Flog.writeline "SQL :" & StrSql
        
        If objRs.EOF Then
            Flog.writeline "EOF :" & StrSql
        End If
        
        Do While Not objRs.EOF
            
            horas = 0
            
            If objRs!thsuma Then
        
                horas = horas + objRs!dgticant
            
            Else
                
                horas = horas + (-objRs!dgticant)
            
            End If
            
            Flog.writeline "Por cada ternro ==> Concepto :" & objRs!conccod
            Flog.writeline "Por cada ternro ==> Thnro :" & objRs!thnro
            Flog.writeline "Por cada ternro ==> dgticant: " & objRs!dgticant
            
            
            If (objRs!conccod = "911") Then
             
                If (objRs!thnro = 14) Then
                    
                    If (objRs!dgticant < 5) Then
                    
                        horas = objRs!dgticant * 20
                    
                    Else
                
                        horas = 100
                    
                    End If
                
                Else
                    
                    If (objRs!dgticant < 1) Then
                        
                        horas = 50
                    
                    Else
                        
                        horas = 100
                    
                    End If
                    
                End If
            
            End If
            
            
            If (objRs!conccod = "910") Then
             
                If (objRs!thnro = 54) Then
                
                    
                    If (objRs!dgticant <= 10) Then
                    
                        horas = Round(objRs!dgticant, 1) * 10
                    
                    Else
                
                        horas = 100
                    
                    End If
                    
                End If
            
            End If
            
            If horas = 0 Then GoTo continuar
            
            StrSql = "SELECT * FROM gti_acunov WHERE gti_acunov.ternro = " & nroter & _
                     " AND gti_acunov.concnro = " & objRs!concnro & _
                     " AND gti_acunov.tpanro = " & objRs!tpanro & _
                     " AND gti_acunov.gpanro = " & nroproc
            
            Flog.writeline "SQL ==> :" & StrSql
            OpenRecordset StrSql, objrsanov, adLockOptimistic
            
            If objrsanov.EOF Then
                
                StrSql = " INSERT INTO  gti_acunov " & _
                         " (ternro,concnro,tpanro,gpanro,acnovvalor) VALUES " & _
                         " ( " & nroter & "," & objRs!concnro & "," & objRs!tpanro & "," & nroproc & "," & horas & ")"
            
                Flog.writeline "inserto en GTI_Acunov ==> :" & StrSql
                
                objConn.Execute StrSql, , adExecuteNoRecords
            
            Else
                
                If (objRs!conccod = 910 Or objRs!conccod = 911) And (objrsanov!acnovvalor + horas >= 100) Then
                    StrSql = " UPDATE gti_acunov " & _
                        " SET acnovvalor = " & 100 & _
                        " WHERE gti_acunov.ternro = " & nroter & _
                        " AND gti_acunov.concnro = " & objRs!concnro & _
                        " AND gti_acunov.tpanro = " & objRs!tpanro & _
                        " AND gti_acunov.gpanro = " & nroproc
                Else
                    StrSql = " UPDATE gti_acunov " & _
                        " SET acnovvalor = " & objrsanov!acnovvalor + horas & _
                        " WHERE gti_acunov.ternro = " & nroter & _
                        " AND gti_acunov.concnro = " & objRs!concnro & _
                        " AND gti_acunov.tpanro = " & objRs!tpanro & _
                        " AND gti_acunov.gpanro = " & nroproc
                End If
                Flog.writeline "Actualizo GTI_Acunov ==> :" & StrSql
                objConn.Execute StrSql, , adExecuteNoRecords
            
            End If
continuar:
            objRs.MoveNext
            
            horas = 0
            
        Loop
        
    
    End If
    
    
End Sub



Public Sub Politica2000(NroProceso As Long, subn As Integer)
    
    Select Case subn
        Case 1
            Call Politica2000Schneider(NroProceso)
        Case 2
            Call Politica2000Expo(NroProceso)
        Case 3
            Call Politica2000MAzul(NroProceso)
    End Select
    
End Sub
Public Sub Politica2000Expo(NroProceso As Long)

Dim fechaDesde As Date
Dim fechaHasta As Date
Dim objConnNexus As New ADODB.Connection
Dim nroLote As Long
Dim rsAD As New ADODB.Recordset
Dim RsAsis As New ADODB.Recordset
Dim objrs_prim As New ADODB.Recordset
Dim objrs_pro As New ADODB.Recordset
Dim strcmdLine As String
Dim FechaTemp As Date
Dim fdesde As Date
Dim fhasta As Date
Dim Cdias As Integer
Dim rs As New ADODB.Recordset
Dim Progreso As Single
Dim CEmpleadosAProc As Integer
Dim IncPorc As Single
   
    OpenConnection "DSN=Nexushr-RHPro", objConnNexus
        
    MyBeginTrans
    
    StrSql = " SELECT MIN( gti_Procacum.gpadesde) AS desde," & _
             "        MAX( gti_Procacum.gpahasta) AS hasta" & _
             " FROM   batch_proceso " & _
             " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro " & _
             " INNER JOIN gti_Procacum   ON Batch_Procacum.gpanro  = gti_Procacum.gpanro " & _
             " WHERE batch_proceso.bpronro = " & NroProceso
    OpenRecordsetRH StrSql, objrs_prim
    
' si està vacio salgo y grabo el estado en el batch_proceso a estado "Sin Datos"
' ---- FGZ 08/04/03
    If objrs_prim.RecordCount = 0 Then
        StrSql = "UPDATE batch_proceso SET bprcestado = 'Sin Datos' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
        Exit Sub
    End If
' ----

    If Not objrs_prim.EOF Then
      fechaDesde = objrs_prim!desde
      fechaHasta = objrs_prim!hasta
    End If
' Como se pueden indicar varios procesos de acumulacion en este proceso batch, se recupera
' la menor y mayor de todas las fechas involucradas. O.D.A. 29/07/2003


    StrSql = "INSERT INTO expo_Lotes(gene_cpto,gene_asis,gene_suce,tran_cpto,tran_asis,tran_suce,gene_usua,gene_hora,gene_fech,fec_des,fec_has,nro_leg_desde,nro_leg_hasta,batch_finalizada) VALUES ("
    StrSql = StrSql & "'Si','Si','Si','No','No','No',null,null,null," & ConvFecha(fechaDesde) & "," & ConvFecha(fechaHasta) & ",null,null,null)"
    objConnNexus.Execute StrSql, , adExecuteNoRecords
    
    ' 22/07/03 FGZ
    nroLote = getLastIdentityExpo(objConnNexus, "expo_Lotes")
    
    
    
    StrSql = " SELECT gti_Procacum.gpadesde," & _
             "        gti_Procacum.gpahasta," & _
             "        Batch_Procacum.gpanro," & _
             "        batch_proceso.bpronro," & _
             "        batch_proceso.iduser" & _
             " FROM   batch_proceso " & _
             " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro " & _
             " INNER JOIN gti_Procacum   ON Batch_Procacum.gpanro  = gti_Procacum.gpanro " & _
             " WHERE batch_proceso.bpronro = " & NroProceso
    OpenRecordsetRH StrSql, objrs_pro
' Se vuelve a armar el SELECT original. O.D.A. 29/07/2003


    ' ------------------------------------------------------------------------
    ' FGZ 30/07/2003
    ' el valor de la columna APLIC_ORIG debería ser la estructura actual de tipo 36 (Centro de empaque)
    ' de cualquiera de los empleados incluidos en el proceso de acumulacion
    ' (GTI_ACUNOV con el GPANRO que corresponde) en lugar de tomar el valor de la estructura 36 del usuario
    '
    ' Así estaba:
    
    'StrSql = "SELECT estructura.estrdabr,estructura.estrcodext FROM usupuedever "
    'StrSql = StrSql & " INNER JOIN estructura ON usupuedever.estrnro = estructura.estrnro AND estructura.tenro = 36 "
    'StrSql = StrSql & " WHERE usupuedever.iduser = '" & objrs_pro!IdUser & "'"
    'OpenRecordsetRH StrSql, objRs
    
    ' StrSql = "INSERT INTO expo_Lotes_batch(nro_lote,aplic_orig) VALUES (" & nroLote & ",'" & objRs!estrcodext & "')"
    ' objConnNexus.Execute StrSql, , adExecuteNoRecords
    '
    '
    ' y yo agregué:
    
    ' FGZ  - 04/08/2003
    ' cambio estruc_actual por his_estructura
    StrSql = " SELECT estructura.estrdabr, estructura.estrcodext FROM gti_acunov " & _
         " INNER JOIN his_estructura ON his_estructura.ternro = gti_acunov.ternro " & _
         " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro AND estructura.tenro = 36 " & _
         " WHERE gti_acunov.gpanro = " & objrs_pro!gpanro & " AND his_estructura.htethasta IS NULL "
    OpenRecordsetRH StrSql, objRs
    
  ' FGZ 30/07/2003
  ' -----------------------------------------------------------------
  
    StrSql = "INSERT INTO expo_Lotes_batch(nro_lote,aplic_orig) VALUES (" & nroLote & ",'" & objRs!estrcodext & "')"
    objConnNexus.Execute StrSql, , adExecuteNoRecords
    
' Seteo las variables de progreso
    CEmpleadosAProc = objrs_pro.RecordCount
    IncPorc = (100 / CEmpleadosAProc)
    Progreso = 0

Do While Not objrs_pro.EOF
    fechaDesde = objrs_pro!gpadesde
    fechaHasta = objrs_pro!gpahasta
' Se setean las fechas correspondientes al proceso de acumulacion que se considera en esta
' vuelta del ciclo. O.D.A. 29/07/2003

    StrSql = " SELECT concepto.ConcCod[1,4],concepto.tconnro,concepto.concnro,empleado.ternro,empleado.empleg,gti_acunov.acnovvalor,gti_acunov.gpanro FROM gti_procacum " & _
             " INNER JOIN gti_acunov ON gti_acunov.gpanro = gti_procacum.gpanro " & _
             " INNER JOIN empleado ON empleado.ternro = gti_acunov.ternro " & _
             " INNER JOIN batch_empleado on batch_empleado.ternro = gti_acunov.ternro " & _
             " AND batch_empleado.bpronro = " & NroProceso & _
             " INNER JOIN Concepto ON Concepto.concnro = gti_acunov.concnro " & _
             " WHERE (gti_acunov.gpanro = " & objrs_pro!gpanro & " )" & _
             " ORDER BY ConcCod[1,4] "
' Agregué el ordenamiento por concepto, para hacer predecible el orden en que se realizan
' las inserciones o actualizaciones. O.D.A. 07/03/2003
' Se toman solo los primeros 4 caracteres del código de los conceptos, ya que existen
' distintas versiones en RHPro que difieren a partir de la 5ta. posición, y que deben
' pasar como un mismo código al liquidador (Feriado Nacional: 0063 y 0063jh).
' O.D.A. 08/04/2003.
' Faltaba el "[1,4]" en el ORDER BY.

    OpenRecordsetRH StrSql, objRs
    
'Flog.writeline "Inicio :" & objrs_pro!gpanro
    
    Do While Not objRs.EOF
        Select Case objRs!tconnro
        Case 1 'Conceptos

          Select Case objRs!conccod
          Case "0055":  'Credito al 100%
                    
'Verifico la existencia de un crédito
            StrSql = "SELECT * FROM expo_Cpto_externo WHERE nro_lote = " & nroLote & _
                     " AND nro_leg = " & objRs!empleg & _
                     " AND tipo_liq = 'MENS' AND cod_cpto = '0054'"
            OpenRecordsetNexus objConnNexus, StrSql, rs
                    
            If rs.EOF Then   'Si no existe
              StrSql = "INSERT INTO expo_Cpto_externo(nro_Leg,tipo_liq,cod_cpto,fec_concepto,fec_activ,fec_vto,valor_base," & _
                       "cantidad,porcentaje,importe_final,centro_costo,nro_lote,nro_correlativo) VALUES (" & _
                       objRs!empleg & ",'MENS','0054'," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & _
                       ",null,0," & objRs!acnovvalor & ",null,null," & nroLote & ",null)"
            Else             'Si ya existe (uno al 50%) actualizo el campo porcentaje
              StrSql = "UPDATE expo_Cpto_externo SET porcentaje = " & objRs!acnovvalor & _
                       "WHERE nro_lote = " & nroLote & _
                       " AND nro_leg = " & objRs!empleg & _
                       " AND tipo_liq = 'MENS' AND cod_cpto = '0054'"
            End If
  
          Case "0054": 'Credito al 50%

'Verifico la existencia de un crédito
            StrSql = "SELECT * FROM expo_Cpto_externo WHERE nro_lote = " & nroLote & _
                     " AND nro_leg = " & objRs!empleg & _
                     " AND tipo_liq = 'MENS' AND cod_cpto = '0054'"
            OpenRecordsetNexus objConnNexus, StrSql, rs
                    
            If rs.EOF Then
              StrSql = "INSERT INTO expo_Cpto_externo(nro_Leg,tipo_liq,cod_cpto,fec_concepto,fec_activ,fec_vto,valor_base," & _
                       "cantidad,porcentaje,importe_final,centro_costo,nro_lote,nro_correlativo) VALUES (" & _
                       objRs!empleg & ",'MENS','" & objRs!conccod & "'," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & _
                       ",null," & objRs!acnovvalor & ",0,null,null," & nroLote & ",null)"
            Else
              StrSql = "UPDATE expo_Cpto_externo SET cantidad = cantidad + " & objRs!acnovvalor & _
                       "WHERE nro_lote = " & nroLote & _
                       " AND nro_leg = " & objRs!empleg & _
                       " AND tipo_liq = 'MENS' AND cod_cpto = '0054'"
            End If

          Case Else  'Otros casos que no sean créditos
            StrSql = "SELECT * FROM expo_Cpto_externo WHERE nro_lote = " & nroLote & _
                     " AND nro_leg = " & objRs!empleg & _
                     " AND tipo_liq = 'MENS' AND cod_cpto = '" & objRs!conccod & "'"
            OpenRecordsetNexus objConnNexus, StrSql, rs
            
            If rs.EOF Then
                StrSql = "INSERT INTO expo_Cpto_externo(nro_Leg,tipo_liq,cod_cpto,fec_concepto,fec_activ,fec_vto,valor_base," & _
                         "cantidad,porcentaje,importe_final,centro_costo,nro_lote,nro_correlativo) VALUES (" & _
                         objRs!empleg & ",'MENS','" & objRs!conccod & "'," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & _
                         ",null," & objRs!acnovvalor & ",0,null,null," & nroLote & ",null)"
            Else
              StrSql = "UPDATE expo_Cpto_externo SET cantidad = cantidad + " & objRs!acnovvalor & _
                       "WHERE nro_lote = " & nroLote & _
                       " AND nro_leg = " & objRs!empleg & _
                       " AND tipo_liq = 'MENS' AND cod_cpto = '" & objRs!conccod & "'"
            End If
          End Select
                
          objConnNexus.Execute StrSql, , adExecuteNoRecords
          If rs.State = adStateOpen Then rs.Close
' Si el codigo del concepto es el "0055" Creditos al 100% en el RHpro, los debe
' insertar en la columna de porcentaje del concepto de credito al 50% "0054" del NexusHR.
' Caso contrario, escribe normalmente.
            
                
        Case 15 'Asistencias
          StrSql = " SELECT * FROM gti_acumdiario " & _
                   " INNER JOIN gti_cab ON gti_cab.ternro = gti_acumdiario.ternro " & _
                   " INNER JOIN gti_det ON gti_det.cgtinro = gti_cab.cgtinro and gti_det.thnro = gti_acumdiario.thnro " & _
                   " INNER JOIN tiph_con on tiph_con.thnro = gti_det.thnro" & _
                   " WHERE tiph_con.concnro = " & objRs!concnro & " AND " & _
                   " gti_acumdiario.ternro = " & objRs!Ternro & _
                   " AND gti_cab.gpanro = " & objRs!gpanro & _
                   " AND adFecha >= " & ConvFecha(fechaDesde) & _
                   " AND adfecha <= " & ConvFecha(fechaHasta) & _
                   " ORDER BY adFecha ASC"
          OpenRecordsetRH StrSql, rsAD
                
          Do While Not rsAD.EOF
            StrSql = "SELECT * FROM expo_Asis_externa " & _
                     " WHERE periodo_pago = '" & Format(fechaDesde, "YYYYMM") & "' " & _
                     " AND   nro_leg      = " & objRs!empleg & _
                     " AND   dia_desde    = " & Day(rsAD!adfecha)
            OpenRecordsetNexus objConnNexus, StrSql, rs
            
            If rs.EOF Then
              StrSql = "INSERT INTO expo_Asis_externa(" & _
                       " Periodo_pago, nro_leg, dia_desde, dia_hasta, estado, nro_lote, nro_correlativo) " & _
                       " VALUES( " & _
                       "'" & Format(fechaDesde, "YYYYMM") & "'," & _
                       objRs!empleg & "," & _
                       Day(rsAD!adfecha) & "," & _
                       Day(rsAD!adfecha) & ",'" & _
                       objRs!conccod & "'," & _
                       nroLote & ",null)"
              objConnNexus.Execute StrSql, , adExecuteNoRecords
            End If
            
            rsAD.MoveNext
          Loop

        Case 16 'Sucesos
          If objRs!conccod <> "SUSF" Then
            StrSql = " SELECT * FROM gti_acumdiario " & _
                     " INNER JOIN gti_cab ON gti_cab.ternro = gti_acumdiario.ternro " & _
                     " INNER JOIN gti_det ON gti_det.cgtinro = gti_cab.cgtinro and gti_det.thnro = gti_acumdiario.thnro " & _
                     " INNER JOIN tiph_con on tiph_con.thnro = gti_det.thnro" & _
                     " WHERE tiph_con.concnro = " & objRs!concnro & " AND " & _
                     " gti_acumdiario.ternro = " & objRs!Ternro & _
                     " AND gti_cab.gpanro = " & objRs!gpanro & _
                     " AND adFecha >= " & ConvFecha(fechaDesde) & _
                     " AND adfecha <= " & ConvFecha(fechaHasta) & _
                     " ORDER BY adFecha ASC"
            OpenRecordsetRH StrSql, rsAD
                
            Do While Not rsAD.EOF
              StrSql = "INSERT INTO expo_Suce_externa(Periodo_pago,nro_leg,dia,cod_suceso,cantidad,horario,nro_lote,nro_correlativo) VALUES (" & _
                       "'" & Format(fechaDesde, "YYYYMM") & "'," & objRs!empleg & ",'" & Format(rsAD!adfecha, "DD") & "','" & objRs!conccod & "'," & rsAD!adcanthoras * 60 & ", Null," & nroLote & ", Null)"
              objConnNexus.Execute StrSql, , adExecuteNoRecords
              rsAD.MoveNext
            Loop
          End If
        End Select
        objRs.MoveNext
        Progreso = Progreso + 1
        Debug.Print Progreso
    Loop
    
    
   ' Actualizar el progreso
   Progreso = Progreso + IncPorc
   StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
   objConn.Execute StrSql, , adExecuteNoRecords

        
  objrs_pro.MoveNext
Loop
    
    
    MyCommitTrans
End Sub


Public Sub Politica2000Schneider(NroProceso As Long)

Dim fechaDesde As Date
Dim fechaHasta As Date
Dim objConnNexus As New ADODB.Connection
Dim nroLote As Long
Dim rsAD As New ADODB.Recordset
Dim RsAsis As New ADODB.Recordset
Dim objrs_pro As New ADODB.Recordset
Dim strcmdLine As String
Dim FechaTemp As Date
Dim fdesde As Date
Dim fhasta As Date
Dim Cdias As Integer
Dim rs As New ADODB.Recordset
Dim Progreso As Single
Dim CEmpleadosAProc As Integer
Dim IncPorc As Single

Dim loRs
   
    OpenConnection "DSN=Nexushr-RHPro", objConnNexus
        
    MyBeginTrans
    
    StrSql = " SELECT gpadesde,gpahasta,Batch_Procacum.gpanro,batch_proceso.bpronro,batch_proceso.iduser" & _
             " From batch_proceso " & _
             " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro " & _
             " INNER JOIN gti_Procacum ON Batch_Procacum.gpanro = gti_Procacum.gpanro " & _
             " Where batch_proceso.bpronro = " & NroProceso
    OpenRecordsetRH StrSql, objrs_pro
    
' si està vacio salgo y grabo el estado en el batch_proceso a estado "Sin Datos"
' ---- FGZ 08/04/03
    If objrs_pro.RecordCount = 0 Then
            
        StrSql = "UPDATE batch_proceso SET bprcestado = 'Sin Datos' WHERE bpronro = " & NroProceso
        objConn.Execute StrSql, , adExecuteNoRecords
    
        Exit Sub
    End If
' ----

    If Not objrs_pro.EOF Then
        fechaDesde = objrs_pro!gpadesde
        fechaHasta = objrs_pro!gpahasta
    End If
    
    
    ' Estaba esto hasta el 20/05/2003
    'StrSql = "INSERT INTO expo_Lotes(gene_cpto,gene_asis,gene_suce,tran_cpto,tran_asis,tran_suce,gene_usua,gene_hora,gene_fech,fec_des,fec_has,nro_leg_desde,nro_leg_hasta,batch_finalizada) VALUES ("
    'StrSql = StrSql & "'Si','Si','Si','No','No','No',null,null,null," & ConvFecha(fechaDesde) & "," & ConvFecha(fechaHasta) & ",null,null,null)"
    'objConnNexus.Execute StrSql, , adExecuteNoRecords
    
    'nroLote = getLastIdentity(objConnNexus, "expo_Lotes")
    ' ------------------
    
    ' FGZ 20/05/2003 ------------------
    ' Puse esto y funciona bien para SQL server 2000
    StrSql = "INSERT INTO expo_Lotes(gene_cpto,gene_asis,gene_suce,tran_cpto,tran_asis,tran_suce,gene_usua,gene_hora,gene_fech,fec_des,fec_has,nro_leg_desde,nro_leg_hasta,batch_finalizada) VALUES ("
    StrSql = StrSql & "'Si','Si','Si','No','No','No',null,null,null," & ConvFecha(fechaDesde) & "," & ConvFecha(fechaHasta) & ",null,null,null)"
    objConnNexus.Execute (StrSql)
    
    StrSql = "SELECT @@IDENTITY AS NewID"
    Set loRs = objConnNexus.Execute(StrSql)
    nroLote = loRs.Fields("NewID").Value
    loRs.Close
    Set loRs = Nothing
    ' -----------------------------------------------------
            
        
    StrSql = "SELECT estructura.estrdabr,estructura.estrcodext FROM usupuedever "
    StrSql = StrSql & " INNER JOIN estructura ON usupuedever.estrnro = estructura.estrnro AND estructura.tenro = 36 "
    StrSql = StrSql & " WHERE usupuedever.iduser = '" & objrs_pro!IdUser & "'"
    OpenRecordsetRH StrSql, objRs
    
    
    StrSql = "INSERT INTO expo_Lotes_batch(nro_lote,aplic_orig) VALUES (" & nroLote & ",'" & objRs!estrcodext & "')"
    objConnNexus.Execute StrSql, , adExecuteNoRecords
    
' Seteo las variables de progreso
CEmpleadosAProc = objrs_pro.RecordCount
IncPorc = (100 / CEmpleadosAProc)
Progreso = 0
    
    
Do While Not objrs_pro.EOF
    
    StrSql = " SELECT concepto.ConcCod[1,4],concepto.tconnro,concepto.concnro,empleado.ternro,empleado.empleg,gti_acunov.acnovvalor,gti_acunov.gpanro FROM gti_procacum " & _
             " INNER JOIN gti_acunov ON gti_acunov.gpanro = gti_procacum.gpanro " & _
             " INNER JOIN empleado ON empleado.ternro = gti_acunov.ternro " & _
             " INNER JOIN batch_empleado on batch_empleado.ternro = gti_acunov.ternro " & _
             " AND batch_empleado.bpronro = " & NroProceso & _
             " INNER JOIN Concepto ON Concepto.concnro = gti_acunov.concnro " & _
             " WHERE (gti_acunov.gpanro = " & objrs_pro!gpanro & " )" & _
             " ORDER BY ConcCod[1,4] "
' Agregué el ordenamiento por concepto, para hacer predecible el orden en que se realizan
' las inserciones o actualizaciones. O.D.A. 07/03/2003
' Se toman solo los primeros 4 caracteres del código de los conceptos, ya que existen
' distintas versiones en RHPro que difieren a partir de la 5ta. posición, y que deben
' pasar como un mismo código al liquidador (Feriado Nacional: 0063 y 0063jh).
' O.D.A. 08/04/2003.
' Faltaba el "[1,4]" en el ORDER BY.

    OpenRecordsetRH StrSql, objRs
    
'Flog.writeline "Inicio :" & objrs_pro!gpanro
    
    Do While Not objRs.EOF
        Select Case objRs!tconnro
        Case 1 'Conceptos

          Select Case objRs!conccod
          Case "0055":  'Credito al 100%
                    
'Verifico la existencia de un crédito
            StrSql = "SELECT * FROM expo_Cpto_externo WHERE nro_lote = " & nroLote & _
                     " AND nro_leg = " & objRs!empleg & _
                     " AND tipo_liq = 'MENS' AND cod_cpto = '0054'"
            OpenRecordset StrSql, rs
                    
            If rs.EOF Then   'Si no existe
              StrSql = "INSERT INTO expo_Cpto_externo(nro_Leg,tipo_liq,cod_cpto,fec_concepto,fec_activ,fec_vto,valor_base," & _
                       "cantidad,porcentaje,importe_final,centro_costo,nro_lote,nro_correlativo) VALUES (" & _
                       objRs!empleg & ",'MENS','0054'," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & _
                       ",null,0," & objRs!acnovvalor & ",null,null," & nroLote & ",null)"
            Else             'Si ya existe (uno al 50%) actualizo el campo porcentaje
              StrSql = "UPDATE expo_Cpto_externo SET porcentaje = " & objRs!acnovvalor & _
                       "WHERE nro_lote = " & nroLote & _
                       " AND nro_leg = " & objRs!empleg & _
                       " AND tipo_liq = 'MENS' AND cod_cpto = '0054'"
            End If
  
          Case "0054": 'Credito al 50%

'Verifico la existencia de un crédito
            StrSql = "SELECT * FROM expo_Cpto_externo WHERE nro_lote = " & nroLote & _
                     " AND nro_leg = " & objRs!empleg & _
                     " AND tipo_liq = 'MENS' AND cod_cpto = '0054'"
            OpenRecordset StrSql, rs
                    
            If rs.EOF Then
              StrSql = "INSERT INTO expo_Cpto_externo(nro_Leg,tipo_liq,cod_cpto,fec_concepto,fec_activ,fec_vto,valor_base," & _
                       "cantidad,porcentaje,importe_final,centro_costo,nro_lote,nro_correlativo) VALUES (" & _
                       objRs!empleg & ",'MENS','" & objRs!conccod & "'," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & _
                       ",null," & objRs!acnovvalor & ",0,null,null," & nroLote & ",null)"
            Else
              StrSql = "UPDATE expo_Cpto_externo SET cantidad = " & objRs!acnovvalor & _
                       "WHERE nro_lote = " & nroLote & _
                       " AND nro_leg = " & objRs!empleg & _
                       " AND tipo_liq = 'MENS' AND cod_cpto = '0054'"
            End If

          Case Else  'Otros casos que no sean créditos
            StrSql = "SELECT * FROM expo_Cpto_externo WHERE nro_lote = " & nroLote & _
                     " AND nro_leg = " & objRs!empleg & _
                     " AND tipo_liq = 'MENS' AND cod_cpto = '0054'"
            OpenRecordset StrSql, rs
            
            If rs.EOF Then
                StrSql = "INSERT INTO expo_Cpto_externo(nro_Leg,tipo_liq,cod_cpto,fec_concepto,fec_activ,fec_vto,valor_base," & _
                         "cantidad,porcentaje,importe_final,centro_costo,nro_lote,nro_correlativo) VALUES (" & _
                         objRs!empleg & ",'MENS','" & objRs!conccod & "'," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & "," & ConvFecha(fechaHasta) & _
                         ",null," & objRs!acnovvalor & ",0,null,null," & nroLote & ",null)"
            Else
              StrSql = "UPDATE expo_Cpto_externo SET cantidad = " & objRs!acnovvalor & _
                       "WHERE nro_lote = " & nroLote & _
                       " AND nro_leg = " & objRs!empleg & _
                       " AND tipo_liq = 'MENS' AND cod_cpto = " & objRs!conccod
            End If
          End Select
                
          objConnNexus.Execute StrSql, , adExecuteNoRecords
          If rs.State = adStateOpen Then rs.Close
' Si el codigo del concepto es el "0055" Creditos al 100% en el RHpro, los debe
' insertar en la columna de porcentaje del concepto de credito al 50% "0054" del NexusHR.
' Caso contrario, escribe normalmente.
            
                
        Case 15 'Asistencias
          StrSql = " SELECT * FROM gti_acumdiario " & _
                   " INNER JOIN gti_cab ON gti_cab.ternro = gti_acumdiario.ternro " & _
                   " INNER JOIN gti_det ON gti_det.cgtinro = gti_cab.cgtinro and gti_det.thnro = gti_acumdiario.thnro " & _
                   " INNER JOIN tiph_con on tiph_con.thnro = gti_det.thnro" & _
                   " WHERE tiph_con.concnro = " & objRs!concnro & " AND " & _
                   " gti_acumdiario.ternro = " & objRs!Ternro & _
                   " AND gti_cab.gpanro = " & objRs!gpanro & _
                   " AND adFecha >= " & ConvFecha(fechaDesde) & _
                   " AND adfecha <= " & ConvFecha(fechaHasta) & _
                   " ORDER BY adFecha ASC"
          OpenRecordsetRH StrSql, rsAD
                
          Do While Not rsAD.EOF
            StrSql = "INSERT INTO expo_Asis_externa(Periodo_pago,nro_leg,dia_desde,dia_hasta,estado,nro_lote,nro_correlativo) VALUES (" & _
                     "'" & Format(fechaDesde, "YYYYMM") & "'," & objRs!empleg & "," & Day(rsAD!adfecha) & "," & Day(rsAD!adfecha) & ",'" & objRs!conccod & "'," & nroLote & ",null)"
            objConnNexus.Execute StrSql, , adExecuteNoRecords
            rsAD.MoveNext
          Loop

        Case 16 'Sucesos
          If objRs!conccod <> "SUSF" Then
            StrSql = " SELECT * FROM gti_acumdiario " & _
                     " INNER JOIN gti_cab ON gti_cab.ternro = gti_acumdiario.ternro " & _
                     " INNER JOIN gti_det ON gti_det.cgtinro = gti_cab.cgtinro and gti_det.thnro = gti_acumdiario.thnro " & _
                     " INNER JOIN tiph_con on tiph_con.thnro = gti_det.thnro" & _
                     " WHERE tiph_con.concnro = " & objRs!concnro & " AND " & _
                     " gti_acumdiario.ternro = " & objRs!Ternro & _
                     " AND gti_cab.gpanro = " & objRs!gpanro & _
                     " AND adFecha >= " & ConvFecha(fechaDesde) & _
                     " AND adfecha <= " & ConvFecha(fechaHasta) & _
                     " ORDER BY adFecha ASC"
            OpenRecordsetRH StrSql, rsAD
                
            Do While Not rsAD.EOF
              StrSql = "INSERT INTO expo_Suce_externa(Periodo_pago,nro_leg,dia,cod_suceso,cantidad,horario,nro_lote,nro_correlativo) VALUES (" & _
                       "'" & Format(fechaDesde, "YYYYMM") & "'," & objRs!empleg & ",'" & Format(rsAD!adfecha, "DD") & "','" & objRs!conccod & "'," & rsAD!adcanthoras * 60 & ", Null," & nroLote & ", Null)"
              objConnNexus.Execute StrSql, , adExecuteNoRecords
              rsAD.MoveNext
            Loop
          End If
        End Select
        objRs.MoveNext
        Progreso = Progreso + 1
        Debug.Print Progreso
    Loop
    
    
      'Actualizo el progreso
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
      
  objrs_pro.MoveNext
Loop
    
    
    MyCommitTrans
End Sub

Public Sub Politica550()

    usaSabadoDomingo = True

End Sub

Public Sub politica560()

    usaTurnoTrasnoche = True

End Sub

Public Sub politica570()

    UsaConversionHoras = True

End Sub
Public Sub politica575()

    usaHorasExtras = True

End Sub

Public Sub politica580()

    usaTopesHorasExtras = True

End Sub

Public Sub politica590()

    usaDesgloseAP = True

End Sub


Public Sub Politica2000MAzul(NroProceso As Long)

Dim fechaDesde As Date
Dim fechaHasta As Date
Dim nroLote As Long
Dim rsAD As New ADODB.Recordset
Dim RsAsis As New ADODB.Recordset
Dim objrs_pro As New ADODB.Recordset
Dim objrs_estr As New ADODB.Recordset
Dim strcmdLine As String
Dim FechaTemp As Date
Dim fdesde As Date
Dim fhasta As Date
Dim Cdias As Integer

Dim cantidad As Integer
Dim Cod_emp As Long

Dim Archivo As String
Dim pos As Integer
Dim freg

Dim pos1 As Integer
Dim pos2 As Integer

Dim Int_Valor As Integer
Dim Dec_Valor As String

    StrSql = " SELECT gpadesde,gpahasta,Batch_Procacum.gpanro,batch_proceso.bpronro " & _
             " From batch_proceso " & _
             " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro " & _
             " INNER JOIN gti_Procacum ON Batch_Procacum.gpanro = gti_Procacum.gpanro " & _
             " Where batch_proceso.bpronro = " & NroProceso
    
    OpenRecordset StrSql, objrs_pro
    

    If Not objrs_pro.EOF Then
        fechaDesde = objrs_pro!gpadesde
        fechaHasta = objrs_pro!gpahasta
    End If
    
    Archivo = "c:\rhprox2\algoliq\" & NroProceso & " - Exportacion NOV a LIQ " & Format(fechaDesde, "DD-MM-YYYY") & " al " & Format(fechaHasta, "DD-MM-YYYY") & ".txt"

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set freg = fs.CreateTextFile(Archivo, True)
    
    MyBeginTrans
   
Cod_emp = 0
cantidad = 4

Do While Not objrs_pro.EOF
   
    StrSql = " SELECT concepto.ConcCod,concepto.tconnro,concepto.concnro,empleado.ternro,empleado.empleg,gti_acunov.acnovvalor,gti_acunov.gpanro,gti_acunov.acnovfecaprob FROM gti_procacum " & _
             " INNER JOIN gti_acunov ON gti_acunov.gpanro = gti_procacum.gpanro " & _
             " INNER JOIN empleado ON empleado.ternro = gti_acunov.ternro " & _
             " INNER JOIN Concepto ON Concepto.concnro = gti_acunov.concnro " & _
             " WHERE (gti_acunov.gpanro = " & objrs_pro!gpanro & " )" & _
             " ORDER BY concepto.conccod "
    OpenRecordset StrSql, objRs
   
    Do While Not objRs.EOF
    
        If objRs!empleg <> Cod_emp Then
         cantidad = 5
         Cod_emp = objRs!empleg
        Else
         cantidad = cantidad + 1
        End If
        
        StrSql = " SELECT estrcodext FROM estruc_actual "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = estruc_actual.estrnro WHERE estruc_actual.tenro = 2 and estruc_actual.ternro = " & objRs!Ternro
        OpenRecordset StrSql, objrs_estr
    
        Int_Valor = Int(objRs!acnovvalor)
        Dec_Valor = (objRs!acnovvalor - Int(objRs!acnovvalor)) * 100
        
        freg.writeline Chr(34) & _
                       Format(objRs!empleg, "00000") & _
                       Format(objRs!conccod, "000") & _
                       Format(Int_Valor, "0000000") & _
                       Format(Dec_Valor, "00") & _
                       "000000000" & _
                       Format(cantidad, "0") & _
                       Format(objRs!acnovfecaprob, "yyyymmdd") & _
                       Chr(34)
        
        objRs.MoveNext
        
    Loop
    
  objrs_pro.MoveNext
Loop
    MyCommitTrans
End Sub


Public Sub Politica2000MAzul_old(NroProceso As Long)

Dim fechaDesde As Date
Dim fechaHasta As Date
Dim objConnNexus As New ADODB.Connection
Dim nroLote As Long
Dim rsAD As New ADODB.Recordset
Dim RsAsis As New ADODB.Recordset
Dim objrs_pro As New ADODB.Recordset
Dim objrs_estr As New ADODB.Recordset
Dim strcmdLine As String
Dim FechaTemp As Date
Dim fdesde As Date
Dim fhasta As Date
Dim Cdias As Integer

Dim Archivo As String
Dim pos As Integer
Dim freg

Dim Progreso As Single
Dim CEmpleadosAProc As Integer
Dim IncPorc As Single


'OpenConnection strconexion, objConn
    
    StrSql = " SELECT gpadesde,gpahasta,Batch_Procacum.gpanro,batch_proceso.bpronro " & _
             " From batch_proceso " & _
             " INNER JOIN Batch_Procacum ON Batch_Procacum.bpronro = batch_proceso.bpronro " & _
             " INNER JOIN gti_Procacum ON Batch_Procacum.gpanro = gti_Procacum.gpanro " & _
             " Where batch_proceso.bpronro = " & NroProceso
    
    OpenRecordset StrSql, objrs_pro
    

    If Not objrs_pro.EOF Then
        fechaDesde = objrs_pro!gpadesde
        fechaHasta = objrs_pro!gpahasta
    End If
    
    
    Archivo = "c:\rhprox2\algoliq\Exportacion NOV a LIQ " & Format(fechaDesde, "DD-MM-YYYY") & " al " & Format(fechaHasta, "DD-MM-YYYY") & ".txt"

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set freg = fs.CreateTextFile(Archivo, True)
    
    MyBeginTrans
   
' Seteo las variables de progreso
CEmpleadosAProc = objrs_pro.RecordCount
IncPorc = (100 / CEmpleadosAProc)
Progreso = 0
   
    
Do While Not objrs_pro.EOF
   
    StrSql = " SELECT concepto.ConcCod,concepto.tconnro,concepto.concnro,empleado.ternro,empleado.empleg,gti_acunov.acnovvalor,gti_acunov.gpanro FROM gti_procacum " & _
             " INNER JOIN gti_acunov ON gti_acunov.gpanro = gti_procacum.gpanro " & _
             " INNER JOIN empleado ON empleado.ternro = gti_acunov.ternro " & _
             " INNER JOIN batch_empleado on batch_empleado.ternro = gti_acunov.ternro " & _
             " AND batch_empleado.bpronro = " & NroProceso & _
             " INNER JOIN Concepto ON Concepto.concnro = gti_acunov.concnro " & _
             " WHERE (gti_acunov.gpanro = " & objrs_pro!gpanro & " )"

    OpenRecordset StrSql, objRs
   
    Do While Not objRs.EOF
    
        'FGZ  - 04/08/2003
        ' Cambio estruc_actual por his_estructura
        StrSql = " SELECT estrcodext FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro WHERE his_estructura.htethasta IS NULL AND his_estructura.tenro = 2 AND his_estructura.ternro = " & objRs!Ternro
        OpenRecordset StrSql, objrs_estr
    
        freg.writeline Chr(34) & Format(objRs!empleg, "00000") & Format(objRs!conccod, "000") & Format(objRs!acnovvalor, "00000") & "00000" & Format(Now, "yyyymmdd") & objrs_estr!estrcodext & Chr(34)
                       
        objRs.MoveNext
    Loop
    
    
    'Actualizo el progreso
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
    objConn.Execute StrSql, , adExecuteNoRecords
    
  objrs_pro.MoveNext
Loop
    MyCommitTrans
End Sub


Public Sub Politica850()

    usaExcedentesHorasNormales = True

End Sub



