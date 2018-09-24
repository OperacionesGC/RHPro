Attribute VB_Name = "BusqAntigVac"
Option Explicit

' Procedimienos de Busqueda de Parámetros
Public Sub bus_Antiguedad(ByVal TipoAnt As String, ByVal Fec_Fin As Date, ByRef dia As Integer, ByRef mes As Integer, ByRef anio As Integer, ByRef DiasHabiles As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad al dia de hoy de un empleado en :
'              dias hábiles(si es menor que un año) o en dias, meses y años en caso contrario.
'              antiguedad.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
Dim fecalta As Date
Dim fecbaja As Date
Dim seguir As Date
Dim q As Integer

Dim NombreCampo As String

Dim rs_Fases As New ADODB.Recordset


NombreCampo = ""
DiasHabiles = 0

Select Case UCase(TipoAnt)
Case "SUELDO":
    NombreCampo = "sueldo"
Case "INDEMNIZACION":
    NombreCampo = "indemnizacion"
Case "VACACIONES":
    NombreCampo = "vacaciones"
Case "REAL":
    NombreCampo = "real"
Case Else
End Select

StrSql = "SELECT * FROM fases WHERE empleado = " & empternro & _
         " AND " & NombreCampo & " = -1" & _
         " ORDER BY altfec,bajfec"
OpenRecordset StrSql, rs_Fases

Do While Not rs_Fases.EOF
    If (IsNull(rs_Fases!altfec)) Or (IsNull(rs_Fases!bajfec) And rs_Fases!estado = 0) Or (rs_Fases!altfec >= Fec_Fin) Then
        GoTo siguiente
    Else
        fecalta = rs_Fases!altfec
    End If
    
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If rs_Fases!estado Then
        fecbaja = Fec_Fin ' solo es un alta, tomar el fecha-fin
    ElseIf rs_Fases!bajfec <= Fec_Fin Then
        fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
    Else
        fecbaja = Fec_Fin ' hasta la fecha ingresada
    End If
    
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    dia = dia + aux1
    mes = mes + aux2 + Int(dia / 30)
    anio = anio + aux3 + Int(mes / 12)
    dia = dia Mod 30
    mes = mes Mod 12

' ------------
'        aux1 = aux1 + 1
'        If aux1 < 365 Then
'            aux3 = 0
'        End If
'        If aux1 = aux2 And aux2 = aux3 And aux3 = 0 Then
'            aux1 = 1
'        End If
'        aux1 = aux1 Mod 30
'        aux2 = aux2 Mod 12
'        dia = dia + aux1
'        Mes = Mes + aux2 + Int(dia / 30)
'
'        If Int(dia / 30) >= 1 Then
'            dia = dia Mod 30
'        End If
'        Anio = Anio + aux3 + Int(Mes / 12)
'        If Int(Mes / 12) >= 1 Then
'            Mes = Mes Mod 12
'        End If
' ------------

    If anio = 0 Then
        Call DiasTrab(fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
siguiente:
    rs_Fases.MoveNext
Loop

If anio <> 0 Then
    DiasHabiles = 0
End If


' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub


Public Sub DiasTrab(ByVal Desde As Date, ByVal hasta As Date, ByRef DiasH As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias trabajados de acuerdo al turno en que se trabaja y
'              de acuerdo a los dias que figuran como feriados en la tabla de feriados.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim d1 As Integer
Dim d2 As Integer
Dim Aux As Integer
Dim aux2 As Integer
Dim dxsem As Integer

Dim rs_pais As New ADODB.Recordset
Dim rs_feriados As New ADODB.Recordset

    dxsem = 5
    
    d1 = Weekday(Desde)
    d2 = Weekday(hasta)
    
    Aux = DateDiff("d", Desde, hasta) + 1
    If Aux < 7 Then
        DiasH = Minimo(Aux, dxsem)
    Else
        If Aux = 7 Then
            DiasH = dxsem
        Else
            aux2 = 8 - d1 + d2
            If aux2 < 7 Then
                aux2 = Minimo(aux2, dxsem)
            Else
                If aux2 = 7 Then
                    aux2 = dxsem
                End If
            End If
            
            If aux2 >= 7 Then
                aux2 = Abs(aux2 - 7) + Int(aux2 / 7) * dxsem
            Else
                aux2 = aux2 + Int((aux2 - aux2) / 7) * dxsem
            End If
        End If
    End If
    
    Aux = 0
    
    StrSql = "SELECT * FROM pais INNER JOIN tercero ON tercero.paisnro = pais.paisnro WHERE tercero.ternro = " & empternro
    OpenRecordset StrSql, rs_pais
    
    If Not rs_pais.EOF Then
        ' Resto los Feriados Nacionales
        StrSql = "SELECT * FROM feriado WHERE tipferinro = 2 " & _
                 " AND fericodext = " & rs_pais!paisnro & _
                 " AND ferifecha >= " & ConvFecha(Desde) & _
                 " AND ferifecha < " & ConvFecha(hasta)
        OpenRecordset StrSql, rs_feriados
        
        Do While Not rs_feriados.EOF
            If Weekday(rs_feriados!ferifecha) > 1 Then
                DiasH = DiasH - 1
            End If
            
            ' Siguiente Feriado
            rs_feriados.MoveNext
        Loop
    End If


    ' Resto los feriados por Convenio
    StrSql = "SELECT * FROM empleado INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro " & _
             " INNER JOIN fer_estr ON fer_estr.tenro = his_estructura.tenro " & _
             " INNER JOIN feriado ON fer_estr.ferinro = feriado.ferinro " & _
             " WHERE empleado.ternro = " & empternro & _
             " AND feriado.tipferinro = 2" & _
             " AND feriado.ferifecha >= " & ConvFecha(Desde) & _
             " AND feriado.ferifecha < " & ConvFecha(hasta)
    OpenRecordset StrSql, rs_feriados
    
    Do While Not rs_feriados.EOF
        If Weekday(rs_feriados!ferifecha) > 1 Then
            DiasH = DiasH - 1
        End If
        
        ' Siguiente Feriado
        rs_feriados.MoveNext
    Loop
    
    
    ' cierro todo y libero
    If rs_pais.State = adStateOpen Then rs_pais.Close
    If rs_feriados.State = adStateOpen Then rs_feriados.Close
        
    Set rs_feriados = Nothing
    Set rs_pais = Nothing

End Sub


Public Sub bus_Estructura(ByVal NroProg As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Estructura a una Fecha
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoEstr As Long      ' Tipo de Estructura
Dim TipoFecha As Integer    ' 1 - Primer dia del año
                            ' 2 - Ultimo dia del año
                            ' 3 - Inicio del proceso
                            ' 4 - Fin del proceso
                            ' 5 - Inicio del Periodo
                            ' 6 - Fin del periodo
                            ' 7 - Today

Dim param_cur As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim objRs As New ADODB.Recordset

Dim Aux_Fecha As Date

   
    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    OpenRecordset StrSql, param_cur
    
    If Not param_cur.EOF Then
        TipoEstr = param_cur!Auxint1
        TipoFecha = param_cur!Auxint2
    Else
        Exit Sub
    End If


    ' busco la ultima fase del empleado  para saber
    ' si entró dentro de la fecha desde y hasta del periodo de vacaciones.
    ' en ese caso la fecha para buscar la estructura es la fecha de alta de la ultima fase.
    
    'Ej.
    'Fecha_Desde del Periodo de Vac. = 01/10/2003
    'Fecha_Hasta del Periodo de Vac. = 30/04/2004
    'y la fecha de alta de la ultima fase es 12/10/2003
    
    ' == > deberia buscar la estructura a la fecha de alta de la ultima fase, es decir, 12/10/2004
    
    StrSql = "SELECT * FROM fases WHERE fases.empleado = " & empternro & " AND fases.vacaciones = -1"
    OpenRecordset StrSql, objRs
    
    If Not objRs.EOF Then
        objRs.MoveLast
        
        If objRs!altfec > fecha_desde Then
            Aux_Fecha = objRs!altfec
        Else
            Aux_Fecha = fecha_desde
        End If
    End If

'antes
'Aux_Fecha = fecha_desde

If Not IsNull(Aux_Fecha) Then
    ' Busco de estructura
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & empternro & " AND " & _
                 " tenro =" & TipoEstr & " AND " & _
                 " (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND " & _
                 " ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
    
        If Not rs_Estructura.EOF Then
            Valor = rs_Estructura!estrnro
            Bien = True
        End If
End If
    

' Cierro todo y libero
If param_cur.State = adStateOpen Then param_cur.Close
Set param_cur = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

If objRs.State = adStateOpen Then objRs.Close
Set objRs = Nothing

End Sub


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

