Attribute VB_Name = "MdlTiposBusquedas"
Option Explicit

Dim objFeriado      As New Feriado

' Procedimienos de Busqueda de Parámetros
Public Sub bus_Antiguedad(ByVal TipoAnt As String, ByVal fechaFin As Date, ByRef Dia As Long, ByRef mes As Long, ByRef Anio As Long, ByRef DiasHabiles As Integer)
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

Flog.writeline "Fecha hasta " & fechaFin
StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro & _
         " AND " & NombreCampo & " = -1" & _
         " ORDER BY altfec,bajfec"
OpenRecordset StrSql, rs_Fases
If rs_Fases.EOF Then
    Flog.writeline "No hay Fases, SQL ==> " & StrSql
End If
Do While Not rs_Fases.EOF
    If (IsNull(rs_Fases!altfec)) Or (IsNull(rs_Fases!bajfec) And rs_Fases!estado = 0) Or (rs_Fases!altfec >= fechaFin) Then
        Flog.writeline "Fase no considerada"
        GoTo siguiente
    Else
        fecalta = rs_Fases!altfec
    End If
    
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If rs_Fases!estado Then
        fecbaja = fechaFin ' solo es un alta, tomar el fecha-fin
    ElseIf rs_Fases!bajfec <= fechaFin Then
        fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
    Else
        fecbaja = fechaFin ' hasta la fecha ingresada
    End If
    Flog.writeline "Diferencia de fechas entre " & fecalta & " y " & fecbaja
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    Dia = Dia + aux1
    mes = mes + aux2 + Int(Dia / 30)
    Anio = Anio + aux3 + Int(mes / 12)
    Dia = Dia Mod 30
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

    If Anio = 0 Then
        Call DiasTrab(fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
siguiente:
    rs_Fases.MoveNext
Loop

If Anio <> 0 Then
    DiasHabiles = 0
End If


' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub

' Procedimienos de Busqueda de Parámetros
Public Sub bus_Antiguedad_R(ByVal TipoAnt As String, ByVal fechaFin As Date, ByRef Dia As Long, ByRef mes As Long, ByRef Anio As Long, ByRef DiasHabiles As Integer)
' ----------------------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad al dia de hoy de un empleado en :
'              dias hábiles(si es menor que un año) o en dias, meses y años en caso contrario.
'              antiguedad. Tiene en cuenta si el empleado trabajo mas de la mitad del año
' Autor      : Gustavo Ring
' Fecha      : 02/03/2009
' Ultima Mod.:
' Descripcion:
' --------------------------------------------------------------------------------------------------------

Dim aux1        As Integer
Dim aux2        As Integer
Dim aux3        As Integer
Dim fecalta     As Date
Dim fecbaja     As Date
Dim seguir      As Date
Dim q           As Integer
Dim fechadesde  As Date
Dim NombreCampo As String
Dim porcentaje  As Boolean

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

fechadesde = CDate("01/01/" & Year(fechaFin))
Flog.writeline "Fecha desde " & fechadesde
Flog.writeline "Fecha hasta " & fechaFin

'StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro
'StrSql = StrSql & " AND " & NombreCampo & " = -1 "
'StrSql = StrSql & " AND ((fases.altfec <= " & ConvFecha(fechaFin) & " AND fases.altfec >= " & ConvFecha(fechaDesde) & ")"
'StrSql = StrSql & " OR (fases.bajfec <= " & ConvFecha(fechaFin) & " AND fases.bajfec >= " & ConvFecha(fechaDesde) & ")"
'StrSql = StrSql & " OR (fases.altfec <= " & ConvFecha(fechaFin) & " AND fases.bajfec is NULL))"
'StrSql = StrSql & " ORDER BY altfec,bajfec "

StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro
StrSql = StrSql & " AND " & NombreCampo & " = -1 "
StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechaFin) & ")"
StrSql = StrSql & " AND (( fases.bajfec >= " & ConvFecha(fechadesde) & " OR fases.bajfec is NULL))"
StrSql = StrSql & " ORDER BY altfec,bajfec "
OpenRecordset StrSql, rs_Fases
If rs_Fases.EOF Then
    Flog.writeline "No hay Fases, SQL ==> " & StrSql
End If

Do While Not rs_Fases.EOF
    If (IsNull(rs_Fases!altfec)) Or (IsNull(rs_Fases!bajfec) And rs_Fases!estado = 0) Or (rs_Fases!altfec >= fechaFin) Then
        Flog.writeline "Fase no considerada"
        GoTo siguiente
    Else
        fecalta = rs_Fases!altfec
    End If
    
    If rs_Fases!altfec < fechadesde Then
        fecalta = fechadesde
    End If
    
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If Not (IsNull(rs_Fases!bajfec)) Then
        If rs_Fases!bajfec <= fechaFin Then
            fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
        Else
            fecbaja = fechaFin ' hasta la fecha ingresada
        End If
    Else
        fecbaja = fechaFin ' hasta la fecha ingresada
    End If
    
    Flog.writeline "Diferencia de fechas entre " & fecalta & " y " & fecbaja
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    Dia = Dia + aux1
    mes = mes + aux2 + Int(Dia / 30)
    Anio = Anio + aux3 + Int(mes / 12)
    Dia = Dia Mod 30
    mes = mes Mod 12

    If Anio = 0 Then
        Call DiasTrab(fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
siguiente:
    rs_Fases.MoveNext
Loop

If Anio <> 0 Then
    DiasHabiles = 0
End If


' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub


' Procedimienos de Busqueda de Parámetros
Public Sub bus_Antiguedad_RV7(ByVal TipoAnt As String, ByVal fechaFin As Date, ByRef Dia As Long, ByRef mes As Long, ByRef Anio As Long, ByRef DiasHabiles As Integer)
' ----------------------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad al dia de hoy de un empleado en :
'              dias hábiles(si es menor que un año) o en dias, meses y años en caso contrario.
'              antiguedad. Tiene en cuenta si el empleado trabajo mas de la mitad del año
' Autor      : Mauricio Zwenger
' Fecha      : 07/08/2014
' Ultima Mod.:
' Descripcion:
' --------------------------------------------------------------------------------------------------------

Dim aux1        As Integer
Dim aux2        As Integer
Dim aux3        As Integer
Dim fecalta     As Date
Dim fecbaja     As Date
Dim seguir      As Date
Dim q           As Integer
Dim fechadesde  As Date
Dim NombreCampo As String
Dim porcentaje  As Boolean

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

fechadesde = CDate("01/01/" & Year(fechaFin))
Flog.writeline "Fecha desde " & fechadesde
Flog.writeline "Fecha hasta " & fechaFin

'StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro
'StrSql = StrSql & " AND " & NombreCampo & " = -1 "
'StrSql = StrSql & " AND ((fases.altfec <= " & ConvFecha(fechaFin) & " AND fases.altfec >= " & ConvFecha(fechaDesde) & ")"
'StrSql = StrSql & " OR (fases.bajfec <= " & ConvFecha(fechaFin) & " AND fases.bajfec >= " & ConvFecha(fechaDesde) & ")"
'StrSql = StrSql & " OR (fases.altfec <= " & ConvFecha(fechaFin) & " AND fases.bajfec is NULL))"
'StrSql = StrSql & " ORDER BY altfec,bajfec "

'StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro
'StrSql = StrSql & " AND " & NombreCampo & " = -1 "
'StrSql = StrSql & " AND (fases.altfec <= " & ConvFecha(fechaFin) & ")"
'StrSql = StrSql & " AND (( fases.bajfec >= " & ConvFecha(fechadesde) & " OR fases.bajfec is NULL))"
'StrSql = StrSql & " ORDER BY altfec,bajfec "

StrSql = "SELECT empfaltagr, empfecbaja FROM empleado Where ternro = " & Ternro

OpenRecordset StrSql, rs_Fases
If rs_Fases.EOF Then
    Flog.writeline "No hay Fases, SQL ==> " & StrSql
End If

If Not Not rs_Fases.EOF Then
    If Not IsNull(rs_Fases!empfaltagr) Then
        fecalta = rs_Fases!empfaltagr
    Else
        Flog.writeline "El empleado no tiene Fecha de Alta!"
        GoTo siguiente
    End If
    
    If rs_Fases!empfaltagr < fechadesde Then
        fecalta = fechadesde
    End If
    
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If Not (IsNull(rs_Fases!empfecbaja)) Then
        If rs_Fases!empfecbaja <= fechaFin Then
            fecbaja = rs_Fases!empfecbaja  ' se trata de un registro completo
        Else
            fecbaja = fechaFin ' hasta la fecha ingresada
        End If
    Else
        fecbaja = fechaFin ' hasta la fecha ingresada
    End If
    
    Flog.writeline "Diferencia de fechas entre " & fecalta & " y " & fecbaja
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    Dia = Dia + aux1
    mes = mes + aux2 + Int(Dia / 30)
    Anio = Anio + aux3 + Int(mes / 12)
    Dia = Dia Mod 30
    mes = mes Mod 12

    If Anio = 0 Then
        Call DiasTrab(fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
End If

siguiente:


If Anio <> 0 Then
    DiasHabiles = 0
End If


' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub


Public Sub bus_Antiguedad_CR(ByVal TipoAnt As String, ByVal fechaFin As Date, ByRef Dia As Long, ByRef mes As Long, ByRef Anio As Long, ByRef DiasHabiles As Integer)
' ----------------------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad al dia de hoy de un empleado en :
'              dias hábiles(si es menor que un año) o en dias, meses y años en caso contrario.
'              antiguedad. Tiene en cuenta si el empleado trabajo mas de la mitad del año
' Autor      : FGZ
' Fecha      : 02/06/2011
' Ultima Mod.:
' Descripcion:
' --------------------------------------------------------------------------------------------------------
Dim aux1        As Integer
Dim aux2        As Integer
Dim aux3        As Integer
Dim fecalta     As Date
Dim fecbaja     As Date
Dim seguir      As Date
Dim q           As Integer
Dim fechadesde  As Date
Dim NombreCampo As String
Dim porcentaje  As Boolean

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

fechadesde = CDate("01/01/" & Year(fechaFin))
Flog.writeline "Fecha desde " & fechadesde
Flog.writeline "Fecha hasta " & fechaFin

'StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro
'StrSql = StrSql & " AND " & NombreCampo & " = -1 "
'StrSql = StrSql & " AND ((fases.altfec <= " & ConvFecha(fechaFin) & " AND fases.altfec >= " & ConvFecha(fechaDesde) & ")"
'StrSql = StrSql & " OR (fases.bajfec <= " & ConvFecha(fechaFin) & " AND fases.bajfec >= " & ConvFecha(fechaDesde) & ")"
'StrSql = StrSql & " OR (fases.altfec <= " & ConvFecha(fechaFin) & " AND fases.bajfec is NULL))"
'StrSql = StrSql & " ORDER BY altfec,bajfec "

StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro
StrSql = StrSql & " AND " & NombreCampo & " = -1 "
StrSql = StrSql & " AND fases.altfec <= " & ConvFecha(fechaFin)
StrSql = StrSql & " ORDER BY altfec,bajfec "
OpenRecordset StrSql, rs_Fases
If rs_Fases.EOF Then
    Flog.writeline "No hay Fases, SQL ==> " & StrSql
End If

Do While Not rs_Fases.EOF
    If (IsNull(rs_Fases!altfec)) Or (IsNull(rs_Fases!bajfec) And rs_Fases!estado = 0) Or (rs_Fases!altfec >= fechaFin) Then
        Flog.writeline "Fase no considerada"
        GoTo siguiente
    Else
        fecalta = rs_Fases!altfec
    End If
    
    'If rs_Fases!altfec < fechaDesde Then
    '    fecalta = fechaDesde
    'End If
    
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If Not (IsNull(rs_Fases!bajfec)) Then
        If rs_Fases!bajfec <= fechaFin Then
            fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
        Else
            fecbaja = fechaFin ' hasta la fecha ingresada
        End If
    Else
        fecbaja = fechaFin ' hasta la fecha ingresada
    End If
    
    Flog.writeline "Diferencia de fechas entre " & fecalta & " y " & fecbaja
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    Dia = Dia + aux1
    mes = mes + aux2 + Int(Dia / 30)
    Anio = Anio + aux3 + Int(mes / 12)
    Dia = Dia Mod 30
    mes = mes Mod 12

    'If Anio = 0 Then
    '    Call DiasTrab(fecalta, fecbaja, aux1)
    '    DiasHabiles = DiasHabiles + aux1
    'End If
    
siguiente:
    rs_Fases.MoveNext
Loop

'If Anio <> 0 Then
'    DiasHabiles = 0
'End If


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
' Ultima Mod.: 17/04/2009. Gustavo Ring.- Se modifico el calculo de los feriados.-
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim d1 As Integer
Dim d2 As Integer
Dim Aux As Integer
Dim aux2 As Integer
Dim dxsem As Integer
Dim excluyeFeriado As Boolean
Dim EsFeriado As Boolean
Dim habiles(7) As Boolean
Dim rs As New ADODB.Recordset
Dim rs_feriados As New ADODB.Recordset

    excluyeFeriado = False
    
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
    
 ' Calculo si el periodo vacacional excluye feriados
    StrSql = "SELECT tpvferiado,tpvhabiles__1,tpvhabiles__2,tpvhabiles__3,tpvhabiles__4,tpvhabiles__5,tpvhabiles__6,tpvhabiles__7 FROM tipovacac WHERE tipvacnro = " & TipoVacacionProporcion
    OpenRecordset StrSql, rs
            
    If Not rs.EOF Then
         excluyeFeriado = rs!tpvferiado
         ' Creo el arreglo con los días feriados
         habiles(1) = rs!tpvhabiles__1
         habiles(2) = rs!tpvhabiles__2
         habiles(3) = rs!tpvhabiles__3
         habiles(4) = rs!tpvhabiles__4
         habiles(5) = rs!tpvhabiles__5
         habiles(6) = rs!tpvhabiles__6
         habiles(7) = rs!tpvhabiles__7
    End If
    
    If Not excluyeFeriado Then
    
    
        ' Busco todos los Feriados
        StrSql = "SELECT * FROM feriado "
        StrSql = StrSql & " WHERE ferifecha >= " & ConvFecha(Desde)
        StrSql = StrSql & " AND ferifecha < " & ConvFecha(hasta)
                 
        OpenRecordset StrSql, rs_feriados
        
        Do While Not rs_feriados.EOF
            
            If habiles(Weekday(rs_feriados!ferifecha)) And objFeriado.Feriado(rs_feriados!ferifecha, Ternro, False) Then
                DiasH = DiasH - 1
            End If
            
            ' Siguiente Feriado
            rs_feriados.MoveNext
        Loop

    End If
    ' cierro todo y libero
    If rs.State = adStateOpen Then rs.Close
    If rs_feriados.State = adStateOpen Then rs_feriados.Close
        
    Set rs_feriados = Nothing
    Set rs = Nothing

End Sub

Public Function Antiguedad(ByRef Dia As Integer, ByRef mes As Integer, ByRef Anio As Integer, ByRef DiasHabiles As Integer) As Integer
' -----------------------------------------------------------------------------------
' Descripcion: Antigued.p. Calcula la antiguedad al dia de hoy de un empleado en :
'               dias hábiles(si es menor que un año) o en dias, meses y años en caso contrario.
'               Retorna 0 si no hubo error y <> 0 en caso contrario
' Autor: FGZ
' Fecha: 31/07/2003
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
Dim fecalta As Date
Dim fecbaja As Date
Dim seguir As Date
Dim q As Integer
Dim NombreCampo As String
Dim rs_Fases As New ADODB.Recordset

DiasHabiles = 0

StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro
OpenRecordset StrSql, rs_Fases

If rs_Fases.EOF Then
    ' ERROR. El empleado no tiene fecha de alta en fases
    Antiguedad = 1
    Exit Function
Else
        fecalta = rs_Fases!altfec
        ' verificar si se trata de un registro completo(alta/baja) o solo de un alta
        If CBool(rs_Fases!estado) Then
            fecbaja = Date  ' solo es un alta ==> tomar el Today (Date)
        Else
            fecbaja = rs_Fases!bajfec   'se trata de un registro completo
        End If
        
        Call Dif_Fechas(fecalta, fecbaja, aux1, aux2, aux3)
'        dia = dia + aux1
'        Mes = Mes + aux2 + Int(dia / 30)
'        Anio = Anio + aux3 + Int(Mes / 12)
'        dia = dia Mod 30
'        Mes = Mes Mod 12
        
' ----------
        aux1 = aux1 Mod 30
        aux2 = aux2 Mod 12
        Dia = Dia + aux1
        mes = mes + aux2 + Int(Dia / 30)
        
        If Int(Dia / 30) >= 1 Then
            Dia = Dia Mod 30
        End If
        If aux1 < 365 Then
            aux3 = 0
        End If
        Anio = Anio + aux3 + Int(mes / 12)
        If Int(mes / 12) >= 1 Then
            mes = mes Mod 12
        End If
' ----------
        If Anio = 0 Then
            Call DiasTrab(fecalta, fecbaja, aux1)
            DiasHabiles = DiasHabiles + aux1
        End If
        Antiguedad = 0
End If

If Anio <> 0 Then
    DiasHabiles = 0
End If

' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Function


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
    valor = 0
   
    ' Obtener los parametros de la Busqueda
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    OpenRecordset StrSql, param_cur
    
    If Not param_cur.EOF Then
        TipoEstr = param_cur!auxint1
        TipoFecha = param_cur!auxint2
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
    
'    StrSql = "SELECT * FROM fases WHERE fases.empleado = " & Ternro & " AND fases.vacaciones = -1"
'    OpenRecordset StrSql, objRs
'
'    If Not objRs.EOF Then
'        objRs.MoveLast
'
'        If objRs!altfec > fecha_desde Then
'            Aux_Fecha = objRs!altfec
'        Else
'            Aux_Fecha = fecha_desde
'        End If
'    End If

    'tocado por maxi 02/02/2006
    Aux_Fecha = fecha_hasta


'antes
'Aux_Fecha = fecha_desde
Flog.writeline
Flog.writeline "Busco estructura a fecha " & Aux_Fecha
If Not IsNull(Aux_Fecha) Then
    ' Busco de estructura
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & Ternro & " AND " & _
                 " tenro =" & TipoEstr & " AND " & _
                 " (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND " & _
                 " ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
    
        If Not rs_Estructura.EOF Then
            valor = rs_Estructura!estrnro
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




Public Function Minimo(ByVal X, ByVal Y)
    
    If X <= Y Then
        Minimo = X
    Else
        Minimo = Y
    End If
    
End Function


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
Public Sub bus_Antiguedad_G(ByVal TipoAnt As String, ByVal fechaFin As Date, ByRef Dia As Long, ByRef mes As Long, ByRef Anio As Long, ByRef DiasHabiles As Integer, Optional ByVal inicioMes As Boolean = False)
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

Dia = 0
mes = 0
Anio = 0

'StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
'         " AND " & NombreCampo & " = -1 "
'OpenRecordset StrSql, rs_Fases

' FGZ -27/01/2004
StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro & _
         " AND " & NombreCampo & " = -1 " & _
         " AND not altfec is null " & _
         " AND not (bajfec is null AND estado = 0)" & _
         " AND altfec <= " & ConvFecha(fechaFin)
OpenRecordset StrSql, rs_Fases
If rs_Fases.EOF Then
    Flog.writeline "No hay fases. SQL :" & StrSql
End If

Do While Not rs_Fases.EOF
'    If (EsNulo(rs_Fases!altfec)) Or (EsNulo(rs_Fases!bajfec) And rs_Fases!estado = 0) Or (rs_Fases!altfec >= FechaFin) Then
'        GoTo siguiente
'    Else
        fecalta = rs_Fases!altfec
'    End If
    If inicioMes Then
       If Day(C_Date(fecalta)) <= 15 Then
          fecalta = C_Date("01" & "/" & Month(C_Date(fecalta)) & "/" & Year(C_Date(fecalta)))
       Else
         If Month(C_Date(fecalta)) = 12 Then
            fecalta = C_Date("01/01/" & (Year(C_Date(fecalta)) + 1))
         Else
            fecalta = C_Date("01" & "/" & (Month(C_Date(fecalta)) + 1) & "/" & Year(C_Date(fecalta)))
         End If
       End If
    End If
    
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If rs_Fases!estado Then
        fecbaja = fechaFin ' solo es un alta, tomar el fecha-fin
    ElseIf rs_Fases!bajfec <= fechaFin Then
        fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
    Else
        fecbaja = fechaFin ' hasta la fecha ingresada
    End If
    
'    Call Dif_Fechas(fecalta, fecbaja, aux1, aux2, aux3)
'    Dia = Dia + aux1
'    Mes = Mes + aux2 + Int(Dia / 30)
'    Anio = Anio + aux3 + Int(Mes / 12)
'    Dia = Dia Mod 30
'    Mes = Mes Mod 12
        
    Flog.writeline Espacios(Tabulador * 4) & "fase de " & fecalta & " a " & fecbaja
        
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    If rs_Fases.RecordCount = 1 Then
        Dia = Dia + aux1
        mes = mes + aux2 '+ Int(dia / 30)
        Anio = Anio + aux3 + Int(mes / 12)
        Dia = Dia Mod 30
        mes = mes Mod 12
    Else
        Dia = Dia + aux1
        mes = mes + aux2 '+ Int(dia / 30)
        Anio = Anio + aux3 + Int(mes / 12)
        Dia = Dia Mod 30
        mes = mes Mod 12
    End If
        
    If Anio = 0 Then
       Call DiasTrab(fecalta, fecbaja, aux1)
       DiasHabiles = DiasHabiles + aux1
    End If
    
siguiente:
    rs_Fases.MoveNext
Loop

If Anio <> 0 Then
    DiasHabiles = 0
End If


' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub


Public Sub bus_Antiguedad_U(ByVal TipoAnt As String, ByVal fechaInicio As Date, ByVal fechaFin As Date, ByRef Dia As Long, ByRef mes As Long, ByRef Anio As Long, ByRef DiasHabiles As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad entre 2 fecha para un empleado en :
'              dias hábiles(si es menor que un año) o en dias, meses y años en caso contrario.
'              antiguedad.p
' Autor      : Margiotta, Emanuel
' Fecha      : 11/08/2010
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

'EAM- Obtiene la fase del empleado
Flog.writeline "Fecha desde " & fechaInicio & " Fecha hasta " & fechaFin
StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro
StrSql = StrSql & " AND " & NombreCampo & " = -1 "
StrSql = StrSql & " AND ((fases.altfec <= " & ConvFecha(fechaFin) & " AND fases.altfec >= " & ConvFecha(fechaInicio) & ")"
StrSql = StrSql & " OR (fases.bajfec <= " & ConvFecha(fechaFin) & " AND fases.bajfec >= " & ConvFecha(fechaInicio) & ")"
StrSql = StrSql & " OR (fases.altfec <= " & ConvFecha(fechaFin) & " AND fases.bajfec is NULL))"
StrSql = StrSql & " ORDER BY altfec,bajfec "
'StrSql = "SELECT * FROM fases WHERE empleado = " & Ternro & _
'         " AND " & NombreCampo & " = -1 " & _
'         " AND not altfec is null " & _
'         " AND not (bajfec is null AND estado = 0)" & _
'         " AND altfec <= " & ConvFecha(FechaFin)
OpenRecordset StrSql, rs_Fases

If rs_Fases.EOF Then
    Flog.writeline "No hay fases. SQL :" & StrSql
End If




Do While Not rs_Fases.EOF
    If (IsNull(rs_Fases!altfec)) Or (IsNull(rs_Fases!bajfec) And rs_Fases!estado = 0) Or (rs_Fases!altfec >= fechaFin) Then
        Flog.writeline "Fase no considerada"
        GoTo siguiente
    End If
    
    
    If rs_Fases!altfec > fechaInicio Then
        fecalta = rs_Fases!altfec
    End If
    
    If Not (IsNull(rs_Fases!bajfec)) Then
        If rs_Fases!bajfec < fechaFin Then
            fecbaja = rs_Fases!bajfec
        End If
    Else
        fecbaja = fechaFin
    End If
    'Verificar si se trata de un registro completo (alta/baja) o solo de un alta
'    If rs_Fases!estado Then
'        fecbaja = FechaFin ' solo es un alta, tomar el fecha-fin
'    ElseIf rs_Fases!bajfec <= FechaFin Then
'        fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
'    Else
'        fecbaja = FechaFin ' hasta la fecha ingresada
'    End If
    
    Flog.writeline "Diferencia de fechas entre " & fecalta & " y " & fecbaja
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    Dia = Dia + aux1
    mes = mes + aux2 + Int(Dia / 30)
    Anio = Anio + aux3 + Int(mes / 12)
    Dia = Dia Mod 30
    mes = mes Mod 12



    If Anio = 0 Then
        Call DiasTrab(fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
siguiente:
    rs_Fases.MoveNext
Loop

If Anio <> 0 Then
    DiasHabiles = 0
End If


' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub

'Public Function bus_Antiguedad_Col(ByVal TipoAnt As String, ByVal FechaInicio As Date, ByVal FechaFin As Date, ByRef Dia As Long, ByRef Mes As Long, ByRef anio As Long, ByRef DiasHabiles As Integer) As Long
Public Function bus_Antiguedad_Col(ByVal fechaInicio As Date, ByVal fechaFin As Date) As Long
    'En realidad busca los dias activos en el periodo
    
    Dim i_desde As Date
    Dim i_hasta As Date
    Dim i_dias As Long
    Dim rs_Fases As New ADODB.Recordset
    
    If CDate(fechaFin) > Date Then
        fechaFin = Date
    End If
    
    i_dias = 0
    
    StrSql = " SELECT * FROM fases WHERE empleado = " & Ternro
    StrSql = StrSql & " AND altfec >= (select altfec from fases where empleado = " & Ternro & " AND fasrecofec = -1) "
    StrSql = StrSql & " AND ( "
    StrSql = StrSql & "   (altfec <= " & ConvFecha(fechaInicio) & " AND (bajfec >=" & ConvFecha(fechaInicio) & " OR bajfec is null )) "
    StrSql = StrSql & "   OR (altfec>=" & ConvFecha(fechaInicio) & " AND altfec<=" & ConvFecha(fechaFin) & ") "
    StrSql = StrSql & "   OR (altfec <= " & ConvFecha(fechaFin) & " AND bajfec is null) "
    StrSql = StrSql & " ) "
    StrSql = StrSql & " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    If rs_Fases.EOF Then
        i_dias = 0
    Else
        Do Until rs_Fases.EOF
            'Si la fecha desde es menor a la del periodo tomo la del periodo como limite
            If CDate(rs_Fases("altfec")) >= CDate(fechaInicio) Then
                i_desde = CDate(rs_Fases("altfec"))
            Else
                i_desde = CDate(fechaInicio)
            End If
            'Si la fecha hasta es mayor a la del periodo tomo la del periodo como limite
            If IsNull(rs_Fases("bajfec")) Then
                i_hasta = fechaFin
            Else
                If CDate(rs_Fases("bajfec")) > CDate(fechaFin) Then
                    i_hasta = CDate(fechaFin)
                Else
                    i_hasta = CDate(rs_Fases("bajfec"))
                End If
            End If
            i_dias = i_dias + DateDiff("d", i_desde, i_hasta) + 1 'Contemplo las fechas desde y hasta inclusive
            rs_Fases.MoveNext
        Loop
    End If
    'diasActivo = i_dias
    i_dias = i_dias - LicenciaGozadas(Ternro, fechaInicio, fechaFin)
    bus_Antiguedad_Col = i_dias

    ' Cierro todo y Libero
    If rs_Fases.State = adStateOpen Then rs_Fases.Close
    Set rs_Fases = Nothing

End Function
