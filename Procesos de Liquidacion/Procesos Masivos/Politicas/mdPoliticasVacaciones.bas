Attribute VB_Name = "mdlPoliticasVac"
Option Explicit

Dim ok As Boolean


Global rsPolitica As New ADODB.Recordset
Global PoliticaOK As Boolean 'si cargo bien o no la politica llamada
'FGZ - 29/09/2004
Global rs_Periodos_Vac As New ADODB.Recordset

Global DiasProporcion As Integer
Global FactorDivision As Single


' Debe haber 1 Variable por cada tipo de parametro posible
Global st_Opcion As Integer
Global st_VentSal As String
Global st_VentEnt As String
Global st_Iteraciones As Integer
Global st_Tolerancia As String
Global st_TipoHora1 As Long
Global st_Distancia As Integer
Global st_TamañoVentana As String
Global st_TipoDia1 As Integer
Global st_CantidadDias As Integer
Global st_FactorDivision As Integer
Global st_Escala As Integer
Global st_ModeloPago As Integer
Global st_ModeloDto As Integer

Public Sub SetearParametrosPolitica(ByVal Detalle As Long, ByRef ok As Boolean)
Dim rsPolitica As New ADODB.Recordset

    ok = False
    
    StrSql = " SELECT * FROM gti_pol_det_param " & _
             " INNER JOIN gti_pol_param ON gti_pol_det_param.polparamnro = gti_pol_param.polparamnro " & _
             " WHERE detpolnro = " & Detalle & _
             " ORDER BY gti_pol_param.polparamnro"
    OpenRecordset StrSql, rsPolitica

    If Not rsPolitica.EOF Then
        ok = True
    End If

    Do While Not rsPolitica.EOF
        Select Case rsPolitica!polparamnro
        Case 1:
            st_Opcion = CInt(rsPolitica!polparamvalor)
        Case 2:
            st_VentSal = Format(rsPolitica!polparamvalor, "0000")
        Case 3:
            ' por ahora esta vacio
        Case 4:
            st_VentEnt = Format(rsPolitica!polparamvalor, "0000")
        Case 5:
            st_Iteraciones = CInt(rsPolitica!polparamvalor)
        Case 6:
            st_Tolerancia = Format(rsPolitica!polparamvalor, "0000")
        Case 7:
            st_Distancia = CInt(rsPolitica!polparamvalor)
        Case 8:
            st_TipoHora1 = CLng(rsPolitica!polparamvalor)
        Case 9:
            st_TamañoVentana = Format(rsPolitica!polparamvalor, "0000")
        Case 10:
            st_TipoDia1 = CInt(rsPolitica!polparamvalor)
        Case 11:
            st_CantidadDias = CInt(rsPolitica!polparamvalor)
        Case 12:
            st_FactorDivision = CInt(rsPolitica!polparamvalor)
        Case 13:
            st_Escala = CInt(rsPolitica!polparamvalor)
        Case 14:
            st_ModeloPago = CInt(rsPolitica!polparamvalor)
        Case 15:
            st_ModeloDto = CInt(rsPolitica!polparamvalor)
        Case Else
        
        End Select
        
        rsPolitica.MoveNext
    Loop


End Sub

Public Sub Politica(Numero As Integer)
' --------------------------------------------------------------
' Descripcion: LLamador de las politicas
' Autor: ?
' Ultima modificacion: FGZ - 28/07/2003
' --------------------------------------------------------------


Dim objRs As New ADODB.Recordset 'Como esta función es recursiva el recordset lo tengo que definir en forma local
Dim StrSql As String
Dim det As Integer
Dim Cabecera As Long
Dim Detalle As Long

    StrSql = "SELECT gti_cabpolitica.cabpolnro,gti_cabpolitica.cabpolnivel, gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro, gti_detpolitica.detpolprograma " & _
        "FROM gti_cabpolitica INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro " & _
        "WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 3 And gti_alcanpolitica.alcpolorigen = " & empternro & " AND gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 "

    OpenRecordset StrSql, objRs
    
    If objRs.EOF Then
        
        ' EPL - 07/10/2003
        StrSql = " SELECT gti_cabpolitica.cabpolnro, gti_cabpolitica.cabpolnivel,gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro,gti_detpolitica.detpolprograma,alcance_testr.alteOrden "
        StrSql = StrSql & " FROM gti_cabpolitica "
        StrSql = StrSql & " INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro "
        StrSql = StrSql & " INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro "
        StrSql = StrSql & " INNER JOIN his_estructura ON gti_alcanpolitica.alcpolorigen = his_estructura.estrnro "
        StrSql = StrSql & " INNER JOIN alcance_tEstr ON his_estructura.tenro = alcance_tEstr.tenro "
        StrSql = StrSql & " INNER JOIN empleado ON empleado.ternro = his_estructura.ternro"
        StrSql = StrSql & " WHERE gti_cabpolitica.cabpolnivel = " & Numero
        StrSql = StrSql & " And gti_alcanpolitica.alcpolnivel = 2 "
        StrSql = StrSql & " And gti_cabpolitica.cabpolestado = -1 "
        StrSql = StrSql & " And gti_alcanpolitica.alcpolestado = -1 "
        StrSql = StrSql & " And alcance_testr.tanro = 1 "
        StrSql = StrSql & " And empleado.ternro = " & empternro
'        StrSql = StrSql & " And his_estructura.htethasta IS NULL "
        StrSql = StrSql & " And (his_estructura.htetdesde <= " & ConvFecha(Fec_Fin) & ")" 'p_fecha) & ")"
        StrSql = StrSql & " And ((" & ConvFecha(Fec_Fin) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
'        StrSql = StrSql & " And ((" & ConvFecha(p_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
        StrSql = StrSql & " ORDER BY alcance_testr.AlteOrden Asc "
        
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            StrSql = " SELECT gti_cabpolitica.cabpolnro,gti_cabpolitica.cabpolnivel, gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro, gti_detpolitica.detpolprograma " & _
             " FROM gti_cabpolitica INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro " & _
             " WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 1 And gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 "

            OpenRecordset StrSql, objRs
        End If
    End If
    
    
    
    If Not objRs.EOF Then
        det = objRs!detpolprograma
        Cabecera = objRs!cabpolnro
        Detalle = objRs!detpolnro
        
        Select Case Numero
        Case 1500: 'Vacaciones de pago/dto
            Call politica1500(det, Cabecera, Detalle)
        Case 1501: 'Proporcion de dias de Vacaciones
            Call politica1501(det, Cabecera, Detalle)
        Case 1502: 'Escala
            Call politica1502(det, Cabecera, Detalle)
        Case 1503: 'Modelo de liq. pago/dto
            Call politica1503(det, Cabecera, Detalle)
        Case 1504: 'Modelo de liq. TTI
            Call politica1504(det, Cabecera, Detalle)
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


'Politica1500AdelantaDescuenta(Date, Date, True, 1)
' corresponde a vacpdo01
'
'Politica1500AdelantaDescuentaTodo(Date, Date, True, 1)
' corresponde a vacpdo04
'
'Politica1500PagayDescuenta(Date, Date, True, 1)
' corresponde a vacpdo02
'
'Politica1500NoLiquida(Date, Date, True, 1)
' corresponde a vacpdo03
'
'Politica1500PagaDescuentaTodo(Date, Date, True, 206)
'Corresponde a vacpdo05.p
'
'Politica1500v_6(Date, Date, True, 1)
'Corresponde a vacpdo06.p

Private Sub politica1501(ByVal subn As Long, ByVal Cabecera As Long, ByVal Detalle As Long)

    Call SetearParametrosPolitica(Detalle, ok)
    DiasProporcion = st_CantidadDias
    FactorDivision = st_FactorDivision
    
End Sub

Private Sub politica1502(ByVal subn As Long, ByVal Cabecera As Long, ByVal Detalle As Long)

    Call SetearParametrosPolitica(Detalle, ok)
    If ok Then
        NroGrilla = st_Escala
        PoliticaOK = True
    Else
        PoliticaOK = False
    End If
    
End Sub

Private Sub politica1503(ByVal subn As Long, ByVal Cabecera As Long, ByVal Detalle As Long)

    Call SetearParametrosPolitica(Detalle, ok)
    If ok Then
        TipDiaPago = st_ModeloPago
        TipDiaDescuento = st_ModeloDto
        PoliticaOK = True
    Else
        PoliticaOK = False
    End If
    
End Sub

Private Sub politica1504(ByVal subn As Long, ByVal Cabecera As Long, ByVal Detalle As Long)

    PoliticaOK = True
End Sub


Private Sub politica1500(ByVal subn As Long, ByVal Cabecera As Long, ByVal Detalle As Long)
Dim Op As Integer

    'Call SetearParametrosPolitica(Detalle, ok)
    'Op = st_Opcion

    Op = subn
    Select Case Op
        Case 1:
            If GeneraPorLicencia Then
                Call Politica1500AdelantaDescuenta
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Do While Not rs_Periodos_Vac.EOF
                    NroVac = rs_Periodos_Vac!vacnro
                    Call Politica1500_v2AdelantaDescuenta
                    
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 2:
            If GeneraPorLicencia Then
                Call Politica1500PagayDescuenta
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Do While Not rs_Periodos_Vac.EOF
                    NroVac = rs_Periodos_Vac!vacnro
                    Call Politica1500_v2AdelantaDescuenta
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 3:
            If GeneraPorLicencia Then
                Call Politica1500NoLiquida
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Do While Not rs_Periodos_Vac.EOF
                    NroVac = rs_Periodos_Vac!vacnro
                    Call Politica1500_v2AdelantaDescuenta
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 4:
            If GeneraPorLicencia Then
                Call Politica1500AdelantaDescuentaTodo
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Do While Not rs_Periodos_Vac.EOF
                    NroVac = rs_Periodos_Vac!vacnro
                    Call Politica1500_v2AdelantaDescuenta
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 5:
            If GeneraPorLicencia Then
                Call Politica1500PagaDescuentaTodo
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Do While Not rs_Periodos_Vac.EOF
                    NroVac = rs_Periodos_Vac!vacnro
                    Call Politica1500_v2AdelantaDescuenta
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 6:
            If GeneraPorLicencia Then
                Call Politica1500v_6
            Else
                Call Politica1500_v2AdelantaDescuenta
            End If
        Case 7:
            If GeneraPorLicencia Then
                Call Politica1500v_TMK
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Do While Not rs_Periodos_Vac.EOF
                    NroVac = rs_Periodos_Vac!vacnro
                    Call Politica1500_v2AdelantaDescuenta
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 8:
            If GeneraPorLicencia Then
                Call Politica1500PagayDescuenta_Mes_a_Mes
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Do While Not rs_Periodos_Vac.EOF
                    NroVac = rs_Periodos_Vac!vacnro
                    Call Politica1500_v2AdelantaDescuenta
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
    End Select
End Sub



Public Function AFecha(m As Integer, d As Integer, a As Integer) As Date
' Reemplaza a la función Date de Progress
'ultimo-mes  = DATE (mes-afecta,30,ano-afecta)
Dim auxi
  
  auxi = Str(m) & "/" & Str(d) & "/" & Str(a)
  AFecha = Format(auxi, "mm/dd/yyyy")

End Function


Public Function FechaLiqVacaciones(anio_vacac As Integer, Ternro As Long) As Date
'Determina a que fecha se liquidan los dias correspondientes de vacaciones
'Parametros:
'   anio_vacac es el año al que estoy liquidando
'   TerNro el código de tercero del empleado.
'   Devuelve la fecha a la que hay que liquidar

Dim rs As New ADODB.Recordset
Dim StrSql As String
Dim mes As Integer
Dim dia As Integer

'Busco en que pais estoy
StrSql = "SELECT * FROM pais WHERE paisnro = 36"
rs.Open StrSql, objConn

If rs!paisdef = 0 Then
' si no estoy en Chile liquido al 31/12
    FechaLiqVacaciones = AFecha(12, 31, anio_vacac)
Else
' estoy en Chile. Liquido a la fecha del aniversario del empleado.
    rs.Close
    StrSql = "SELECT terfecnac FROM tercero WHERE ternro = " & Ternro
    rs.Open StrSql, objConn
    
    mes = Month(rs!terfecnac)
    dia = Day(rs!terfecnac)
    FechaLiqVacaciones = AFecha(mes, dia, anio_vacac)
End If

rs.Close
Set rs = Nothing

End Function


Public Sub Politica1500AdelantaDescuenta()
'-----------------------------------------------------------------------
' paga adelantado todo y descuenta por mes lo que corresponde generar todos los dias de pago
'-----------------------------------------------------------------------

Dim rs As New ADODB.Recordset
Dim mes_inicio As Integer
Dim ano_inicio As Integer
Dim mes_fin    As Integer
Dim ano_fin    As Integer
Dim fin_licencia  As Boolean
Dim mes_afecta    As Integer
Dim ano_afecta    As Integer
Dim primero_mes   As Date
Dim ultimo_mes    As Date
Dim Dias_Afecta   As Integer
Dim dias_pendientes As Integer
Dim dias_restantes As Integer

Dim dias_ya_tomados As Integer
Dim fecha_limite    As Date
Dim anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

Call Politica(1503)
If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503"
        Exit Sub
End If


'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn
' Si no hay licencia me voy
If rs.EOF Then
    Flog.writeline "Licencia inexistente.", vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    mes_inicio = Month(rs!elfechadesde)
    ano_inicio = Year(rs!elfechadesde)
    mes_fin = Month(rs!elfechahasta)
    ano_fin = Year(rs!elfechahasta)
End If
rs.Close

'/* VERIFICAR el reproceso y manejar la depuraci¢n */
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & "AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        StrSql = "SELECT empleg FROM empleado WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Legajo = rs!empleg
        rs.Close
        StrSql = "SELECT terape,ternom FROM tercero WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Nombre = rs!terape & ", " & rs!ternom
        'MsgBox "No se puede Reprocesar una Licencia con pagos y/o descuentos Liquidados. " & Legajo & " " & Nombre, vbCritical
        rs.Close
    Else
    '/* DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA */
        rs.Close
        'StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
StrSql = "SELECT * FROM tipdia WHERE tdnro = " & TipoLicencia
rs.Open StrSql, objConn
If rs.EOF Then
    Flog.writeline " Tipo de licencia (" & TipoLicencia & ") inexistente."
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.Close
End If

StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
OpenRecordset StrSql, rs

'jornal = IIf(rs!folinro = 2, True, False)
rs.Close

'/* POLITICA  - ANALISIS */

StrSql = "SELECT * FROM emp_lic "
If TipoLicencia = 2 Then
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
End If
StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If TipoLicencia = 2 Then
    NroVac = rs!vacnro
End If
Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1

'Terminar
'RUN generar_pago(mes_inicio, ano_inicio, dias_afecta, 3)
'Call generar_pago(mes_inicio, ano_inicio, Dias_Afecta, 3, Jornal, nrolicencia)
'Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
'fgz - 09/01/2003

'Call Generar_PagoDescuento(primer_mes, primer_ano, TipDiaPago, Dias_Afecta, 3) 'Pago
Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'Pago

'IF RETURN-VALUE <> ""
'          THEN UNDO main, RETURN.
          
'/* tantos descuentos como meses afecte */
fin_licencia = False
mes_afecta = mes_inicio
ano_afecta = ano_inicio
dias_pendientes = 0
dias_restantes = Dias_Afecta

'/* determinar los dias que afecta para el primer mes de decuento */
'/* Genera 30 dias, para todos los meses */
anio_bisiesto = EsBisiesto(ano_afecta)

If mes_afecta = 2 Then
' Date en progress es una función
' la sintaxis es DATE(month,day,year)

    primero_mes = AFecha(mes_afecta, 1, ano_afecta)
    ultimo_mes = IIf(anio_bisiesto = True, AFecha(mes_afecta + 1, 1, ano_afecta), AFecha(mes_afecta + 1, 2, ano_afecta))
Else
    primero_mes = AFecha(mes_afecta, 1, ano_afecta)
    'ultimo_mes = AFecha(mes_afecta, 30, ano_afecta)
    If mes_afecta = 12 Then
        ultimo_mes = AFecha(mes_afecta, 31, ano_afecta)
    Else
        ultimo_mes = AFecha(mes_afecta + 1, 1, ano_afecta) - 1
    End If
End If

If (rs!elfechahasta <= ultimo_mes) Then
'/* termina en el mes */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
'/* continua en el mes siguiente */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, ultimo_mes) + 1
End If

Do While Not fin_licencia
    dias_restantes = dias_restantes - Dias_Afecta
    'Revisar
    'RUN generar_descuento(mes_afecta, ano_afecta, dias_afecta, 4)
    'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
    Call Generar_PagoDescuento(mes_afecta, ano_afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
    'IF RETURN-VALUE <> ""
    '         THEN UNDO main, RETURN.
    
    '/* determinar a que continua en el proximo mes */
    If (mes_afecta = 12) Then
        mes_afecta = 1
        ano_afecta = ano_afecta + 1
    Else
        mes_afecta = mes_afecta + 1
    End If
    If (mes_afecta > mes_fin) Or (mes_afecta = 1 And mes_fin = 12) Then
        fin_licencia = True
    End If
    
    '/* determinar los d¡as que afecta */

    anio_bisiesto = EsBisiesto(ano_afecta)

    If mes_afecta = 2 Then
        primero_mes = AFecha(mes_afecta, 1, ano_afecta)
        ultimo_mes = IIf(anio_bisiesto = True, AFecha(mes_afecta + 1, 1, ano_afecta), AFecha(mes_afecta + 1, 2, ano_afecta))
    Else
        primero_mes = AFecha(mes_afecta, 1, ano_afecta)
        'ultimo_mes = AFecha(mes_afecta, 30, ano_afecta)
        If mes_afecta = 12 Then
            ultimo_mes = AFecha(mes_afecta, 31, ano_afecta)
        Else
            ultimo_mes = AFecha(mes_afecta + 1, 1, ano_afecta) - 1
        End If
    End If

    If (rs!elfechahasta <= ultimo_mes) Then
    '/* termina en el mes */
        If (dias_restantes < DateDiff("d", primero_mes, ultimo_mes)) Then
            Dias_Afecta = dias_restantes
        Else
            Dias_Afecta = DateDiff("d", primero_mes, rs!elfechahasta) + 1
        End If
    Else
    '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", primero_mes, ultimo_mes) + 1
    End If
Loop

rs.Close
Set rs = Nothing

End Sub


Private Sub generar_pago(ByVal mes_aplicar As Integer, ano_aplicar As Integer, Dias_Afecta As Integer, anti_vac As Integer, Jornal As Boolean, nrolicencia As Long)

Dim rs As New Recordset
Dim StrSql As String
Dim TipoDia As Integer

'Abrir tipdia
StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
OpenRecordset StrSql, rs
TipoDia = IIf(Jornal, rs!tdinteger4, rs!tdinteger1)
rs.Close

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM periodo WHERE pliqanio= " & ano_aplicar & _
" AND pliqmes = " & mes_aplicar
OpenRecordset StrSql, rs

If rs.EOF Then
'   MsgBox "Periodo de liquidación inexistente para generar el Pago de la Lic.Vacaciones del:  " & ano_aplicar & " - " & mes_aplicar, vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
End If

StrSql = "INSERT INTO vacpagdesc(emp_licnro,tprocnro,pago_dto,pliqnro,cantdias,manual)" & _
" VALUES(" & nrolicencia & "," & TipoDia & "," & anti_vac & "," & rs!PliqNro & "," & Dias_Afecta & "," & "0)"
' Cierro el recordset de la liquidacion
rs.Close

'Ejecuto la consulta
objConn.Execute StrSql, , adExecuteNoRecords

'Libero
Set rs = Nothing

End Sub

 
Function EsBisiesto(anio As Integer) As Boolean
If (anio Mod 4) = 0 Then
    If (((anio Mod 100) <> 0) And ((anio Mod 400) = 0)) Or _
        (((anio Mod 100) = 0) And ((anio Mod 400) = 0)) Or _
        (((anio Mod 100) <> 0) And ((anio Mod 400) <> 0)) Then
           EsBisiesto = True
       Else
           EsBisiesto = False
    End If
 Else
    EsBisiesto = False
End If

End Function


Private Sub generar_descuento(mes_aplicar As Integer, ano_aplicar As Integer, Dias_Afecta As Integer, anti_vac As Integer, Jornal As Boolean, nro_lic As Long)

Dim rs As New Recordset
Dim StrSql As String
Dim TipoDia As Integer

'Abrir tipdia
StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
OpenRecordset StrSql, rs
TipoDia = IIf(Jornal, rs!tdinteger5, rs!tdinteger2)
rs.Close

StrSql = "SELECT * FROM periodo WHERE pliqanio= " & ano_aplicar & _
" AND pliqmes = " & mes_aplicar
OpenRecordset StrSql, rs
If rs.EOF Then
'    MsgBox "Periodo de liquidación inexistente para generar el Descuento de la Lic.Vacaciones del:  " & ano_aplicar & " - " & mes_aplicar, vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
End If

StrSql = "INSERT INTO vacpagdesc(emp_licnro,tprocnro,pago_dto,pliqnro,cantdias,manual)" & _
" VALUES(" & nro_lic & "," & TipoDia & "," & anti_vac & "," & rs!PliqNro & "," & Dias_Afecta & ",0)"
' Cierro el recordset de la liquidacion
rs.Close



'Ejecuto la consulta
objConn.Execute StrSql, , adExecuteNoRecords

'Libero
Set rs = Nothing
                       
End Sub


Public Sub Politica1500PagayDescuenta()
' corresponde a vacpdo02

'/* ***************************************************************
'PAGA Y DESCUENTA POR MES SIN TOPE DE 30 DIAS FIJOS.
'
'*************************************************************** */

Dim rs As New Recordset
Dim StrSql As String

Dim mes_inicio As Integer
Dim ano_inicio As Integer
Dim mes_fin    As Integer
Dim ano_fin    As Integer
Dim fin_licencia  As Boolean
Dim mes_afecta    As Integer
Dim ano_afecta    As Integer
Dim primero_mes   As Date
Dim ultimo_mes    As Date
Dim Dias_Afecta   As Integer
Dim dias_pendientes As Integer
Dim dias_restantes As Integer

'DEF BUFFER buf_lic FOR emp_lic.
Dim dias_ya_tomados As Integer
Dim fecha_limite    As Date
Dim anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

'/**************************************************************************************************/

Call Politica(1503)
If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503"
        Exit Sub
End If


'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn
' Si no hay licencia me voy
If rs.EOF Then
    'MsgBox "Licencia inexistente.", vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    'NroTer = emp_lic.empleado
    NroTer = rs!Empleado
    mes_inicio = Month(rs!elfechadesde)
    ano_inicio = Year(rs!elfechadesde)
    mes_fin = Month(rs!elfechahasta)
    ano_fin = Year(rs!elfechahasta)
End If
rs.Close

'/* VERIFICAR el reproceso y manejar la depuraci¢n */
If Not Reproceso Then
    '/* si no es reproceso y existe el desglose de pago/descuento, salir */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Exit Sub
    End If
    rs.Close
Else
    '/*  verificar si no tiene pagos/descuentos ya procesados, sino depurarlos */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & "AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        StrSql = "SELECT empleg FROM empleado WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Legajo = rs!empleg
        rs.Close
        StrSql = "SELECT terape,ternom FROM tercero WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Nombre = rs!terape & ", " & rs!ternom
        'MsgBox "No se puede Reprocesar una Licencia con pagos y/o descuentos Liquidados. " & Legajo & " " & Nombre, vbCritical
        rs.Close
    Else
    '/* DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA */
        rs.Close
        StrSql = "DELETE vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


StrSql = "SELECT * FROM tipdia WHERE tdnro = " & TipoLicencia
rs.Open StrSql, objConn
If rs.EOF Then
    'MsgBox " Tipo de dia de Vacaciones (2) inexistente."
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.Close
End If

StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
rs.Open StrSql, objConn
'jornal = IIf(rs!folinro = 2, True, False)
rs.Close

'/* POLITICA  - ANALISIS */

fin_licencia = False
mes_afecta = mes_inicio
ano_afecta = ano_inicio
dias_pendientes = 0

'/* determinar los dias que afecta para el primer mes de descuento */
primero_mes = AFecha(mes_afecta, 1, ano_afecta)
ultimo_mes = IIf(mes_afecta = 12, DateAdd("d", -1, AFecha(1, 1, ano_afecta + 1)), _
DateAdd("d", -1, AFecha(mes_afecta + 1, 1, ano_afecta)))

StrSql = "SELECT * FROM emp_lic "
If TipoLicencia = 2 Then
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
End If
StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If TipoLicencia = 2 Then
    NroVac = rs!vacnro
End If

'StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'rs.Open StrSql, objConn

If (rs!elfechahasta <= ultimo_mes) Then
'/* termina en el mes */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
'/* continua en el mes siguiente */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, ultimo_mes) + 1
End If
       
Do While Not fin_licencia
    'RUN generar_pago  (mes_afecta,ano_afecta,dias_afecta,3)
    'Call generar_pago(mes_inicio, ano_inicio, Dias_Afecta, 3, Jornal, nrolicencia)
    Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
    'RUN generar_descuento  (mes_afecta,ano_afecta,dias_afecta,4)
    'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
    Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaDescuento, Dias_Afecta, 4) 'Pago
    
    If (mes_afecta = 12) Then
        mes_afecta = 1
        ano_afecta = ano_afecta + 1
    Else
        mes_afecta = mes_afecta + 1
    End If
    
    If (mes_afecta > mes_fin) Or (mes_afecta = 1 And mes_fin = 12) Then
        fin_licencia = True
    End If

    '/* determinar los d­as que afecta */
    primero_mes = AFecha(mes_afecta, 1, ano_afecta)
    ultimo_mes = IIf(mes_afecta = 12, DateAdd("d", -1, AFecha(1, 1, ano_afecta + 1)), _
    DateAdd("d", -1, AFecha(mes_afecta + 1, 1, ano_afecta)))
    
    If (rs!elfechahasta <= ultimo_mes) Then
    '/* termina en el mes */
        Dias_Afecta = DateDiff("d", primero_mes, rs!elfechahasta) + 1
    Else
    '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", primero_mes, ultimo_mes) + 1
    End If

Loop

End Sub


Public Sub Politica1500NoLiquida()
' corresponde a vacpdo03

'/* ***************************************************************
' NO LIQUIDA:Pago adelantado entero: mes anterior al inicio de las vacaciones.
' NO LIQUIDA:Descuento adelantado entero: mes de inicio de las vacaciones.
' PAGO adelantado entero: mes de inicio de las vacaciones.
' DESCUENTO x mes tomado con tope de 30 dias mensuales fijos: mes de inicio de las vacaciones.
'
' Nota: Para que liquide tocar el licdes.p y licpag.p que esta en /par.
'
'*************************************************************** */

Dim rs As New Recordset
Dim StrSql As String

Dim mes_inicio As Integer
Dim ano_inicio As Integer
Dim mes_fin    As Integer
Dim ano_fin    As Integer
Dim fin_licencia  As Boolean
Dim mes_afecta    As Integer
Dim ano_afecta    As Integer
Dim primero_mes   As Date
Dim ultimo_mes    As Date
Dim Dias_Afecta   As Integer
Dim dias_pendientes As Integer
Dim dias_restantes As Integer

'DEF BUFFER buf_lic FOR emp_lic.
Dim dias_ya_tomados As Integer
Dim fecha_limite    As Date
Dim anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

'/**************************************************************************************************/

Call Politica(1503)
If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503"
        Exit Sub
End If


'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn
' Si no hay licencia me voy
If rs.EOF Then
    'MsgBox "Licencia inexistente.", vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    'NroTer = emp_lic.empleado
    NroTer = rs!Empleado
    mes_inicio = Month(rs!elfechadesde)
    ano_inicio = Year(rs!elfechadesde)
    mes_fin = Month(rs!elfechahasta)
    ano_fin = Year(rs!elfechahasta)
End If
rs.Close

'/* VERIFICAR el reproceso y manejar la depuraci¢n */
If Not Reproceso Then
    '/* si no es reproceso y existe el desglose de pago/descuento, salir */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Exit Sub
    End If
    rs.Close
Else
    '/*  verificar si no tiene pagos/descuentos ya procesados, sino depurarlos */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & "AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        StrSql = "SELECT empleg FROM empleado WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Legajo = rs!empleg
        rs.Close
        StrSql = "SELECT terape,ternom FROM tercero WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Nombre = rs!terape & ", " & rs!ternom
        'MsgBox "No se puede Reprocesar una Licencia con pagos y/o descuentos Liquidados. " & Legajo & " " & Nombre, vbCritical
        rs.Close
    Else
    '/* DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA */
        rs.Close
        StrSql = "DELETE vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
rs.Open StrSql, objConn
If rs.EOF Then
    'MsgBox " Tipo de dia de Vacaciones (2) inexistente."
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.Close
End If

StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
rs.Open StrSql, objConn
'jornal = IIf(rs!folinro = 2, True, False)
rs.Close

'/* POLITICA  - ANALISIS */
StrSql = "SELECT * FROM emp_lic "
If TipoLicencia = 2 Then
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
End If
StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If TipoLicencia = 2 Then
    NroVac = rs!vacnro
End If

'StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'rs.Open StrSql, objConn

Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1

'/* generar pago de adelanto */
If (mes_inicio = 1) Then
    'RUN generar_pago (12,(ano_inicio - 1),dias_afecta,1)
    'Call generar_pago(12, (ano_inicio - 1), Dias_Afecta, 1, Jornal, nrolicencia)
    Call Generar_PagoDescuento(12, (ano_inicio - 1), TipDiaPago, Dias_Afecta, 1) 'Pago
Else
    'RUN generar_pago ((mes_inicio - 1),ano_inicio,dias_afecta,1)
    'Call generar_pago((mes_inicio - 1), ano_inicio, Dias_Afecta, 1, Jornal, nrolicencia)
    Call Generar_PagoDescuento((mes_inicio - 1), ano_inicio, TipDiaPago, Dias_Afecta, 1) 'PAgo
End If

'/* GENERAR DESCUENTO DE ANTICIPO */
'RUN generar_descuento (mes_inicio,ano_inicio,dias_afecta,2)
'Call generar_descuento(mes_inicio, ano_inicio, Dias_Afecta, 2, Jornal, nrolicencia)
Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaDescuento, Dias_Afecta, 2) 'PAgo
'       IF RETURN-VALUE <> ""
'          THEN UNDO main, RETURN.
          
'/* GENERAR EL PAGO DE VACACIONES */
'RUN generar_pago (mes_inicio,ano_inicio,dias_afecta,3)
'Call generar_pago(mes_inicio, ano_inicio, Dias_Afecta, 3, Jornal, nrolicencia)
Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'PAgo

'       IF RETURN-VALUE <> ""
'          THEN UNDO main, RETURN.

'/* GENERAR TANTOS DESCUENTOS SEGUN CORRESPONDA AL MES */
fin_licencia = False
mes_afecta = mes_inicio
ano_afecta = ano_inicio

'/* determinar los d­as que afecta para el primer mes de descuento */
primero_mes = AFecha(mes_afecta, 1, ano_afecta)
ultimo_mes = IIf(mes_afecta = 12, DateAdd("d", -1, AFecha(1, 1, ano_afecta + 1)), _
DateAdd("d", -1, AFecha(mes_afecta + 1, 1, ano_afecta)))

'Revisar comparacion de fechas
If (rs!elfechahasta <= ultimo_mes) Then
'/* termina en el mes */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
'/* continua en el mes siguiente */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, ultimo_mes) + 1
End If

Do While Not fin_licencia
    If Dias_Afecta > 30 Then
        dias_pendientes = Dias_Afecta - 30
        Dias_Afecta = 30
    End If

    'RUN generar_descuento  (mes_afecta,ano_afecta,dias_afecta,4)
    'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
    Call Generar_PagoDescuento((mes_inicio - 1), ano_inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
    'IF RETURN-VALUE <> ""
        'THEN UNDO main, RETURN.
    If (mes_afecta = 12) Then
        mes_afecta = 1
        ano_afecta = ano_afecta + 1
    Else
        mes_afecta = mes_afecta + 1
    End If
    If (mes_afecta > mes_fin) Or (mes_afecta = 1 And mes_fin = 12) Then
        fin_licencia = True
    End If

    '/* determinar los d¡as que afecta */
    primero_mes = AFecha(mes_afecta, 1, ano_afecta)
    ultimo_mes = IIf(mes_afecta = 12, DateAdd("d", -1, AFecha(1, 1, ano_afecta + 1)), _
    DateAdd("d", -1, AFecha(mes_afecta + 1, 1, ano_afecta)))
       
    If (rs!elfechahasta <= ultimo_mes) Then
    '/* termina en el mes */
        Dias_Afecta = DateDiff("d", primero_mes, rs!elfechahasta) + 1 + dias_pendientes
    Else
    '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", primero_mes, ultimo_mes) + 1 + dias_pendientes
    End If

Loop

rs.Close
Set rs = Nothing

End Sub


Public Sub Politica1500AdelantaDescuentaTodo()
' corresponde a vacpdo04

'/* ***************************************************************
' paga adelantado, descuenta adelantado todo
'
'*************************************************************** */

Dim rs As New Recordset
Dim StrSql As String

Dim mes_inicio As Integer
Dim ano_inicio As Integer
Dim mes_fin    As Integer
Dim ano_fin    As Integer
Dim fin_licencia  As Boolean
Dim mes_afecta    As Integer
Dim ano_afecta    As Integer
Dim primero_mes   As Date
Dim ultimo_mes    As Date
Dim Dias_Afecta   As Integer
Dim dias_pendientes As Integer
Dim dias_restantes As Integer

'DEF BUFFER buf_lic FOR emp_lic.
Dim dias_ya_tomados As Integer
Dim fecha_limite    As Date
Dim anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

'/**************************************************************************************************/
Call Politica(1503)
If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503"
        Exit Sub
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn
' Si no hay licencia me voy
If rs.EOF Then
    'MsgBox "Licencia inexistente.", vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    'NroTer = emp_lic.empleado
    NroTer = rs!Empleado
    mes_inicio = Month(rs!elfechadesde)
    ano_inicio = Year(rs!elfechadesde)
    mes_fin = Month(rs!elfechahasta)
    ano_fin = Year(rs!elfechahasta)
End If
rs.Close

'/* VERIFICAR el reproceso y manejar la depuraci¢n */
If Not Reproceso Then
    '/* si no es reproceso y existe el desglose de pago/descuento, salir */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Exit Sub
    End If
    rs.Close
Else
    '/*  verificar si no tiene pagos/descuentos ya procesados, sino depurarlos */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & "AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        StrSql = "SELECT empleg FROM empleado WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Legajo = rs!empleg
        rs.Close
        StrSql = "SELECT terape,ternom FROM tercero WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Nombre = rs!terape & ", " & rs!ternom
        'MsgBox "No se puede Reprocesar una Licencia con pagos y/o descuentos Liquidados. " & Legajo & " " & Nombre, vbCritical
        rs.Close
    Else
    '/* DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA */
        rs.Close
        StrSql = "DELETE vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
rs.Open StrSql, objConn
If rs.EOF Then
    'MsgBox " Tipo de dia de Vacaciones (2) inexistente."
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.Close
End If

StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
rs.Open StrSql, objConn
'jornal = IIf(rs!folinro = 2, True, False)
rs.Close

'Revisar
'Politica = tipdia.tdformapagdto
                                
'/* POLITICA  - sin ANALISIS */

'StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'rs.Open StrSql, objConn
StrSql = "SELECT * FROM emp_lic "
If TipoLicencia = 2 Then
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
End If
StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If TipoLicencia = 2 Then
    NroVac = rs!vacnro
End If

Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1

'Terminar
'RUN generar_pago(mes_inicio, ano_inicio, dias_afecta, 3)
'Call generar_pago(mes_inicio, ano_inicio, Dias_Afecta, 3, Jornal, nrolicencia)
Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'PAgo

'IF RETURN-VALUE <> ""
'          THEN UNDO main, RETURN.
          
'RUN generar-descuento (mes-inicio,ano-inicio,dias-afecta,4).
'Call generar_descuento(mes_inicio, ano_inicio, Dias_Afecta, 4, Jornal, nrolicencia)
Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
'IF RETURN-VALUE <> ""
'         THEN UNDO main, RETURN.

rs.Close
Set rs = Nothing

End Sub



Public Sub Politica1500PagaDescuentaTodo()
'Corresponde a vacpdo05.p

'/* ***************************************************************
'  EN CUANTO SE TOMA AL MENOS UN DIA DE VACACIONES, SE LE PAGA Y DESCUENTA TODOS LOS D™AS DE
'  VACACIONES QUE LE CORRESPONDEN PARA EL A¾O
'  TODO ESTO LO HACE MIENTRAS NO PAGO NI DESCONTO NINGUN DIA DE DICHA LICENCIA
'
'*************************************************************** */

Dim rs As New Recordset
Dim StrSql As String

Dim mes_inicio As Integer
Dim ano_inicio As Integer
Dim mes_fin    As Integer
Dim ano_fin    As Integer
Dim fin_licencia  As Boolean
Dim mes_afecta    As Integer
Dim ano_afecta    As Integer
Dim primero_mes   As Date
Dim ultimo_mes    As Date
Dim Dias_Afecta   As Integer
Dim dias_pendientes As Integer
Dim dias_restantes As Integer

'DEF BUFFER buf_lic FOR emp_lic.
Dim dias_ya_tomados As Integer
Dim fecha_limite    As Date
Dim anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

Dim ya_se_pago As Boolean
Dim blv As New ADODB.Recordset
Dim bel As New ADODB.Recordset

'/**************************************************************************************************/
Call Politica(1503)
If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503"
        Exit Sub
End If


'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn
' Si no hay licencia me voy
If rs.EOF Then
    'MsgBox "Licencia inexistente.", vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    'NroTer = emp_lic.empleado
    NroTer = rs!Empleado
    mes_inicio = Month(rs!elfechadesde)
    ano_inicio = Year(rs!elfechadesde)
    mes_fin = Month(rs!elfechahasta)
    ano_fin = Year(rs!elfechahasta)
End If
rs.Close

'/* VERIFICAR el reproceso y manejar la depuraci¢n */
If Not Reproceso Then
    '/* si no es reproceso y existe el desglose de pago/descuento, salir */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Exit Sub
    End If
    rs.Close
Else
    '/*  verificar si no tiene pagos/descuentos ya procesados, sino depurarlos */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & "AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        StrSql = "SELECT empleg FROM empleado WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Legajo = rs!empleg
        rs.Close
        StrSql = "SELECT terape,ternom FROM tercero WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Nombre = rs!terape & ", " & rs!ternom
        'MsgBox "No se puede Reprocesar una Licencia con pagos y/o descuentos Liquidados. " & Legajo & " " & Nombre, vbCritical
        rs.Close
    Else
    '/* DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA */
        rs.Close
        StrSql = "DELETE vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
rs.Open StrSql, objConn
If rs.EOF Then
    'MsgBox " Tipo de dia de Vacaciones (2) inexistente."
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.Close
End If

StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
rs.Open StrSql, objConn
'jornal = IIf(rs!folinro = 2, True, False)
rs.Close

'Politica = tipdia.tdformapagdto

'/* BUSCAR EL PERIODO DE LA LICENCIA */
'       FIND lic_vacacion OF emp_lic NO-LOCK.
StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn

'       FIND vacacion OF lic_vacacion NO-LOCK.
StrSql = "SELECT * FROM vacacion WHERE vacnro= " & rs!vacnro
rs.Close
rs.Open StrSql, objConn

'       DEF BUFFER blv FOR lic_vacacion.
'       DEF BUFFER bel FOR emp_lic.
'       DEF VAR ya-se-pago AS LOG INITIAL FALSE.
ya_se_pago = False

StrSql = "SELECT * FROM lic_vacacion" & _
" INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro" & _
" WHERE emp_lic.empleado = " & NroTer & " AND vacnro= " & rs!vacnro
rs.Close
blv.Open StrSql, objConn
'       FOR EACH blv OF vacacion NO-LOCK,
'           EACH bel OF blv WHERE bel.empleado=emp_lic.empleado NO-LOCK:
'           FIND FIRST vacpagdesc OF blv NO-LOCK NO-ERROR.
'           IF AVAILABLE vacpagdesc
'              THEN DO:
'              ASSIGN ya-se-pago = TRUE.
'              LEAVE.
'              END.
'       END.
Do While Not blv.EOF
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro= " & blv!emp_licnro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        ya_se_pago = True
        rs.Close
        Exit Do
    End If
    blv.MoveNext
    rs.Close
Loop

blv.Close

If Not ya_se_pago Then
    StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
    rs.Open StrSql, objConn
       
    StrSql = "SELECT * FROM vacacion WHERE vacnro = " & rs!vacnro
    rs.Close
    
    rs.Open StrSql, objConn
          
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & rs!vacnro & _
    " AND ternro = " & NroTer
    rs.Close
    
    rs.Open StrSql, objConn
    
    NroVac = rs!vacnro
    
    'Call generar_pago(mes_inicio, ano_inicio, rs!vdiascorcant, 3, Jornal, nrolicencia)
    Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'PAgo
    'Call generar_descuento(mes_inicio, ano_inicio, rs!vdiascorcant, 4, Jornal, nrolicencia)
    Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
    rs.Close
End If

'rs.Close
Set rs = Nothing

End Sub


Public Sub Politica1500v_6()
'Corresponde a vacpdo06.p
'/* ***************************************************************
'  EN CUANTO SE TOMA AL MENOS UN DIA DE VACACIONES, SE LE PAGA Y DESCUENTA TODOS LOS D™AS DE
'  VACACIONES QUE LE CORRESPONDEN PARA EL A¾O
'  TODO ESTO LO HACE MIENTRAS NO PAGO NI DESCONTO NINGUN DIA DE DICHA LICENCIA
'
'*************************************************************** */
Dim rs As New Recordset
Dim StrSql As String

Dim mes_inicio As Integer
Dim ano_inicio As Integer
Dim mes_fin    As Integer
Dim ano_fin    As Integer
Dim fin_licencia  As Boolean
Dim mes_afecta    As Integer
Dim ano_afecta    As Integer
Dim primero_mes   As Date
Dim ultimo_mes    As Date
Dim Dias_Afecta   As Integer
Dim dias_pendientes As Integer
Dim dias_restantes As Integer

'DEF BUFFER buf_lic FOR emp_lic.
Dim dias_ya_tomados As Integer
Dim fecha_limite    As Date
Dim anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

Dim ya_se_pago As Boolean
Dim cantdias As Integer
Dim blv As New ADODB.Recordset
Dim bel As New ADODB.Recordset

'/**************************************************************************************************/
Call Politica(1503)
If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503"
        Exit Sub
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn
' Si no hay licencia me voy
If rs.EOF Then
    'MsgBox "Licencia inexistente.", vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    'NroTer = emp_lic.empleado
    NroTer = rs!Empleado
    mes_inicio = Month(rs!elfechadesde)
    ano_inicio = Year(rs!elfechadesde)
    mes_fin = Month(rs!elfechahasta)
    ano_fin = Year(rs!elfechahasta)
End If
rs.Close

'/* VERIFICAR el reproceso y manejar la depuraci¢n */
If Not Reproceso Then
    '/* si no es reproceso y existe el desglose de pago/descuento, salir */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Exit Sub
    End If
    rs.Close
Else
    '/*  verificar si no tiene pagos/descuentos ya procesados, sino depurarlos */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & "AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        StrSql = "SELECT empleg FROM empleado WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Legajo = rs!empleg
        rs.Close
        StrSql = "SELECT terape,ternom FROM tercero WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Nombre = rs!terape & ", " & rs!ternom
        'MsgBox "No se puede Reprocesar una Licencia con pagos y/o descuentos Liquidados. " & Legajo & " " & Nombre, vbCritical
        rs.Close
    Else
    '/* DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA */
        rs.Close
        StrSql = "DELETE vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
rs.Open StrSql, objConn
If rs.EOF Then
    'MsgBox " Tipo de dia de Vacaciones (2) inexistente."
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.Close
End If

StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
rs.Open StrSql, objConn
'jornal = IIf(rs!folinro = 2, True, False)
rs.Close

'Politica = tipdia.tdformapagdto

'/* BUSCAR EL PERIODO DE LA LICENCIA */
'       FIND lic_vacacion OF emp_lic NO-LOCK.
StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn

'       FIND vacacion OF lic_vacacion NO-LOCK.
StrSql = "SELECT * FROM vacacion WHERE vacnro= " & rs!vacnro
rs.Close
rs.Open StrSql, objConn

'       DEF BUFFER blv FOR lic_vacacion.
'       DEF BUFFER bel FOR emp_lic.
'       DEF VAR ya-se-pago AS LOG INITIAL FALSE.
ya_se_pago = False

'Busco todas las licencias del empleado en el periodo
StrSql = "SELECT * FROM lic_vacacion" & _
" INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro" & _
" WHERE emp_lic.empleado = " & NroTer & " AND vacnro= " & rs!vacnro
rs.Close
blv.Open StrSql, objConn
'       FOR EACH blv OF vacacion NO-LOCK,
'           EACH bel OF blv WHERE bel.empleado=emp_lic.empleado NO-LOCK:
'           FIND FIRST vacpagdesc OF blv NO-LOCK NO-ERROR.
'           IF AVAILABLE vacpagdesc
'              THEN DO:
'              ASSIGN ya-se-pago = TRUE.
'              LEAVE.
'              END.
'       END.
Do While Not blv.EOF
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro= " & blv!emp_licnro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        ya_se_pago = True
        rs.Close
        Exit Do
    End If
    blv.MoveNext
    rs.Close
Loop

blv.Close

If Not ya_se_pago Then
    StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
    rs.Open StrSql, objConn
       
    StrSql = "SELECT * FROM vacacion WHERE vacnro = " & rs!vacnro
    rs.Close
    
    rs.Open StrSql, objConn
          
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & rs!vacnro & _
    " AND ternro = " & NroTer
    rs.Close
    
    rs.Open StrSql, objConn
    cantdias = rs!vdiascorcant
    
    NroVac = rs!vacnro
    
    If mes_inicio = 1 Then
        'Call generar_pago(12, ano_inicio - 1, CantDias, 3, Jornal, nrolicencia)
        Call Generar_PagoDescuento(12, ano_inicio - 1, TipDiaPago, Dias_Afecta, 3) 'PAgo
    Else
        'Call generar_pago(mes_inicio - 1, ano_inicio, CantDias, 3, Jornal, nrolicencia)
        Call Generar_PagoDescuento(mes_inicio - 1, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'PAgo
    End If
    
    rs.Close
    
    StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
    rs.Open StrSql, objConn
    
    fin_licencia = False
    mes_afecta = mes_inicio
    ano_afecta = ano_inicio
    dias_pendientes = 0
    dias_restantes = cantdias

    anio_bisiesto = EsBisiesto(ano_afecta)

    primero_mes = AFecha(mes_afecta, 1, ano_afecta)
    ultimo_mes = IIf(mes_afecta = 12, DateAdd("d", -1, AFecha(1, 1, ano_afecta + 1)), _
    DateAdd("d", -1, AFecha(mes_afecta + 1, 1, ano_afecta)))

    If (rs!elfechahasta <= ultimo_mes) Then
        '/* termina en el mes */
        Dias_Afecta = cantdias
    Else
        '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", rs!elfechadesde, ultimo_mes) + 1
    End If

    Do While Not fin_licencia
        dias_restantes = dias_restantes - Dias_Afecta
        'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
        Call Generar_PagoDescuento(mes_afecta, ano_afecta, TipDiaDescuento, Dias_Afecta, 4) 'PAgo
        
        '/* determinar a que continua en el proximo mes */
        If (mes_afecta = 12) Then
            mes_afecta = 1
            ano_afecta = ano_afecta + 1
        Else
            mes_afecta = mes_afecta + 1
        End If

        If (mes_afecta > mes_fin) Or (mes_afecta = 1 And mes_fin = 12) Then
            fin_licencia = True
        End If

        '/* determinar los d¡as que afecta */

        anio_bisiesto = EsBisiesto(ano_afecta)

        primero_mes = AFecha(mes_afecta, 1, ano_afecta)
        ultimo_mes = IIf(mes_afecta = 12, DateAdd("d", -1, AFecha(1, 1, ano_afecta + 1)), _
        DateAdd("d", -1, AFecha(mes_afecta + 1, 1, ano_afecta)))

        If (rs!elfechahasta <= ultimo_mes) Then
        '/* termina en el mes */
            If (dias_restantes < DateDiff("d", primero_mes, ultimo_mes)) Then
                Dias_Afecta = dias_restantes
            Else
                Dias_Afecta = DateDiff("d", primero_mes, rs!elfechahasta) + 1
            End If
        Else
        '/* continua en el mes siguiente */
            Dias_Afecta = DateDiff("d", primero_mes, ultimo_mes) + 1
        End If

    Loop

End If

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub


Public Sub Politica1500_v2AdelantaDescuenta()
' vacpdo50.p
'Genera Pago/dto por dias correspondientes
'paga adelantado todo y descuenta por mes lo que corresponde generar todos los dias de pago

Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_empleado As New ADODB.Recordset
Dim rs_emplic As New ADODB.Recordset
Dim rs_tipdia As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Vacacion As New ADODB.Recordset

Dim Jornal As Boolean
Dim Dias_Afecta As Integer
Dim mes_inicio As Integer
Dim ano_inicio As Integer
Dim mes_fin As Integer
Dim ano_fin As Integer

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro & _
         " AND vacdiascor.vacnro = " & NroVac
OpenRecordset StrSql, rs_vacdiascor

If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo"
    Exit Sub
End If

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_Vacacion

If rs_Vacacion.EOF Then
    'Flog.writeline "No hay dias correspondientes a ese periodo"
    Exit Sub
End If

mes_inicio = Month(rs_Vacacion!vacfecdesde)
ano_inicio = Year(rs_Vacacion!vacfecdesde)
mes_fin = Month(rs_Vacacion!vacfechasta)
ano_fin = Year(rs_Vacacion!vacfechasta)

Dias_Afecta = rs_vacdiascor!vdiascorcant

    StrSql = " SELECT estrnro FROM his_estructura " & _
             " WHERE ternro = " & Ternro & " AND " & _
             " tenro = 22 AND " & _
             " (htetdesde <= " & ConvFecha(fecha_desde) & ") AND " & _
             " ((" & ConvFecha(fecha_desde) & " <= htethasta) or (htethasta is null))"
    OpenRecordset StrSql, rs_Estructura

    If Not rs_Estructura.EOF Then
        If rs_Estructura!estrnro = 1196 Then
            Jornal = False
        Else
            Jornal = True
        End If
    Else
        Flog.writeline "No se encuentra estructura de forma de Liquidacion del empleado"
        Exit Sub
    End If



'StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
'OpenRecordset StrSql, rs_tipdia
'If Not rs_tipdia.EOF Then
'    If Jornal Then
'        TipDiaPago = rs_tipdia!tdinteger4
'        TipDiaDescuento = rs_tipdia!tdinteger5
'    Else
'        TipDiaPago = rs_tipdia!tdinteger1
'        TipDiaDescuento = rs_tipdia!tdinteger2
'    End If
'Else
'    Flog.writeline "Tipo de dia de Vacaciones (2) inexistente"
'    Exit Sub
'End If

Call Politica(1503)
If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503"
        Exit Sub
End If

If Reproceso Then
    StrSql = "SELECT * FROM emp_lic WHERE empleado = " & Ternro & _
             " AND tdnro = 2 "
             
    OpenRecordset StrSql, rs_emplic
    
    If Not rs_emplic.EOF Then
        StrSql = "DELETE FROM vacpagdesc WHERE vacnro = " & NroVac & _
                 " AND ( emp_licnro = " & rs_emplic!emp_licnro & _
                 " OR emp_licnro = 0 AND empleado = " & Ternro & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 1) 'Pago
Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaDescuento, Dias_Afecta, 2) 'Descuento

' Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs_empleado.State = adStateOpen Then rs_empleado.Close
If rs_emplic.State = adStateOpen Then rs_emplic.Close
If rs_tipdia.State = adStateOpen Then rs_tipdia.Close

Set rs_vacdiascor = Nothing
Set rs_empleado = Nothing
Set rs_emplic = Nothing
Set rs_tipdia = Nothing

End Sub


Private Sub Generar_PagoDescuento(ByVal mes_aplicar As Integer, ByVal ano_aplicar As Integer, ByVal TipDia As Integer, ByVal Dias_Afecta As Integer, ByVal anti_vac As Integer)
Dim rs As New Recordset
Dim StrSql As String
Dim TipoDia As Integer
Dim PliqNro As Long

If Not IsNull(mes_aplicar) And Not IsNull(ano_aplicar) Then
    StrSql = "SELECT * FROM periodo WHERE pliqanio = " & ano_aplicar & _
    " AND pliqmes = " & mes_aplicar
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Flog.writeline "Periodo de liquidación inexistente para generar el Descuento de la Lic.Vacaciones del:  " & ano_aplicar & " - " & mes_aplicar
        rs.Close
        Set rs = Nothing
        Exit Sub
    Else
        PliqNro = rs!PliqNro
    End If
End If

If Not IsNull(PliqNro) Then
    StrSql = "INSERT INTO vacpagdesc (emp_licnro,ternro,tprocnro,pago_dto,pliqnro,vacnro,cantdias,manual)" & _
             " VALUES(" & _
             nrolicencia & "," & _
             Ternro & "," & _
             TipDia & "," & _
             anti_vac & "," & _
             PliqNro & "," & _
             NroVac & "," & _
             Dias_Afecta & "," & _
             "0" & _
             ")"
Else
    StrSql = "INSERT INTO vacpagdesc (emp_licnro,ternro,tprocnro,pago_dto,vacnro,cantdias,manual)" & _
             " VALUES(" & _
             "0," & _
             Ternro & "," & _
             TipDia & "," & _
             anti_vac & "," & _
             NroVac & "," & _
             Dias_Afecta & "," & _
             "0" & _
             ")"
End If
objConn.Execute StrSql, , adExecuteNoRecords

End Sub




Public Sub Politica1500v_TMK()
'-----------------------------------------------------------------------
' Customizacion para Temaiken
' paga adelantado todo (en el periodo de liq anterior al que corresponde la licencia.
' descuenta por mes lo que corresponde generar todos los dias de pago
'-----------------------------------------------------------------------

Dim rs As New ADODB.Recordset
Dim mes_inicio As Integer
Dim ano_inicio As Integer
Dim mes_fin    As Integer
Dim ano_fin    As Integer
Dim fin_licencia  As Boolean
Dim mes_afecta    As Integer
Dim ano_afecta    As Integer
Dim primero_mes   As Date
Dim ultimo_mes    As Date
Dim Dias_Afecta   As Integer
Dim dias_pendientes As Integer
Dim dias_restantes As Integer

Dim dias_ya_tomados As Integer
Dim fecha_limite    As Date
Dim anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

Dim Aux_mes_inicio As Integer
Dim Aux_ano_inicio As Integer

Call Politica(1503)
If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503"
        Exit Sub
End If


'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn
' Si no hay licencia me voy
If rs.EOF Then
    Flog.writeline "Licencia inexistente.", vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    mes_inicio = Month(rs!elfechadesde)
    ano_inicio = Year(rs!elfechadesde)
    mes_fin = Month(rs!elfechahasta)
    ano_fin = Year(rs!elfechahasta)
End If
rs.Close

'/* VERIFICAR el reproceso y manejar la depuraci¢n */
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & "AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        StrSql = "SELECT empleg FROM empleado WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Legajo = rs!empleg
        rs.Close
        StrSql = "SELECT terape,ternom FROM tercero WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Nombre = rs!terape & ", " & rs!ternom
        'MsgBox "No se puede Reprocesar una Licencia con pagos y/o descuentos Liquidados. " & Legajo & " " & Nombre, vbCritical
        rs.Close
    Else
    '/* DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA */
        rs.Close
        'StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
StrSql = "SELECT * FROM tipdia WHERE tdnro = " & TipoLicencia
rs.Open StrSql, objConn
If rs.EOF Then
    Flog.writeline " Tipo de licencia (" & TipoLicencia & ") inexistente."
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.Close
End If

StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
OpenRecordset StrSql, rs

'jornal = IIf(rs!folinro = 2, True, False)
rs.Close

'/* POLITICA  - ANALISIS */

StrSql = "SELECT * FROM emp_lic "
If TipoLicencia = 2 Then
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
End If
StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If TipoLicencia = 2 Then
    NroVac = rs!vacnro
End If

'StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'OpenRecordset StrSql, rs
Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1

'Terminar
'RUN generar_pago(mes_inicio, ano_inicio, dias_afecta, 3)
'Call generar_pago(mes_inicio, ano_inicio, Dias_Afecta, 3, Jornal, nrolicencia)
'Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
'fgz - 09/01/2003

'FGZ - 08/09/2004
' El pago debe ser un mes antes del mes de comienzo del Periodo (TMK)
If primer_mes = 1 Then
    Aux_mes_inicio = 12
    Aux_ano_inicio = primer_ano - 1
Else
    Aux_mes_inicio = primer_mes - 1
    Aux_ano_inicio = primer_ano
End If
Call Generar_PagoDescuento(Aux_mes_inicio, Aux_ano_inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
'Call Generar_PagoDescuento(primer_mes, primer_ano, TipDiaPago, Dias_Afecta, 3) 'Pago

'IF RETURN-VALUE <> ""
'          THEN UNDO main, RETURN.
          
'/* tantos descuentos como meses afecte */
fin_licencia = False
mes_afecta = mes_inicio
ano_afecta = ano_inicio
dias_pendientes = 0
dias_restantes = Dias_Afecta

'/* determinar los dias que afecta para el primer mes de decuento */
'/* Genera 30 dias, para todos los meses */
anio_bisiesto = EsBisiesto(ano_afecta)

If mes_afecta = 2 Then
' Date en progress es una función
' la sintaxis es DATE(month,day,year)

    primero_mes = AFecha(mes_afecta, 1, ano_afecta)
    ultimo_mes = IIf(anio_bisiesto = True, AFecha(mes_afecta + 1, 1, ano_afecta), AFecha(mes_afecta + 1, 2, ano_afecta))
Else
    primero_mes = AFecha(mes_afecta, 1, ano_afecta)
    'ultimo_mes = AFecha(mes_afecta, 30, ano_afecta)
    If mes_afecta = 12 Then
        ultimo_mes = AFecha(mes_afecta, 31, ano_afecta)
    Else
        ultimo_mes = AFecha(mes_afecta + 1, 1, ano_afecta) - 1
    End If
End If

If (rs!elfechahasta <= ultimo_mes) Then
'/* termina en el mes */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
'/* continua en el mes siguiente */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, ultimo_mes) + 1
End If

Do While Not fin_licencia
    dias_restantes = dias_restantes - Dias_Afecta
    'Revisar
    'RUN generar_descuento(mes_afecta, ano_afecta, dias_afecta, 4)
    'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
    Call Generar_PagoDescuento(mes_afecta, ano_afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
    'IF RETURN-VALUE <> ""
    '         THEN UNDO main, RETURN.
    
    '/* determinar a que continua en el proximo mes */
    If (mes_afecta = 12) Then
        mes_afecta = 1
        ano_afecta = ano_afecta + 1
    Else
        mes_afecta = mes_afecta + 1
    End If
    If (mes_afecta > mes_fin) Or (mes_afecta = 1 And mes_fin = 12) Then
        fin_licencia = True
    End If
    
    '/* determinar los d¡as que afecta */

    anio_bisiesto = EsBisiesto(ano_afecta)

    If mes_afecta = 2 Then
        primero_mes = AFecha(mes_afecta, 1, ano_afecta)
        ultimo_mes = IIf(anio_bisiesto = True, AFecha(mes_afecta + 1, 1, ano_afecta), AFecha(mes_afecta + 1, 2, ano_afecta))
    Else
        primero_mes = AFecha(mes_afecta, 1, ano_afecta)
        'ultimo_mes = AFecha(mes_afecta, 30, ano_afecta)
        If mes_afecta = 12 Then
            ultimo_mes = AFecha(mes_afecta, 31, ano_afecta)
        Else
            ultimo_mes = AFecha(mes_afecta + 1, 1, ano_afecta) - 1
        End If
    End If

    If (rs!elfechahasta <= ultimo_mes) Then
    '/* termina en el mes */
        If (dias_restantes < DateDiff("d", primero_mes, ultimo_mes)) Then
            Dias_Afecta = dias_restantes
        Else
            Dias_Afecta = DateDiff("d", primero_mes, rs!elfechahasta) + 1
        End If
    Else
    '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", primero_mes, ultimo_mes) + 1
    End If
Loop

rs.Close
Set rs = Nothing

End Sub


Public Sub Politica1500PagayDescuenta_Mes_a_Mes()
' corresponde a vacpdo02

'/* ***************************************************************
'PAGA Y DESCUENTA POR MES SIN TOPE DE 30 DIAS FIJOS.
'
'*************************************************************** */

Dim rs As New Recordset
Dim StrSql As String

Dim mes_inicio As Integer
Dim ano_inicio As Integer
Dim mes_fin    As Integer
Dim ano_fin    As Integer
Dim fin_licencia  As Boolean
Dim mes_afecta    As Integer
Dim ano_afecta    As Integer
Dim primero_mes   As Date
Dim ultimo_mes    As Date
Dim Dias_Afecta   As Integer
Dim dias_pendientes As Integer
Dim dias_restantes As Integer

'DEF BUFFER buf_lic FOR emp_lic.
Dim dias_ya_tomados As Integer
Dim fecha_limite    As Date
Dim anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

'/**************************************************************************************************/

Call Politica(1503)
If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503"
        Exit Sub
End If


'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
rs.Open StrSql, objConn
' Si no hay licencia me voy
If rs.EOF Then
    'MsgBox "Licencia inexistente.", vbCritical
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    'NroTer = emp_lic.empleado
    NroTer = rs!Empleado
    mes_inicio = Month(rs!elfechadesde)
    ano_inicio = Year(rs!elfechadesde)
    mes_fin = Month(rs!elfechahasta)
    ano_fin = Year(rs!elfechahasta)
End If
rs.Close

'/* VERIFICAR el reproceso y manejar la depuraci¢n */
If Not Reproceso Then
    '/* si no es reproceso y existe el desglose de pago/descuento, salir */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Exit Sub
    End If
    rs.Close
Else
    '/*  verificar si no tiene pagos/descuentos ya procesados, sino depurarlos */
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & "AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        StrSql = "SELECT empleg FROM empleado WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Legajo = rs!empleg
        rs.Close
        StrSql = "SELECT terape,ternom FROM tercero WHERE ternro = " & NroTer
        rs.Open StrSql, objConn
        Nombre = rs!terape & ", " & rs!ternom
        'MsgBox "No se puede Reprocesar una Licencia con pagos y/o descuentos Liquidados. " & Legajo & " " & Nombre, vbCritical
        rs.Close
    Else
    '/* DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA */
        rs.Close
        StrSql = "DELETE vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


StrSql = "SELECT * FROM tipdia WHERE tdnro = " & TipoLicencia
rs.Open StrSql, objConn
If rs.EOF Then
    'MsgBox " Tipo de dia de Vacaciones (2) inexistente."
    rs.Close
    Set rs = Nothing
    Exit Sub
Else
    rs.Close
End If

StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
rs.Open StrSql, objConn
'jornal = IIf(rs!folinro = 2, True, False)
rs.Close

'/* POLITICA  - ANALISIS */

fin_licencia = False
mes_afecta = mes_inicio
ano_afecta = ano_inicio
dias_pendientes = 0

'/* determinar los dias que afecta para el primer mes de descuento */
primero_mes = AFecha(mes_afecta, 1, ano_afecta)
ultimo_mes = IIf(mes_afecta = 12, DateAdd("d", -1, AFecha(1, 1, ano_afecta + 1)), _
DateAdd("d", -1, AFecha(mes_afecta + 1, 1, ano_afecta)))

StrSql = "SELECT * FROM emp_lic "
If TipoLicencia = 2 Then
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
End If
StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If TipoLicencia = 2 Then
    NroVac = rs!vacnro
End If

'StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'rs.Open StrSql, objConn

If (rs!elfechahasta <= ultimo_mes) Then
'/* termina en el mes */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
'/* continua en el mes siguiente */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, ultimo_mes) + 1
End If
       
Do While Not fin_licencia
    'RUN generar_pago  (mes_afecta,ano_afecta,dias_afecta,3)
    'Call generar_pago(mes_inicio, ano_inicio, Dias_Afecta, 3, Jornal, nrolicencia)
    'Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
    'FGZ - 15/12/2004
    Call Generar_PagoDescuento(mes_afecta, ano_afecta, TipDiaPago, Dias_Afecta, 3) 'Pago
    'RUN generar_descuento  (mes_afecta,ano_afecta,dias_afecta,4)
    'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
    'Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaDescuento, Dias_Afecta, 4) 'Pago
    'FGZ - 15/12/2004
    Call Generar_PagoDescuento(mes_afecta, ano_afecta, TipDiaDescuento, Dias_Afecta, 4) 'Pago
    
    If (mes_afecta = 12) Then
        mes_afecta = 1
        ano_afecta = ano_afecta + 1
    Else
        mes_afecta = mes_afecta + 1
    End If
    
    If (mes_afecta > mes_fin) Or (mes_afecta = 1 And mes_fin = 12) Then
        fin_licencia = True
    End If

    '/* determinar los d­as que afecta */
    primero_mes = AFecha(mes_afecta, 1, ano_afecta)
    ultimo_mes = IIf(mes_afecta = 12, DateAdd("d", -1, AFecha(1, 1, ano_afecta + 1)), _
    DateAdd("d", -1, AFecha(mes_afecta + 1, 1, ano_afecta)))
    
    If (rs!elfechahasta <= ultimo_mes) Then
    '/* termina en el mes */
        Dias_Afecta = DateDiff("d", primero_mes, rs!elfechahasta) + 1
    Else
    '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", primero_mes, ultimo_mes) + 1
    End If

Loop

End Sub


