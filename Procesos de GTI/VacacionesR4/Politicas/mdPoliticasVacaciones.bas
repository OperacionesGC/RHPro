Attribute VB_Name = "mdlPoliticasVac"
Option Explicit

Global rsPolitica As New ADODB.Recordset
Global PoliticaOK As Boolean 'si cargo bien o no la politica llamada
Global rs_Periodos_Vac As New ADODB.Recordset

Global DiasProporcion As Integer
Global FactorDivision As Single
Global TipoVacacionProporcion As Long
Global TipoVacacionProporcionCorr As Long
Global BaseAntiguedad As Integer

Global fecha_desde As Date
Global fecha_hasta As Date
Global Periodo_Anio As Long


Global Total_Dias_A_Generar As Long
Global TotalGeneral_Dias_A_Generar As Long
Global Genera_Dto_DiasCorr As Boolean
Global Genera_Pagos As Boolean
Global Genera_Descuentos As Boolean
Global Generar_Fecha_Desde As Date
Global Aux_Generar_Fecha_Desde As Date
Global GenerarSituacionRevista As Boolean
Global Diashabiles_LV As Boolean
Global DiasAcordados As Boolean
Global CalculaVencimientos As Boolean
Global Dias_efect_trab_anio As Boolean
Global Dias_Bonificacion As Boolean

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
Global st_TipoDia2 As Integer
Global st_CantidadDias As Integer
Global st_FactorDivision As Integer
Global st_Escala As Integer
Global st_ModeloPago As Integer
Global st_ModeloDto As Integer
Global st_redondeo As Integer
Global st_BaseAntiguedad As Integer
Global st_Factor As Double
Global st_ModeloPago2 As Integer
Global st_ModeloDto2 As Integer
Global st_Dias As Long
Global st_Modifica As String
Global st_ListaF1 As String
Global st_ListaF2 As String
Global st_ListaF3 As String
Global st_TipoHora2 As Long
Global st_Minutos As Integer
Global st_Anormalidad As String
Global st_ModeloPais As String 'NG - 07/11/2013



'FGZ - 04/03/2010 - Variables agregadas
Global st_Continua As Boolean
Global st_ListaTH As String
Global st_Anormalidad2 As String
Global st_Anormalidad3 As String
Global st_TipoHora3 As Long
Global st_Tolerancia2 As String
Global st_Tolerancia3 As String
Global st_Dia As Long
Global st_Mes As Long


Global CantidadLicenciasProcesadas As Integer
Global Subturno_Genera As Integer
'EAM- 12/07/2010
Global Lic_Descuento As String

Global Ya_Pago As Boolean
Global alcannivel As Integer

Global listapgdto 'mdf   la habia declarado en el modulo de pagos y descuentos

Public Sub SetearParametrosPolitica(ByVal Detalle As Long, ByRef ok As Boolean)
'NG - 07/11/213 - Se agrego case 45
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

    
    st_Dias = 180
    
    Do While Not rsPolitica.EOF
        Select Case rsPolitica!polparamnro
        Case 1:
            st_Opcion = CInt(rsPolitica!polparamvalor)
        Case 2:
            st_VentSal = Format(rsPolitica!polparamvalor, "0000")
        Case 3:     'FGZ - 05/11/2008
            If Not EsNulo(rsPolitica!polparamvalor) Then
                If UCase(rsPolitica!polparamvalor) = "SI" Or rsPolitica!polparamvalor = "-1" Or rsPolitica!polparamvalor = "1" Then
                    st_Continua = True
                Else
                    st_Continua = False
                End If
            Else
                st_Continua = True
            End If
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
        Case 16:
            st_redondeo = CInt(rsPolitica!polparamvalor)
        Case 17:
            st_ListaTH = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 18:
            st_BaseAntiguedad = CInt(rsPolitica!polparamvalor)
        Case 19:
            st_Factor = CDbl(rsPolitica!polparamvalor)
        Case 20:
            'EAM- 19/11/2012 - Se agrego la validación para que seté en 0 cuando no esta configurado el parametro.
            st_ModeloPago2 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, 0)
        Case 21:
            'EAM- 19/11/2012 - Se agrego la validación para que seté en 0 cuando no esta configurado el parametro.
            st_ModeloDto2 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, 0)
        Case 22: ' Diego Rosso 12/11/2007
            st_Modifica = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 23:    'FGZ - 31/01/2008
            st_ListaF1 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 24:    'FGZ - 31/01/2008
            st_ListaF2 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 25:    'FGZ - 31/01/2008
            st_ListaF3 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 26:    'FGZ - 18/06/2008
            st_TipoHora2 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 27:    'FGZ - 18/06/2008
            st_Minutos = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 28:    'FGZ - 05/11/2008
            st_Anormalidad = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 29:    'GER - 05/05/2009
            st_Dias = CLng(rsPolitica!polparamvalor)
        Case 30:    'FGZ - 18/01/2010
            st_Tolerancia2 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 31:    'FGZ - 18/01/2010
            st_Tolerancia3 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 32:    'FGZ - 18/01/2010
            st_TipoHora3 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 33:    'FGZ - 18/01/2010
            st_Anormalidad2 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 34:    'FGZ - 18/01/2010
            st_Anormalidad3 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, "")
        Case 35:
            st_Dia = IIf(Not EsNulo(rsPolitica!polparamvalor), CLng(rsPolitica!polparamvalor), 0)
        Case 36:
            st_Mes = IIf(Not EsNulo(rsPolitica!polparamvalor), CLng(rsPolitica!polparamvalor), 0)
        Case 44:    'EAM- 18/01/2011 Se agregó para configurar el tipo de vacacines corridos y si se quiere cualquier otro
            st_TipoDia2 = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, 0)
        Case 45:    'NG - 07/11/2013 Se agregó para configurar el Modelo de país a utilizar.
            st_ModeloPais = IIf(Not EsNulo(rsPolitica!polparamvalor), rsPolitica!polparamvalor, 0)
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
' Ultima modificacion: FGZ - 28/01/2008 cambié fecha_desde por aux_fecha
' Ultima modificacion: Gonzalez Nicolás - 07/11/2013 - Se agregó llamada a la 1515
' --------------------------------------------------------------
Dim objRs As New ADODB.Recordset 'Como esta función es recursiva el recordset lo tengo que definir en forma local
Dim StrSql As String
Dim det As Integer
Dim cabecera As Long
Dim Detalle As Long
Dim Aux_Fecha As Date

    'FGZ - 28/01/2008 -----------------
    If fecha_desde > Date Then
        Aux_Fecha = fecha_desde
    Else
        If fecha_hasta > Date Then
            Aux_Fecha = Date
        Else
            Aux_Fecha = fecha_hasta
        End If
    End If
    'FGZ - 28/01/2008 -----------------
    StrSql = "SELECT gti_cabpolitica.cabpolnro,gti_cabpolitica.cabpolnivel, gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro, gti_detpolitica.detpolprograma "
    StrSql = StrSql & " FROM gti_cabpolitica INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro "
    'StrSql = StrSql & " WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 3 And gti_alcanpolitica.alcpolorigen = " & Empleado.Ternro & " AND gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 "
    StrSql = StrSql & " WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 3 And gti_alcanpolitica.alcpolorigen = " & Ternro & " AND gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 " 'mdf
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
        StrSql = StrSql & " And empleado.ternro = " & Ternro
        StrSql = StrSql & " And (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ")"
        StrSql = StrSql & " And ((" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
        StrSql = StrSql & " ORDER BY alcance_testr.AlteOrden Asc "
        'StrSql = StrSql & " And empleado.ternro = " & Empleado.Ternro
        'StrSql = StrSql & " And his_estructura.htethasta IS NULL "
        'StrSql = StrSql & " And (his_estructura.htetdesde <= " & ConvFecha(p_fecha) & ")"
        'StrSql = StrSql & " And ((" & ConvFecha(p_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
        'StrSql = StrSql & " ORDER BY alcance_testr.AlteOrden Asc "
        OpenRecordset StrSql, objRs
        If objRs.EOF Then
            StrSql = " SELECT gti_cabpolitica.cabpolnro,gti_cabpolitica.cabpolnivel, gti_alcanpolitica.alcpolnivel, gti_alcanpolitica.alcpolorigen, gti_detpolitica.detpolnro, gti_detpolitica.detpolprograma "
            StrSql = StrSql & " FROM gti_cabpolitica INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro "
            StrSql = StrSql & " WHERE gti_cabpolitica.cabpolnivel = " & Numero & " And gti_alcanpolitica.alcpolnivel = 1 And gti_cabpolitica.cabpolestado = -1 And gti_alcanpolitica.alcpolestado = -1 "
            OpenRecordset StrSql, objRs
        End If
    End If
    
    
    
    If Not objRs.EOF Then
        det = objRs!detpolprograma
        cabecera = objRs!cabpolnro
        Detalle = objRs!detpolnro
        Flog.writeline "Inicio Politica " & Numero
        If depurar Then GeneraTraza Empleado.Ternro, p_fecha, "Nª " & Numero, Str(det)
        Select Case Numero
        Case 1500: 'Vacaciones de pago/dto
            Call politica1500(det, cabecera, Detalle)
        Case 1501: 'Proporcion de dias de Vacaciones
            Call politica1501(det, cabecera, Detalle)
        Case 1502: 'Escala
            Call politica1502(det, cabecera, Detalle)
        Case 1503: 'Modelo de liq. pago/dto
            Call politica1503(det, cabecera, Detalle)
        Case 1504: 'Modelo de liq. TTI
            Call politica1504(det, cabecera, Detalle)
        Case 1505: 'Tipo base antiguedad
            Call politica1505(det, cabecera, Detalle)
        Case 1506: 'Descuentos por Dias Correspondientes
            Call politica1506(det, cabecera, Detalle)
        Case 1507: 'Generacion Pagos / Descuentos
            Call politica1507(det, cabecera, Detalle)
        Case 1508:  'Dtos de dias de licencias
            Call politica1508(det, cabecera, Detalle)
        Case 1509: 'Situacion de revista
            Call politica1509(det, cabecera, Detalle)
        Case 1510: 'Licencias se gozan por Dias habiles
            Call politica1510(det, cabecera, Detalle)
        Case 1511: 'Vacaciones Acordadas
            Call politica1511(det, cabecera, Detalle)
        Case 1512: 'Vencimientos de vacaciones
            Call politica1512(det, cabecera, Detalle)
        Case 1513: 'Días efectivamente trabajados en el ultimo año.
            Call politica1513(det, cabecera, Detalle)
        Case 1514: 'Bonificación de días de vacaciones
            Call politica1514(det, cabecera, Detalle)
        Case 1515: 'Setea versiones de los modelos a utilizar.
            Call politica1515(det, cabecera, Detalle)
        Case 1516: 'EAM- Descuento de días de vacaciones por licencias.
            Call politica1516(det, cabecera, Detalle)
        Case Else
            Flog.writeline "Politica No Codificada."
        End Select
        Flog.writeline "Fin Politica " & Numero
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

Private Sub politica1501(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
    'EAM- Se setea el redondeo en 2 que es default si no está configurado
    st_redondeo = 2
    st_Opcion = subn
    Call SetearParametrosPolitica(Detalle, ok)
    DiasProporcion = st_CantidadDias
    FactorDivision = st_FactorDivision
    If CInt(st_BaseAntiguedad) > 0 Then
        BaseAntiguedad = CInt(st_BaseAntiguedad)
    Else
        BaseAntiguedad = 6
    End If
    
        
    'EAM (18-01-12)- Se toma el tipodia1 para determinar el tipo de vacaciones del empleado y el tipodia2 para calcular los dias corridos o lo que se configure.
    'luego se busca en escala lo que corresponde para cada tipo de vacacion
    TipoVacacionProporcion = IIf(Not EsNulo(st_TipoDia1), st_TipoDia1, 0)
    TipoVacacionProporcionCorr = IIf(Not EsNulo(st_TipoDia2), st_TipoDia2, 0)
    

    
End Sub


Private Sub politica1502(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)

    Call SetearParametrosPolitica(Detalle, ok)
    If ok Then
        NroGrilla = st_Escala
        PoliticaOK = True
    Else
        PoliticaOK = False
    End If
    
End Sub

Private Sub politica1503(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)

    Call SetearParametrosPolitica(Detalle, ok)
    If ok Then
        TipDiaPago = st_ModeloPago
        TipDiaDescuento = st_ModeloDto
        TipDiaPago2 = st_ModeloPago2
        TipDiaDescuento2 = st_ModeloDto2
        PoliticaOK = True
    Else
        PoliticaOK = False
    End If
    
End Sub

Private Sub politica1504(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)

    PoliticaOK = True
End Sub


Private Sub politica1500(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
Dim Op As Integer

    Op = subn
    Flog.writeline
    Flog.writeline "Programa " & Op
    Select Case Op
        Case 1:
            
            If GeneraPorLicencia Then
                Flog.writeline "La versión 1 (Licencia) de la politica 1500 es reemplazada por la versión 12."
            Else
                Flog.writeline "La versión 1 (Días Correspondientes) de la politica 1500 es reemplazada por la versión 4."

            End If
        Case 2:
            If GeneraPorLicencia Then
                Flog.writeline "La versión 2 (Licencia) de la politica 1500 es reemplazada por la versión 4."
            Else
                Flog.writeline "La versión 1 (Días Correspondientes) de la politica 1500 es reemplazada por la versión 4."

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
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v2AdelantaDescuenta
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If

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
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro

                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v2AdelantaDescuenta
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If

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
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v2AdelantaDescuenta
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 6:
            If GeneraPorLicencia Then
                Flog.writeline "La versión 6 (Licencia) deja de tener soporte y funcionalidad. Ver otras versiones en el manual"
            Else
                Flog.writeline "La versión 6 (Días Correspondientes) deja de tener soporte y funcionalidad. Ver otras versiones en el manual"

            End If
        Case 7: 'TMK
            If GeneraPorLicencia Then
                Call Politica1500v_TMK
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v2AdelantaDescuenta
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 8: '???
            If GeneraPorLicencia Then
                Flog.writeline "La versión 8 (Licencia) de la politica 1500 es reemplazada por la versión 22."
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro

                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        'Call Politica1500_v2AdelantaDescuenta
                        Call Politica1500_V2PagaDescuenta_PorMes
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 9: 'Glencore
            If GeneraPorLicencia Then
                Call Politica1500AdelantaDescuenta_Jornal
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro

                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v2AdelantaDescuenta_Jornal
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 10: 'Schering
            If GeneraPorLicencia Then
                Call Politica1500_AdelantaDescuenta_Schering
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro

                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v2AdelantaDescuenta
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 11: 'al menos un dia en el mes y genera pago en mes actual
            If GeneraPorLicencia Then
                CantidadLicenciasProcesadas = CantidadLicenciasProcesadas + 1
                Call Politica1500v_7
            Else
               If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                       Call Politica1500_v2AdelantaDescuenta
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    
                    
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
            
        Case 12: 'Paga la totalidad de días de la licencia (un solo pago).
            If GeneraPorLicencia Then
                Call Politica1500AdelantaDescuenta_Nueva
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        'Call Politica1500_v2AdelantaDescuenta
                        Call Politica1500_V2AdelantaDescuenta_Nueva
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 13: '
            If GeneraPorLicencia Then
                Flog.writeline "La versión 13 (Licencia) de la politica 1500 es reemplazada por la versión 12."
            Else
                Flog.writeline "La versión 1 (Días Correspondientes) de la politica 1500 es reemplazada por la versión 4."

            End If
        Case 14: '
            If GeneraPorLicencia Then
                'CantidadLicenciasProcesadas = CantidadLicenciasProcesadas + 1
                Call Politica1500v_14
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Período entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Período " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v14_DiasCorresp
                    Else
                        Flog.writeline "Ya se generaron todos los días correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 15: 'AGD
            If GeneraPorLicencia Then
                Call Politica1500_PagoDescuento_AGD
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_V2PagoDescuento_AGD
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 16: ' SMT
            If GeneraPorLicencia Then
                Call Politica1500v_16
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Período entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Período " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v16_DiasCorresp
                    Else
                        Flog.writeline "Ya se generaron todos los días correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 17: 'ARLEI
            If GeneraPorLicencia Then
                Call Politica1500_AdelantaDescuenta_ARLEI
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro

                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v2AdelantaDescuenta
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 18: 'Papelbril
            If GeneraPorLicencia Then
                Flog.writeline "La versión 8 (Licencia) de la politica 1500 es reemplazada por la versión 22."
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro

                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        'Call Politica1500_V2PagaDescuenta_PorMes
                        Call Politica1500_V2PagaDescuenta_PorMes_Papelbril
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 19: 'Monresa (Uruguay)
            If GeneraPorLicencia Then
                'Call Politica1500_Uruguay
                Call Politica1500_Monresa
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro

                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_V2PagaDescuenta_PorMes_Papelbril
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 20: ' BAPRO
'            If GeneraPorLicencia Then
'                Call Politica1500_v20
            If GeneraPorLicencia Then
                Flog.writeline "La versión 13 (Licencia) de la politica 1500 es reemplazada por la versión 12."

            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v20_PorDiasCorresp
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        
        Case 21: 'Cooperativa Seguro
            
            If GeneraPorLicencia Then
                'Ver
                'Call Politica1500AdelantaDescuenta
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        'Call Politica1500_v3AdelantaDescuenta
                        Call Politica1500_V3AdelantaDescuenta_Nueva
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    
                    rs_Periodos_Vac.MoveNext
                Loop
                
            End If
        Case 22: 'TIMBO
            
            If GeneraPorLicencia Then
                Call Politica1500_V2PagaDescuenta_PorQuincena
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_V3AdelantaDescuenta_Nueva
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
            
        Case 23: 'PORTUGAL | SOLO GENERA PAGOS
            If GeneraPorLicencia Then
                Call Politica1500_PagoLic_PT
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro

                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_PagoXdiasCorr_PT
                        'Call Politica1500_V2PagaDescuenta_PorMes_Papelbril
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    rs_Periodos_Vac.MoveNext
                Loop
            End If

        Case 24:
            If GeneraPorLicencia Then
                'Al menos una licencia, paga todo y descuenta todo (por Quincena o mensual)
                CantidadLicenciasProcesadas = CantidadLicenciasProcesadas + 1
                Call Politica1500v_24
            Else
               If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                       Call Politica1500_v2AdelantaDescuenta
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If
                    
                    
                    rs_Periodos_Vac.MoveNext
                Loop
            End If
            
           Case 25: 'HORWATH LITORAL - AMR
            If GeneraPorLicencia Then
                Flog.writeline "Esta versión solo se utiliza para días correspondientes."
                Exit Sub
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        'Call Politica1500_v20_PorDiasCorresp
                        Call Politica1500_v25_PorDiasCorresp
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
            End If
        Case 26: 'nueva politica seba
            If GeneraPorLicencia Then
                CantidadLicenciasProcesadas = CantidadLicenciasProcesadas + 1
                Call Politica1500v_7_1("L")
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_V4AdelantaDescuenta_Nueva
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
                
            End If
            
        Case 27: 'EAM (v) - Nueva versión que toma los días habiles y cada 5 días paga 7
            If GeneraPorLicencia Then
                CantidadLicenciasProcesadas = CantidadLicenciasProcesadas + 1
                'Call Politica1500v_7_1("L")
                'Call Politica1500_V2PagaDescuenta_PorQuincena   MDF
                Call Politica1500_V2PagaDescuenta_PorQuincena_sykes   'mdf, cree esta nueva porq usaba la misma q la 22
                
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Periodo entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_V4AdelantaDescuenta_Nueva
                    Else
                        Flog.writeline "Ya se generaron todos los dias correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
                
            End If
        Case 28: ' SGS
            If GeneraPorLicencia Then
                Call Politica1500_v3PagaDescuenta
            Else
                Flog.writeline "Esta versión solo se utiliza para licencias."
                Exit Sub
            End If
        Case 29: 'Version para TATA- Pago todo y descuenta por mes. Por dias habiles y teniendo en ceunta la fecha de iniciao de la licencia
            If GeneraPorLicencia Then
                Call Politica1500v_TATA
            Else
                If Not rs_Periodos_Vac.EOF Then
                    rs_Periodos_Vac.MoveFirst
                Else
                    Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
                End If
                Flog.writeline
                Flog.writeline "Para Cada Período entre el rango de fechas "
                Flog.writeline
                Do While Not rs_Periodos_Vac.EOF
                    Flog.writeline "Período " & rs_Periodos_Vac!vacnro
                    NroVac = rs_Periodos_Vac!vacnro
                    If Total_Dias_A_Generar > 0 Then
                        Call Politica1500_v14_DiasCorresp
                    Else
                        Flog.writeline "Ya se generaron todos los días correspondientes pretendidos"
                    End If

                    rs_Periodos_Vac.MoveNext
                Loop
            End If
            
            
    End Select
End Sub



Public Function AFecha(ByVal M As Integer, ByVal D As Integer, ByVal a As Integer) As Date
' Reemplaza a la función Date de Progress
'ultimo-mes  = DATE (mes-afecta,30,ano-afecta)
Dim auxi
  
  auxi = Str(M) & "/" & Str(D) & "/" & Str(a)
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
Dim Dia As Integer

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
    Dia = Day(rs!terfecnac)
    FechaLiqVacaciones = AFecha(mes, Dia, anio_vacac)
End If

rs.Close
Set rs = Nothing

End Function

Public Sub Politica1500AdelantaDescuenta_Jornal()
'-----------------------------------------------------------------------
'Customizacion para Glencore.
' Jornales= paga adelantado todo
' y
' No Jornales= paga adelantado todo y descuenta
'       por mes lo que corresponde generar todos los dias de pago.
'Fecha: 27/07/2005
'Autor: FGZ
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Restantes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date

Dim rs As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

On Error GoTo CE

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs

' Si no hay licencia me voy
If rs.EOF Then
    Flog.wirteline Espacios(Tabulador * 1) & "Licencia inexistente."
    GoTo CE
Else
    Aux_Fecha = rs!elfechahasta
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'EAM- Verifica si existe el tipo de licencia
If Not ExisteTipoLicencia(TipoLicencia) Then
    Flog.writeline Espacios(Tabulador * 1) & "Tipo de dia de Vacaciones (2) inexistente."
    GoTo CE
End If


'EAM- Busca la forma de liquidacion True-->Jornal | false -> Mensual
Jornal = FormaDeLiquidacion(Aux_Fecha, 22)


'POLITICA  - ANALISIS
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


' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(rs!elfechadesde) > 15 Then
        TipDiaPago = TipDiaDescuento
    End If
End If


'fgz - 09/01/2003
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago


'/* tantos descuentos como meses afecte */
Fin_Licencia = False
Mes_Afecta = Mes_Inicio
Ano_Afecta = Ano_Inicio
Dias_Restantes = Dias_Afecta

'/* determinar los dias que afecta para el primer mes de decuento */
'/* Genera 30 dias, para todos los meses */
Anio_bisiesto = EsBisiesto(Ano_Afecta)

If Mes_Afecta = 2 Then
' Date en progress es una función
' la sintaxis es DATE(month,day,year)

    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
Else
    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    'ultimo_mes = AFecha(mes_afecta, 30, ano_afecta)
    If Mes_Afecta = 12 Then
        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
    Else
        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    End If
End If

If (rs!elfechahasta <= Ultimo_Mes) Then
'/* termina en el mes */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
'/* continua en el mes siguiente */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
End If

Do While Not Fin_Licencia
    Dias_Restantes = Dias_Restantes - Dias_Afecta
    
    'Revisar
    If Not Jornal Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
    End If
    
    '/* determinar a que continua en el proximo mes */
    If (Mes_Afecta = 12) Then
        Mes_Afecta = 1
        Ano_Afecta = Ano_Afecta + 1
    Else
        Mes_Afecta = Mes_Afecta + 1
    End If
    If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
        Fin_Licencia = True
    End If
    
    '/* determinar los d¡as que afecta */
    Anio_bisiesto = EsBisiesto(Ano_Afecta)

    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    Else
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        'ultimo_mes = AFecha(mes_afecta, 30, ano_afecta)
        If Mes_Afecta = 12 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
        Else
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    End If

    If (rs!elfechahasta <= Ultimo_Mes) Then
    '/* termina en el mes */
        If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
            Dias_Afecta = Dias_Restantes
        Else
            Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
        End If
    Else
    '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
    End If
Loop
GoTo ProcesadoOK

CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing
End Sub

Public Sub Politica1500AdelantaDescuenta_Nueva()
'-----------------------------------------------------------------------
'Descripcion: paga adelantado todo y descuenta por mes lo que corresponde generar
' Autor     :
'Ult Mod    : FGZ - 12-12-2005
'             Cuando el legajo es jornal realiza los cortes y topeos por quincena
'Ult Mod    : FGZ - 07/04/2006
            'Aux_TipDiaDescuento = TipDiaDescuento
            ' GER - 04/09/2008 - Se arreglo los descuentos
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer

Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Aux_Fecha As Date
Dim Dias_Acumulados As Integer

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer
Dim rs As New ADODB.Recordset

On Error GoTo CE

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs

'EAM- Si no hay licencia deja de analizar
If rs.EOF Then
    Flog.writeline "Licencia inexistente.", vbCritical
    GoTo CE
Else
    Aux_Fecha = rs!elfechahasta
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
    Flog.writeline "Licencia desde " & rs!elfechadesde & " al " & rs!elfechahasta
End If
rs.Close

Dias_Acumulados = 0
Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline "Ya existe Pago/Descuento. Se debe Reprocesar"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        Flog.writeline "DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA"
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


StrSql = "SELECT * FROM tipdia WHERE tdnro = " & TipoLicencia
OpenRecordset StrSql, rs
If rs.EOF Then
    Flog.writeline " Tipo de licencia (" & TipoLicencia & ") inexistente."
    GoTo CE
End If

'EAM- Busca la forma de liquidacion True-->Jornal | false -> Mensual
Jornal = FormaDeLiquidacion(Aux_Fecha, 22)


'POLITICA  - ANALISIS
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
Flog.writeline "Dias afecta " & Dias_Afecta


Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento

'Verifica a que quincena corresponde el primer descuento
If Jornal Then
    'reviso a que quincena corresponde
    If Day(rs!elfechadesde) > 15 Then
        'es segunda quincena
        Aux_TipDiaDescuento = st_ModeloDto2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = st_ModeloDto
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(rs!elfechahasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
Else
    Quincena_Inicio = 1
End If
Quincena_Siguiente = Quincena_Inicio

If Genera_Pagos Then
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
End If

'tantos descuentos como meses afecte
Fin_Licencia = False
Mes_Afecta = Mes_Inicio
Ano_Afecta = Ano_Inicio
Dias_Pendientes = 0
Dias_Restantes = Dias_Afecta

'determinar los dias que afecta para el primer mes de decuento
'Genera 30 dias, para todos los meses
Anio_bisiesto = EsBisiesto(Ano_Afecta)

If Mes_Afecta = 2 Then
    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    If Not Jornal Then
        'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    Else
        If Quincena_Siguiente = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    End If
Else
    
    If Not Jornal Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Mes_Afecta = 12 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
        Else
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    Else
        If Quincena_Siguiente = 1 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    End If
End If

If (rs!elfechahasta <= Ultimo_Mes) Then
    'termina en el mes
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
    'continua en el mes siguiente
    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
End If


    'FGZ - 27/02/2006 No estaba inicializando estas variables y por lo tanto
    '      siempre comenzaba en la primera quincena (cuando Jornal)
    Primero_Mes = AFecha(Mes_Afecta, Day(rs!elfechadesde), Ano_Afecta)
    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1

    'Aux_TipDiaPago = TipDiaPago
    Aux_TipDiaDescuento = TipDiaDescuento
    
    If Jornal Then
        'reviso a que quincena corresponde
        If Day(Primero_Mes) > 15 Then
            'es segunda quincena
            Aux_TipDiaDescuento = st_ModeloDto2
            Quincena_Siguiente = 2
        Else
            'Es primera quincena
            Aux_TipDiaDescuento = st_ModeloDto
            Quincena_Siguiente = 1
        End If
    End If
    
    
    
Do While Not (Fin_Licencia) Or Dias_Acumulados > 0
    
    If Jornal Then
        If Dias_Afecta > 15 Then
            Dias_Afecta = 15
            Dias_Acumulados = 1
        End If
        If Dias_Acumulados > 0 And Dias_Afecta < 15 Then
            If (Dias_Acumulados + Dias_Afecta) <= 15 Then
                Dias_Afecta = Dias_Afecta + Dias_Acumulados
                Dias_Acumulados = 0
            End If
        End If
    Else
        If Dias_Afecta > 30 Then
            Dias_Afecta = 30
            Dias_Acumulados = 0
        End If
        If Dias_Acumulados > 0 And Dias_Afecta < 30 Then
            If (Dias_Acumulados + Dias_Afecta) <= 30 Then
                Dias_Afecta = Dias_Afecta + Dias_Acumulados
                Dias_Acumulados = 0
            End If
        End If
        'Ver si va
        'Aux_TipDiaDescuento = 7
    End If
    
    Dias_Restantes = Dias_Restantes - Dias_Afecta
    
    If Genera_Descuentos And (Dias_Afecta <> 0) Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
    End If
            
    If Jornal Then
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
    End If
    
    'determinar a que continua en el proximo mes
    If (Mes_Afecta = 12) Then
        If Not Jornal Then
            Mes_Afecta = 1
            Ano_Afecta = Ano_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    Else
        If Not Jornal Then
            Mes_Afecta = Mes_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    End If
    
    If (Ano_Afecta = Ano_Fin) Then
        If (Mes_Afecta > Mes_Fin) Then
            Fin_Licencia = True
        Else
            If (Mes_Afecta = Mes_Fin) And Jornal Then
                If Quincena_Siguiente > Quincena_Fin Then
                    Fin_Licencia = True
                End If
            End If
        End If
    Else
        If (Ano_Afecta > Ano_Fin) Then
            Fin_Licencia = True
        End If
    End If
    
    'determinar los dias que afecta
    Anio_bisiesto = EsBisiesto(Ano_Afecta)



    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                'Ultimo_Mes = IIf(Anio_bisiesto, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If


    If (rs!elfechahasta <= Ultimo_Mes) Then
    'termina en el mes
        If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
            Dias_Afecta = Dias_Restantes
        Else
            Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
        End If
    Else
        'continua en el mes siguiente
        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
    End If
    
    If Not Jornal Then
        If Dias_Afecta > 30 Then
           Dias_Afecta = 30
        End If
    Else
        If Dias_Afecta > 15 Then
           Dias_Afecta = 15
        End If
    End If
    
    If Jornal Then
        'reviso a que quincena corresponde
        If Day(Primero_Mes) > 15 Then
            'es segunda quincena
            Aux_TipDiaDescuento = st_ModeloDto2
        Else
            'Es primera quincena
            Aux_TipDiaDescuento = st_ModeloDto
        End If
    End If
    
Loop
GoTo ProcesadoOK
CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing
End Sub

Public Sub Politica1500_v20()
'-----------------------------------------------------------------------
'Descripcion: paga adelantado todo y descuenta por mes lo que corresponde generar
' Autor     :
'Ult Mod    : FGZ - 12-12-2005
'             Cuando el legajo es jornal realiza los cortes y topeos por quincena
'Ult Mod    : FGZ - 07/04/2006
            'Aux_TipDiaDescuento = TipDiaDescuento
            ' GER - 04/09/2008 - Se arreglo los descuentos
            ' FAF - Version nueva a partir de 12 - Politica1500AdelantaDescuenta_Nueva.
            ' Paga en el modelo definido en la politica 1503. La version anterior paga en el modelo 7
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date
Dim Dias_Acumulados As Integer


Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim rs As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
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
    Aux_Fecha = rs!elfechahasta
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)

    Flog.writeline "Licencia desde " & rs!elfechadesde & " al " & rs!elfechahasta
End If
rs.Close

Dias_Acumulados = 0
Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        Flog.writeline "DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA"
        rs.Close
        StrSql = "DELETE FROM vacpagdesc"
        StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND (pago_dto = 3" 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

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


'FGZ - 27/07/2005
'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If

'POLITICA  - ANALISIS
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
Flog.writeline "Dias afecta " & Dias_Afecta


Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento

' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(rs!elfechadesde) > 15 Then
        'es segunda quincena
        Aux_TipDiaDescuento = st_ModeloDto2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = st_ModeloDto
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(rs!elfechahasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
Else
    Quincena_Inicio = 1
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio

If Genera_Pagos Then
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
End If

'tantos descuentos como meses afecte
Fin_Licencia = False
Mes_Afecta = Mes_Inicio
Ano_Afecta = Ano_Inicio
Dias_Pendientes = 0
Dias_Restantes = Dias_Afecta

'determinar los dias que afecta para el primer mes de decuento
'Genera 30 dias, para todos los meses
Anio_bisiesto = EsBisiesto(Ano_Afecta)

If Mes_Afecta = 2 Then
    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    If Not Jornal Then
        'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    Else
        If Quincena_Siguiente = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    End If
Else
    
    If Not Jornal Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Mes_Afecta = 12 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
        Else
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    Else
        If Quincena_Siguiente = 1 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    End If
End If

If (rs!elfechahasta <= Ultimo_Mes) Then
    'termina en el mes
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
    'continua en el mes siguiente
    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
End If


    'FGZ - 27/02/2006 No estaba inicializando estas variables y por lo tanto
    '      siempre comenzaba en la primera quincena (cuando Jornal)
    Primero_Mes = AFecha(Mes_Afecta, Day(rs!elfechadesde), Ano_Afecta)
    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1

    Aux_TipDiaPago = TipDiaPago
    Aux_TipDiaDescuento = TipDiaDescuento
    
    If Jornal Then
        'reviso a que quincena corresponde
        If Day(Primero_Mes) > 15 Then
            'es segunda quincena
            Aux_TipDiaDescuento = st_ModeloDto2
            Quincena_Siguiente = 2
        Else
            'Es primera quincena
            Aux_TipDiaDescuento = st_ModeloDto
            Quincena_Siguiente = 1
        End If
    End If
    
    
    
Do While Not (Fin_Licencia) Or Dias_Acumulados > 0
    
    If Jornal Then
        If Dias_Afecta > 15 Then
            Dias_Afecta = 15
            Dias_Acumulados = 1
        End If
        If Dias_Acumulados > 0 And Dias_Afecta < 15 Then
            If (Dias_Acumulados + Dias_Afecta) <= 15 Then
                Dias_Afecta = Dias_Afecta + Dias_Acumulados
                Dias_Acumulados = 0
            End If
        End If
    Else
        If Dias_Afecta > 30 Then
            Dias_Afecta = 30
            Dias_Acumulados = 1
        End If
        If Dias_Acumulados > 0 And Dias_Afecta < 30 Then
            If (Dias_Acumulados + Dias_Afecta) <= 30 Then
                Dias_Afecta = Dias_Afecta + Dias_Acumulados
                Dias_Acumulados = 0
            End If
        End If
        'Aux_TipDiaDescuento = 7
    End If
    
    Dias_Restantes = Dias_Restantes - Dias_Afecta
    
    If Genera_Descuentos And (Dias_Afecta <> 0) Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
    End If
            
    If Jornal Then
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
    End If
    
    'determinar a que continua en el proximo mes
    If (Mes_Afecta = 12) Then
        If Not Jornal Then
            Mes_Afecta = 1
            Ano_Afecta = Ano_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    Else
        If Not Jornal Then
            Mes_Afecta = Mes_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    End If
    
    If (Ano_Afecta = Ano_Fin) Then
        If (Mes_Afecta > Mes_Fin) Then
            Fin_Licencia = True
        Else
            If (Mes_Afecta = Mes_Fin) And Jornal Then
                If Quincena_Siguiente > Quincena_Fin Then
                    Fin_Licencia = True
                End If
            End If
        End If
    Else
        If (Ano_Afecta > Ano_Fin) Then
            Fin_Licencia = True
        End If
    End If
    
    'determinar los dias que afecta
    Anio_bisiesto = EsBisiesto(Ano_Afecta)

'    If Mes_Afecta = 2 Then
'        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'        Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
'    Else
'        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'        'ultimo_mes = AFecha(mes_afecta, 30, ano_afecta)
'        If Mes_Afecta = 12 Then
'            Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
'        Else
'            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'        End If
'    End If



    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                'Ultimo_Mes = IIf(Anio_bisiesto, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If


    If (rs!elfechahasta <= Ultimo_Mes) Then
    'termina en el mes
        If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
            Dias_Afecta = Dias_Restantes
        Else
            Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
        End If
    Else
        'continua en el mes siguiente
        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
    End If
    
    If Not Jornal Then
        If Dias_Afecta > 30 Then
           Dias_Afecta = 30
        End If
    Else
        If Dias_Afecta > 15 Then
           Dias_Afecta = 15
        End If
    End If
    
    'FGZ - 17/02/2006 - Revisar esta inicializacion (creo que no debe ir)
    'Aux_TipDiaPago = TipDiaPago
    'Aux_TipDiaDescuento = TipDiaDescuento
    
    If Jornal Then
        'reviso a que quincena corresponde
        If Day(Primero_Mes) > 15 Then
            'es segunda quincena
            'Aux_TipDiaPago = TipDiaDescuento
            
            'FGZ - 07/04/2006
            Aux_TipDiaDescuento = st_ModeloDto2
        Else
            'Es primera quincena
            Aux_TipDiaDescuento = st_ModeloDto
        End If
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

 
Function EsBisiesto(Anio As Integer) As Boolean
If (Anio Mod 4) = 0 Then
    If (((Anio Mod 100) <> 0) And ((Anio Mod 400) = 0)) Or _
        (((Anio Mod 100) = 0) And ((Anio Mod 400) = 0)) Or _
        (((Anio Mod 100) <> 0) And ((Anio Mod 400) <> 0)) Then
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


'Public Sub Politica1500PagayDescuenta()
''corresponde a vacpdo02
''***************************************************************
''PAGA Y DESCUENTA POR MES SIN TOPE DE 30 DIAS FIJOS.
''***************************************************************
'Dim rs As New Recordset
'Dim StrSql As String
'
'Dim Mes_Inicio As Integer
'Dim Ano_Inicio As Integer
'Dim Mes_Fin    As Integer
'Dim Ano_Fin    As Integer
'Dim Fin_Licencia  As Boolean
'Dim Mes_Afecta    As Integer
'Dim Ano_Afecta    As Integer
'Dim Primero_Mes   As Date
'Dim Ultimo_Mes    As Date
'Dim Dias_Afecta   As Integer
'Dim Dias_Pendientes As Integer
'Dim Dias_Restantes As Integer
'
'Dim Dias_ya_tomados As Integer
'Dim Fecha_limite    As Date
'Dim Anio_bisiesto   As Boolean
'Dim Jornal As Boolean
'Dim Legajo As Long
'Dim NroTer As Long
'Dim Nombre As String
'
''/**************************************************************************************************/
'
'Call Politica(1503)
'If Not PoliticaOK Then
'    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
'    TipDiaPago = 3
'    TipDiaDescuento = 3
'End If
'
'
''Busco la licencia dentro del intervalo especificado
'StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'OpenRecordset StrSql, rs
'
'If rs.EOF Then
'    Flog.writeline "Licencia inexistente"
'    GoTo CE
'Else
'    'NroTer = emp_lic.empleado
'    NroTer = rs!Empleado
'    Mes_Inicio = Month(rs!elfechadesde)
'    Ano_Inicio = Year(rs!elfechadesde)
'    Mes_Fin = Month(rs!elfechahasta)
'    Ano_Fin = Year(rs!elfechahasta)
'End If
'
'Genera_Pagos = False
'Genera_Descuentos = False
'Call Politica(1507)
'If Not PoliticaOK Then
'    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
'    Genera_Pagos = True
'    Genera_Descuentos = True
'End If
'
''VERIFICAR el reproceso y manejar la depuracion
'If Not Reproceso Then
'    'si no es reproceso y existe el desglose de pago/descuento, salir
'    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
'    OpenRecordset StrSql, rs
'    If Not rs.EOF Then
'        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
'        GoTo CE
'    End If
'    rs.Close
'Else
'    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
'    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
'    OpenRecordset StrSql, rs
'    If Not rs.EOF Then
'        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
'    Else
'        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
'        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
'        objConn.Execute StrSql, , adExecuteNoRecords
'    End If
'End If
'
'StrSql = "SELECT * FROM tipdia WHERE tdnro = " & TipoLicencia
'OpenRecordset StrSql, rs
'If rs.EOF Then
'    Flog.writeline "Tipo de dia de Vacaciones (2) inexistente."
'    GoTo CE
'End If
'
'
''/* POLITICA  - ANALISIS */
'Fin_Licencia = False
'Mes_Afecta = Mes_Inicio
'Ano_Afecta = Ano_Inicio
'Dias_Pendientes = 0
'
''/* determinar los dias que afecta para el primer mes de descuento */
'Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
'DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
'
'StrSql = "SELECT * FROM emp_lic "
'If TipoLicencia = 2 Then
'    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
'End If
'StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
'OpenRecordset StrSql, rs
'If TipoLicencia = 2 Then
'    NroVac = rs!vacnro
'End If
'
'
'If (rs!elfechahasta <= Ultimo_Mes) Then
''/* termina en el mes */
'    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
'Else
''/* continua en el mes siguiente */
'    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
'End If
'
'Do While Not Fin_Licencia
'    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
'    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaDescuento, Dias_Afecta, 4) 'Pago
'
'    If (Mes_Afecta = 12) Then
'        Mes_Afecta = 1
'        Ano_Afecta = Ano_Afecta + 1
'    Else
'        Mes_Afecta = Mes_Afecta + 1
'    End If
'
'    If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
'        Fin_Licencia = True
'    End If
'
'    '/* determinar los d­as que afecta */
'    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'    Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
'    DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
'
'    If (rs!elfechahasta <= Ultimo_Mes) Then
'    '/* termina en el mes */
'        Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
'    Else
'    '/* continua en el mes siguiente */
'        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
'    End If
'
'Loop
'    GoTo ProcesadoOK
'CE:
'    Flog.writeline " ------------------------------------------------------------"
'    Flog.writeline "Error procesando Empleado:" & Ternro
'    Flog.writeline Err.Description
'    Flog.writeline "SQL: " & StrSql
'    Flog.writeline " ------------------------------------------------------------"
'
'ProcesadoOK:
'    rs.Close
'    Set rs = Nothing
'End Sub


Public Sub Politica1500NoLiquida()
' corresponde a vacpdo03
'***************************************************************
' NO LIQUIDA:Pago adelantado entero: mes anterior al inicio de las vacaciones.
' NO LIQUIDA:Descuento adelantado entero: mes de inicio de las vacaciones.
' PAGO adelantado entero: mes de inicio de las vacaciones.
' DESCUENTO x mes tomado con tope de 30 dias mensuales fijos: mes de inicio de las vacaciones.
'
' Nota: Para que liquide tocar el licdes.p y licpag.p que esta en /par.
'***************************************************************
Dim rs As New Recordset
Dim StrSql As String

Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer

Dim NroTer As Long

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline Espacios(Tabulador * 1) & "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
' Si no hay licencia me voy
If rs.EOF Then
    Flog.writeline Espacios(Tabulador * 1) & "Licencia inexistente.", vbCritical
    GoTo CE
Else
    NroTer = rs!Empleado
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline Espacios(Tabulador * 1) & "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'EAM- Verifica si existe el tipo de licencia
If Not ExisteTipoLicencia(2) Then
    Flog.writeline Espacios(Tabulador * 1) & "Tipo de dia de Vacaciones (2) inexistente."
    GoTo CE
End If

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

'/* generar pago de adelanto */
If (Mes_Inicio = 1) Then
    Call Generar_PagoDescuento(12, (Ano_Inicio - 1), TipDiaPago, Dias_Afecta, 1) 'Pago
Else
    Call Generar_PagoDescuento((Mes_Inicio - 1), Ano_Inicio, TipDiaPago, Dias_Afecta, 1) 'Pago
End If

'/* GENERAR DESCUENTO DE ANTICIPO */
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaDescuento, Dias_Afecta, 2) 'Pago
          
'/* GENERAR EL PAGO DE VACACIONES */
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'PAgo

Fin_Licencia = False
Mes_Afecta = Mes_Inicio
Ano_Afecta = Ano_Inicio

'determinar los días que afecta para el primer mes de descuento
Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))

'Revisar comparacion de fechas
If (rs!elfechahasta <= Ultimo_Mes) Then
    'termina en el mes
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
    'continua en el mes siguiente */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
End If

Do While Not Fin_Licencia
    If Dias_Afecta > 30 Then
        Dias_Pendientes = Dias_Afecta - 30
        Dias_Afecta = 30
    End If
    
    Call Generar_PagoDescuento((Mes_Inicio - 1), Ano_Inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
    'IF RETURN-VALUE <> ""
        'THEN UNDO main, RETURN.
    If (Mes_Afecta = 12) Then
        Mes_Afecta = 1
        Ano_Afecta = Ano_Afecta + 1
    Else
        Mes_Afecta = Mes_Afecta + 1
    End If
    If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
        Fin_Licencia = True
    End If

    '/* determinar los días que afecta */
    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
    DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
       
    If (rs!elfechahasta <= Ultimo_Mes) Then
        '/* termina en el mes */
        Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1 + Dias_Pendientes
    Else
    '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1 + Dias_Pendientes
    End If

Loop

GoTo ProcesadoOK
CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing
End Sub



Public Sub Politica1500AdelantaDescuentaTodo()
'***************************************************************
' paga adelantado, descuenta adelantado todo los días de la licencia.
'***************************************************************
Dim rs As New Recordset
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin As Integer
Dim Ano_Fin As Integer
Dim Dias_Afecta   As Integer
Dim NroTer As Long


On Error GoTo CE

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs

' Si no hay licencia me voy
If rs.EOF Then
    Flog.wirteline Espacios(Tabulador * 1) & "Licencia inexistente."
    GoTo CE
Else
    NroTer = rs!Empleado
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'Si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'Verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'EAM- Verifica si existe el tipo de licencia
If Not ExisteTipoLicencia(2) Then
    Flog.writeline Espacios(Tabulador * 1) & "Tipo de dia de Vacaciones (2) inexistente."
    GoTo CE
End If

'/* POLITICA  - sin ANALISIS */
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

'Paga y descuenta el total que de dias afectados. (Dias de Licencia)
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
GoTo ProcesadoOK

CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing
End Sub




Public Sub Politica1500PagaDescuentaTodo()
'Corresponde a vacpdo05.p
'/* ***************************************************************
'El paga y descuenta el total de la licencia mes de goce de la misma, sin tope mensual
'Sólo genera pagos y descuentos 1 vez por período. Si se vuelve a procesar para pagar una nueva licencia, y ya hay una paga, no se generarán nuevos pagos/descuentos a menos que se reprocese.
'******************************************************************
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Dias_Afecta   As Integer
Dim Jornal As Boolean
Dim NroTer As Long
Dim ya_se_pago As Boolean
Dim Aux_Fecha As Date
Dim blv As New ADODB.Recordset
Dim rs As New ADODB.Recordset

On Error GoTo CE

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs

' Si no hay licencia me voy
If rs.EOF Then
    Flog.wirteline Espacios(Tabulador * 1) & "Licencia inexistente." & nrolicencia
    GoTo CE
Else
    Aux_Fecha = rs!elfechahasta
    NroTer = rs!Empleado
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline Espacios(Tabulador * 1) & "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'EAM- Verifica si existe el tipo de licencia
If Not ExisteTipoLicencia(2) Then
    Flog.writeline Espacios(Tabulador * 1) & "Tipo de dia de Vacaciones (2) inexistente."
    GoTo CE
End If

'EAM- Busca la forma de liquidacion True-->Jornal | false -> Mensual
'Jornal = FormaDeLiquidacion(Aux_Fecha, 22)


'BUSCAR EL PERIODO DE LA LICENCIA
StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If Not rs.EOF Then
    StrSql = "SELECT * FROM vacacion WHERE vacnro= " & rs!vacnro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        NroVac = rs!vacnro
    End If
End If
ya_se_pago = False

StrSql = "SELECT * FROM lic_vacacion " & _
        " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro " & _
        " WHERE emp_lic.empleado = " & NroTer & " AND vacnro= " & NroVac
OpenRecordset StrSql, blv
'Recorre todas las licencias en el período de vacacion
Do While Not blv.EOF
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro= " & blv!emp_licnro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        ya_se_pago = True
        Flog.writeline Espacios(Tabulador * 1) & "Ya se pagó"
        GoTo CE
    End If
    blv.MoveNext
Loop


If Not ya_se_pago Then
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

    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
End If

GoTo ProcesadoOK

CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing
    Set blv = Nothing
End Sub

Public Sub Politica1500v_7()
'***************************************************************
' En cuanto se toma al menos un día de vacaciones, se le paga todos los días correspondientes y descuenta los dias correspondientes
' a partir de la fecha de inicio de la licencia.
'***************************************************************
Dim rs As New Recordset
Dim StrSql As String

Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Restantes As Integer
Dim fechaAux As Date

Dim Jornal As Boolean
Dim NroTer As Long

Dim ya_se_pago As Boolean
Dim cantdias As Integer
Dim blv As New ADODB.Recordset

On Error GoTo CE

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT Empleado,elfechadesde,elfechahasta,vacnro FROM emp_lic " & _
        " INNER JOIN lic_vacacion ON emp_lic.emp_licnro= lic_vacacion.emp_licnro" & _
        " WHERE lic_vacacion.emp_licnro = " & nrolicencia
OpenRecordset StrSql, rs

' Si no hay licencia me voy
If rs.EOF Then
    Flog.wirteline Espacios(Tabulador * 1) & "Licencia inexistente." & nrolicencia
    GoTo CE
Else
    NroTer = rs!Empleado
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
    NroVac = rs!vacnro
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


'EAM- Verifica si existe el tipo de licencia
If Not ExisteTipoLicencia(2) Then
    Flog.writeline Espacios(Tabulador * 1) & "Tipo de dia de Vacaciones (2) inexistente."
    GoTo CE
End If


ya_se_pago = False
'Busco todas las licencias del empleado en el periodo
StrSql = "SELECT * FROM lic_vacacion" & _
        " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro" & _
        " WHERE emp_lic.empleado = " & NroTer & " AND vacnro= " & NroVac
OpenRecordset StrSql, blv


Do While Not blv.EOF
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro= " & blv!emp_licnro
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        ya_se_pago = True
        Exit Do
    End If
    blv.MoveNext
Loop

'EAM- Si cantidad es uno quiere decir que todavía no genero el pago. Se verifica porque cicla tantas veces como licencias encontradas
If CantidadLicenciasProcesadas = 1 Then
    'EAM- Obtiene la cantidad de días correspondientes
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & NroVac & " AND ternro = " & NroTer & " AND (venc is null OR venc = 0)"
    OpenRecordset StrSql, rs
    cantdias = rs!vdiascorcant
   
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, cantdias, 3) 'Pago

    StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
    OpenRecordset StrSql, rs
    
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Restantes = cantdias
    fechaAux = rs!elfechadesde
    

    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
    DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))

    If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
        Dias_Afecta = Dias_Restantes
        fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
    Else
        'Continua en el mes siguiente
        If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
            Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
            fechaAux = Ultimo_Mes
        Else
            Dias_Afecta = 30
            fechaAux = Ultimo_Mes
        End If
    End If

    
    Do While Not Fin_Licencia
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
        
        '/* determinar a que continua en el proximo mes */
        If (Mes_Afecta = 12) Then
            Mes_Afecta = 1
            Ano_Afecta = Ano_Afecta + 1
        Else
            Mes_Afecta = Mes_Afecta + 1
        End If

        If (Dias_Restantes <= 0) Then
            Fin_Licencia = True
        End If

        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
        DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))


        If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
            Dias_Afecta = Dias_Restantes
            fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
        Else
            'Continua en el mes siguiente
            If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
                Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
                fechaAux = Ultimo_Mes
            Else
                Dias_Afecta = 30
                fechaAux = Ultimo_Mes
            End If
            
        End If


    Loop
End If

GoTo ProcesadoOK

CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing
    Set blv = Nothing

End Sub


'Public Sub Politica1500v_7()
''Corresponde a vacpdo06.p
''***************************************************************
''  EN CUANTO SE TOMA AL MENOS UN DIA DE VACACIONES, SE LE PAGA Y DESCUENTA TODOS LOS D™AS DE
''  VACACIONES QUE LE CORRESPONDEN PARA EL A¾O
''  TODO ESTO LO HACE MIENTRAS NO PAGO NI DESCONTO NINGUN DIA DE DICHA LICENCIA
''***************************************************************
'Dim rs As New Recordset
'Dim StrSql As String
'
'Dim Mes_Inicio As Integer
'Dim Ano_Inicio As Integer
'Dim Mes_Fin    As Integer
'Dim Ano_Fin    As Integer
'Dim Fin_Licencia  As Boolean
'Dim Mes_Afecta    As Integer
'Dim Ano_Afecta    As Integer
'Dim Primero_Mes   As Date
'Dim Ultimo_Mes    As Date
'Dim Dias_Afecta   As Integer
''Dim Dias_Pendientes As Integer
'Dim Dias_Restantes As Integer
'
'Dim Anio_bisiesto   As Boolean
'Dim Jornal As Boolean
'Dim NroTer As Long
'
'Dim ya_se_pago As Boolean
'Dim cantdias As Integer
'Dim blv As New ADODB.Recordset
'
'On Error GoTo CE
'
'Call Politica(1503)
'If Not PoliticaOK Then
'    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
'    TipDiaPago = 3
'    TipDiaDescuento = 3
'End If
'
''Busco la licencia dentro del intervalo especificado
'StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'OpenRecordset StrSql, rs
'
'' Si no hay licencia me voy
'If rs.EOF Then
'    Flog.wirteline Espacios(Tabulador * 1) & "Licencia inexistente." & nrolicencia
'    GoTo CE
'Else
'    NroTer = rs!Empleado
'    Mes_Inicio = Month(rs!elfechadesde)
'    Ano_Inicio = Year(rs!elfechadesde)
'    Mes_Fin = Month(rs!elfechahasta)
'    Ano_Fin = Year(rs!elfechahasta)
'End If
'
'
'Genera_Pagos = False
'Genera_Descuentos = False
'Call Politica(1507)
'If Not PoliticaOK Then
'    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
'    Genera_Pagos = True
'    Genera_Descuentos = True
'End If
'
''VERIFICAR el reproceso y manejar la depuracion
'If Not Reproceso Then
'    'si no es reproceso y existe el desglose de pago/descuento, salir
'    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
'    OpenRecordset StrSql, rs
'    If Not rs.EOF Then
'        Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
'        GoTo CE
'    End If
'Else
'    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
'    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
'    OpenRecordset StrSql, rs
'    If Not rs.EOF Then
'        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
'        GoTo CE
'    Else
'        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
'        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
'        objConn.Execute StrSql, , adExecuteNoRecords
'    End If
'End If
'
'
''EAM- Verifica si existe el tipo de licencia
'If Not ExisteTipoLicencia(2) Then
'    Flog.writeline Espacios(Tabulador * 1) & "Tipo de dia de Vacaciones (2) inexistente."
'    GoTo CE
'End If
'
'
''/* BUSCAR EL PERIODO DE LA LICENCIA */
''       FIND lic_vacacion OF emp_lic NO-LOCK.
'StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
'OpenRecordset StrSql, rs
'
'If Not rs.EOF Then
'    NroVac = rs!vacnro
'Else
'    GoTo CE
'End If
'
'ya_se_pago = False
'
''Busco todas las licencias del empleado en el periodo
'StrSql = "SELECT * FROM lic_vacacion" & _
'        " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro" & _
'        " WHERE emp_lic.empleado = " & NroTer & " AND vacnro= " & NroVac
'OpenRecordset StrSql, blv
'
'
'Do While Not blv.EOF
'    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro= " & blv!emp_licnro
'    OpenRecordset StrSql, rs
'    If Not rs.EOF Then
'        ya_se_pago = True
'        Exit Do
'    End If
'    blv.MoveNext
'Loop
'
''EAM- Si cantidad es uno quiere decir que todavía no genero el pago. Se verifica porque cicla tantas veces como licencias encontradas
'If CantidadLicenciasProcesadas = 1 Then
''    StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
''    OpenRecordset StrSql, rs
'
''    StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
''    OpenRecordset StrSql, rs
'
'    StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & NroVac & " AND ternro = " & NroTer & " AND (venc is null OR venc = 0)"
'    OpenRecordset StrSql, rs
'    cantdias = rs!vdiascorcant
'
'    'NroVac = rs!vacnro
'
'    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, cantdias, 3) 'Pago
''    If Mes_Inicio = 1 Then
''        'Call generar_pago(12, ano_inicio - 1, CantDias, 3, Jornal, nrolicencia)
''        Call Generar_PagoDescuento(12, Ano_Inicio - 1, TipDiaPago, Dias_Afecta, 3) 'PAgo
''    Else
''        'Call generar_pago(mes_inicio - 1, ano_inicio, CantDias, 3, Jornal, nrolicencia)
''        Call Generar_PagoDescuento(Mes_Inicio - 1, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'PAgo
''    End If
'End If
'
'    StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'    OpenRecordset StrSql, rs
'
'    Fin_Licencia = False
'    Mes_Afecta = Mes_Inicio
'    Ano_Afecta = Ano_Inicio
'    'Dias_Pendientes = 0
'    'Dias_Restantes = rs!elcantdias
'    Dias_Restantes = cantdias
'
'    Anio_bisiesto = EsBisiesto(Ano_Afecta)
'
'    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'    Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
'    DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
'
'    If (rs!elfechahasta <= Ultimo_Mes) Then
'        Dias_Afecta = rs!elcantdias
'    Else
'        '/* continua en el mes siguiente */
'        Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
'    End If
'
'    Do While Not Fin_Licencia
'        Dias_Restantes = Dias_Restantes - Dias_Afecta
'        'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
'        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
'
'        '/* determinar a que continua en el proximo mes */
'        If (Mes_Afecta = 12) Then
'            Mes_Afecta = 1
'            Ano_Afecta = Ano_Afecta + 1
'        Else
'            Mes_Afecta = Mes_Afecta + 1
'        End If
'
'        If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
'            Fin_Licencia = True
'        End If
'
'        '/* determinar los d¡as que afecta */
'        Anio_bisiesto = EsBisiesto(Ano_Afecta)
'
'        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'        Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
'        DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
'
'        If (rs!elfechahasta <= Ultimo_Mes) Then
'        '/* termina en el mes */
'            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
'                Dias_Afecta = Dias_Restantes
'            Else
'                Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
'            End If
'        Else
'        '/* continua en el mes siguiente */
'            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
'        End If
'
'    Loop
'
'GoTo ProcesadoOK
'
'CE:
'    Flog.writeline " ------------------------------------------------------------"
'    Flog.writeline "Error procesando Empleado:" & Ternro
'    Flog.writeline Err.Description
'    Flog.writeline "SQL: " & StrSql
'    Flog.writeline " ------------------------------------------------------------"
'
'ProcesadoOK:
'    rs.Close
'    Set rs = Nothing
'    Set blv = Nothing
'
'End Sub

'Public Sub Politica1500v_13()
''Corresponde a vacpdo06.p
''***************************************************************
''  EN CUANTO SE TOMA AL MENOS UN DIA DE VACACIONES, SE LE PAGA Y DESCUENTA TODOS LOS D™AS DE
''  VACACIONES QUE LE CORRESPONDEN PARA EL A¾O
''  TODO ESTO LO HACE MIENTRAS NO PAGO NI DESCONTO NINGUN DIA DE DICHA LICENCIA
''***************************************************************
'' FAF - 11/08/2006 - Es copia de Politica1500v_7() con la unica diferencia que topea a 30 dias
'Dim rs As New Recordset
'Dim StrSql As String
'
'Dim Mes_Inicio As Integer
'Dim Ano_Inicio As Integer
'Dim Mes_Fin    As Integer
'Dim Ano_Fin    As Integer
'Dim Fin_Licencia  As Boolean
'Dim Mes_Afecta    As Integer
'Dim Ano_Afecta    As Integer
'Dim Primero_Mes   As Date
'Dim Ultimo_Mes    As Date
'Dim Dias_Afecta   As Integer
'Dim Dias_Pendientes As Integer
'Dim Dias_Restantes As Integer
'
'Dim Dias_ya_tomados As Integer
'Dim Fecha_limite    As Date
'Dim Anio_bisiesto   As Boolean
'Dim Jornal As Boolean
'Dim Legajo As Long
'Dim NroTer As Long
'Dim Nombre As String
'
'Dim ya_se_pago As Boolean
'Dim cantdias As Integer
'Dim blv As New ADODB.Recordset
'Dim bel As New ADODB.Recordset
'
'Call Politica(1503)
'If Not PoliticaOK Then
'    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
'    TipDiaPago = 3
'    TipDiaDescuento = 3
'    'Exit Sub
'End If
'
''Busco la licencia dentro del intervalo especificado
'StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'rs.Open StrSql, objConn
'' Si no hay licencia me voy
'If rs.EOF Then
'    'MsgBox "Licencia inexistente.", vbCritical
'    rs.Close
'    Set rs = Nothing
'    Exit Sub
'Else
'    NroTer = rs!Empleado
'    Mes_Inicio = Month(rs!elfechadesde)
'    Ano_Inicio = Year(rs!elfechadesde)
'    Mes_Fin = Month(rs!elfechahasta)
'    Ano_Fin = Year(rs!elfechahasta)
'End If
'rs.Close
'
'Genera_Pagos = False
'Genera_Descuentos = False
'Call Politica(1507)
'If Not PoliticaOK Then
'    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
'    Genera_Pagos = True
'    Genera_Descuentos = True
'End If
'
''VERIFICAR el reproceso y manejar la depuracion
'If Not Reproceso Then
'    'si no es reproceso y existe el desglose de pago/descuento, salir
'    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
'    rs.Open StrSql, objConn
'    If Not rs.EOF Then
'        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
'        rs.Close
'        Set rs = Nothing
'        Exit Sub
'    End If
'    rs.Close
'Else
'    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
'    StrSql = "SELECT * FROM vacpagdesc "
'    StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
'    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
'    rs.Open StrSql, objConn
'    If Not rs.EOF Then
'        rs.Close
'        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
'    Else
'        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
'        rs.Close
'        StrSql = "DELETE FROM vacpagdesc"
'        StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
'        If Genera_Pagos And Genera_Descuentos Then
'            StrSql = StrSql & " AND (pago_dto = 3" 'pagos
'            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
'        Else
'            If Genera_Pagos Then
'                StrSql = StrSql & " AND pago_dto = 3" 'pagos
'            End If
'            If Genera_Descuentos Then
'                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
'            End If
'        End If
'        objConn.Execute StrSql, , adExecuteNoRecords
'    End If
'End If
'
'
'
'
''/* POLITICA  - ANALISIS */
'StrSql = "SELECT * FROM emp_lic "
'If TipoLicencia = 2 Then
'    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
'End If
'StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
'OpenRecordset StrSql, rs
'If TipoLicencia = 2 Then
'    NroVac = rs!vacnro
'End If
'Dias_Afecta = rs!elcantdias
''Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
'
''Terminar
''RUN generar_pago(mes_inicio, ano_inicio, dias_afecta, 3)
''Call generar_pago(mes_inicio, ano_inicio, Dias_Afecta, 3, Jornal, nrolicencia)
''Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
''fgz - 09/01/2003
'
''Call Generar_PagoDescuento(primer_mes, primer_ano, TipDiaPago, Dias_Afecta, 3) 'Pago
'If Genera_Pagos Then
'    If Pliq_Nro <> 0 Then
'        Mes_Inicio = Pliq_Mes
'        Ano_Inicio = Pliq_Anio
'    End If
'    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
'End If
'rs.Close
'
'
'
'
'
'
'
''StrSql = "SELECT * FROM tipdia WHERE tdnro = 2"
''rs.Open StrSql, objConn
''If rs.EOF Then
'    'MsgBox " Tipo de dia de Vacaciones (2) inexistente."
''    rs.Close
''    Set rs = Nothing
''    Exit Sub
''Else
''    rs.Close
''End If
'
''StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
''rs.Open StrSql, objConn
''rs.Close
'
''/* BUSCAR EL PERIODO DE LA LICENCIA */
''       FIND lic_vacacion OF emp_lic NO-LOCK.
''StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
''rs.Open StrSql, objConn
'
''       FIND vacacion OF lic_vacacion NO-LOCK.
''StrSql = "SELECT * FROM vacacion WHERE vacnro= " & rs!vacnro
''rs.Close
''rs.Open StrSql, objConn
'
''ya_se_pago = False
'
''Busco todas las licencias del empleado en el periodo
''StrSql = "SELECT * FROM lic_vacacion" & _
''" INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro" & _
''" WHERE emp_lic.empleado = " & NroTer & " AND vacnro= " & rs!vacnro
''rs.Close
''blv.Open StrSql, objConn
''Do While Not blv.EOF
''    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro= " & blv!emp_licnro
''    rs.Open StrSql, objConn
''    If Not rs.EOF Then
''        ya_se_pago = True
''        rs.Close
''        Exit Do
''    End If
''    blv.MoveNext
''    rs.Close
''Loop
'
''blv.Close
'
''If CantidadLicenciasProcesadas = 1 Then
'
''    StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
''    rs.Open StrSql, objConn
'
''    StrSql = "SELECT * FROM vacacion WHERE vacnro = " & rs!vacnro
''    rs.Close
'
''    rs.Open StrSql, objConn
'
''    StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & rs!vacnro & _
''    " AND ternro = " & NroTer
''    rs.Close
'
''    rs.Open StrSql, objConn
''    cantdias = rs!vdiascorcant
'
''    NroVac = rs!vacnro
''    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, cantdias, 3) 'PAgo
'
''    rs.Close
'
''End If
'
'    StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'    rs.Open StrSql, objConn
'
'    Fin_Licencia = False
'    Mes_Afecta = Mes_Inicio
'    Ano_Afecta = Ano_Inicio
'    Dias_Pendientes = 0
'    Dias_Restantes = rs!elcantdias
'
'    Anio_bisiesto = EsBisiesto(Ano_Afecta)
'
'    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
''   Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
''   DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
'
'    If Mes_Afecta = 2 Then
'        If Anio_bisiesto Then
'            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
'        Else
'            Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
'        End If
'    Else
'        Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
'    End If
'
'    If (rs!elfechahasta <= Ultimo_Mes) Then
'        '/* termina en el mes */
'        Dias_Afecta = rs!elcantdias
'    Else
'        '/* continua en el mes siguiente */
'        Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
'    End If
'
'    Do While Not Fin_Licencia
'
'        Dias_Restantes = Dias_Restantes - Dias_Afecta
'
'        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'PAgo
'
'        '/* determinar a que continua en el proximo mes */
'        If (Mes_Afecta = 12) Then
'            Mes_Afecta = 1
'            Ano_Afecta = Ano_Afecta + 1
'        Else
'            Mes_Afecta = Mes_Afecta + 1
'        End If
'
'        If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
'            Fin_Licencia = True
'        End If
'
'        '/* determinar los d¡as que afecta */
'
'        Anio_bisiesto = EsBisiesto(Ano_Afecta)
'
'        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'
'        If Mes_Afecta = 2 Then
'            If Anio_bisiesto Then
'                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
'            Else
'                Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
'            End If
'        Else
'            Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
'        End If
'
'        If (rs!elfechahasta <= Ultimo_Mes) Then
'        '/* termina en el mes */
'            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
'                Dias_Afecta = Dias_Restantes
'            Else
'                Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
'            End If
'        Else
'        '/* continua en el mes siguiente */
'            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
'        End If
'
'    Loop
'
'If rs.State = adStateOpen Then rs.Close
'Set rs = Nothing
'
'End Sub

Public Sub Politica1500_v2AdelantaDescuenta()
'-----------------------------------------------------------------------
'vacpdo50.p
'Genera Pago/dto por dias correspondientes
'paga adelantado todo y descuenta por mes lo que corresponde generar todos los dias de pago
'Fecha:
'Autor: FGZ
'-----------------------------------------------------------------------
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim Jornal As Boolean
Dim Dias_Afecta As Integer
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin As Integer
Dim Ano_Fin As Integer
Dim Aux_TotalDiasCorrespondientes As Long

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer


StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones " & NroVac
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If

'Mes_Inicio = Month(rs_Vacacion!vacfecdesde)
'Ano_Inicio = Year(rs_Vacacion!vacfecdesde)
'Mes_Fin = Month(rs_Vacacion!vacfechasta)
'Ano_Fin = Year(rs_Vacacion!vacfechasta)


Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If


'FGZ - 14/10/2005
'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
    
'Busco la forma de liquidacion
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
        Else
            Jornal = False
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
    End If
Else
    'Flog.writeline "No se encuentra estructura de forma de Liquidacion del empleado"
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
    'Exit Sub
End If

PoliticaOK = False
Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaPago = TipDiaDescuento
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = TipDiaPago
    End If
End If


If Pliq_Nro <> 0 Then
    Mes_Inicio = Pliq_Mes
    Ano_Inicio = Pliq_Anio
End If
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

If Genera_Dto_DiasCorr Then
    
    'ojo que el dto si es jornal ==> se parte
    
    If Pliq_Nro <> 0 Then
        Mes_Inicio = Pliq_Mes
        Ano_Inicio = Pliq_Anio
    End If
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento

Else
    Flog.writeline "No se generan los descuentos."
End If


' Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
End Sub

Public Sub Politica1500_v3PagaDescuenta()
'-----------------------------------------------------------------------
'vacpdo50.p
'Parecida a Politica1500_v2AdelantaDescuenta, solo que paga y descuenta todo en funcion al primer dia de licencia la cantidad de dias que corresponda
'Fecha:
'Autor: FGZ
'-----------------------------------------------------------------------
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim Jornal As Boolean
Dim Dias_Afecta As Integer
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin As Integer
Dim Ano_Fin As Integer
Dim Aux_Fecha As Date
Dim NroTer As Long
Dim Aux_TotalDiasCorrespondientes As Long

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer
Dim continuar As Boolean 'mdf-

continuar = True

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs

' Si no hay licencia me voy
If rs.EOF Then
    Flog.wirteline Espacios(Tabulador * 1) & "Licencia inexistente." & nrolicencia
    Exit Sub
Else
    Aux_Fecha = rs!elfechahasta
    NroTer = rs!Empleado
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
    fecha_hasta = rs!elfechahasta
    fecha_desde = rs!elfechadesde
End If
rs.Close


StrSql = "SELECT * FROM vacacion "
StrSql = StrSql & " WHERE vacfecdesde <= " & ConvFecha(fecha_hasta)
StrSql = StrSql & " AND  vacfechasta >= " & ConvFecha(fecha_desde)
StrSql = StrSql & "and not vacnro in (" & listapgdto & ") and vacestado = -1 " 'mdf, control, cuando mas de un periodo se superpone puede analizar mas de una vez un mismo periodo
StrSql = StrSql & " ORDER BY vacnro"
OpenRecordset StrSql, rs_Periodos_Vac
Do While Not rs_Periodos_Vac.EOF And continuar
            If Not rs_Periodos_Vac.EOF Then
               ' rs_Periodos_Vac.MoveFirst
                Flog.writeline "Periodo " & rs_Periodos_Vac!vacnro
                NroVac = rs_Periodos_Vac!vacnro
                listapgdto = listapgdto & "," & NroVac
            Else
                Flog.writeline "No se encontraron Periodos de Vacaciones entre las Fechas Desde " & fecha_desde & " y Hasta " & fecha_hasta
            End If
            
            StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
            OpenRecordset StrSql, rs_vacacion
            If rs_vacacion.EOF Then
                Flog.writeline "No existe el periodo de vacaciones " & NroVac
                'Exit Sub    'mirar aca-mdf
                 GoTo avanzar
            End If
            
            StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
            StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
            StrSql = StrSql & " AND (venc is null OR venc = 0)"
            OpenRecordset StrSql, rs_vacdiascor
            If rs_vacdiascor.EOF Then
                Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
               ' Exit Sub    'mirar aca mdf
               GoTo avanzar
            End If
            
            'Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
            'Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
            'If Not EsNulo(fecha_hasta) Then
            '    Mes_Fin = Month(fecha_hasta)
            '    Ano_Fin = Year(fecha_hasta)
            'Else
            '    Mes_Fin = Month(fecha_desde)
            '    Ano_Fin = Year(fecha_desde)
            'End If
            
            
            'FGZ - 14/10/2005
            'Reviso que la cantidad de dias correspondientes del periodo
            'no supere la cantidad de dias que quedan por tomar
            Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
            'If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
            '    Dias_Afecta = Total_Dias_A_Generar
            'Else
                Dias_Afecta = rs_vacdiascor!vdiascorcant
            'End If
            'Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
            Total_Dias_A_Generar = Dias_Afecta
                
            'Busco la forma de liquidacion
            StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
            StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
            StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
            StrSql = StrSql & " his_estructura.tenro = 22 AND "
            StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
            StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                If Not IsNull(rs_Estructura!estrcodext) Then
                    If rs_Estructura!estrcodext = "2" Then
                        Jornal = True
                    Else
                        Jornal = False
                    End If
                Else
                    Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
                    Jornal = False
                End If
            Else
                'Flog.writeline "No se encuentra estructura de forma de Liquidacion del empleado"
                Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
                Jornal = False
                Flog.writeline "Mensual"
                'Exit Sub
            End If
            
            PoliticaOK = False
            Call Politica(1503)
            If Not PoliticaOK Then
                Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
                TipDiaPago = 3
                TipDiaDescuento = 3
                'Exit Sub
            End If
            
            
            Genera_Pagos = False
            Genera_Descuentos = False
            Call Politica(1507)
            If Not PoliticaOK Then
                Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
                Genera_Pagos = True
                Genera_Descuentos = True
            End If
            
            'VERIFICAR el reproceso y manejar la depuracion
            If Not Reproceso Then
                'si no es reproceso y existe el desglose de pago/descuento, salir
                StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
                StrSql = StrSql & " AND ternro =" & Ternro
                rs.Open StrSql, objConn
                If Not rs.EOF Then
                    If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
                        Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
                        'Restauro la cantidad de dias porque no se generaron
                        Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
                       ' Exit Sub 'mirar aca mdf
                        rs.Close
                       GoTo avanzar
                    Else
                        Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
                    End If
                End If
                rs.Close
            Else
                'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
                StrSql = "SELECT * FROM vacpagdesc "
                StrSql = StrSql & " WHERE vacnro = " & NroVac
                StrSql = StrSql & " AND ternro =" & Ternro
                StrSql = StrSql & " AND vacpagdesc.pronro is not null"
                rs.Open StrSql, objConn
                If Not rs.EOF Then
                    Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
                    rs.Close
                   ' Exit Sub 'mirar aca mdf
                   GoTo avanzar
                Else
                    'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
                    rs.Close
                    StrSql = "DELETE vacpagdesc "
                    StrSql = StrSql & " WHERE vacnro = " & NroVac
                    StrSql = StrSql & " AND ternro =" & Ternro
                    If Genera_Pagos And Genera_Descuentos Then
                        StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
                        StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
                    Else
                        If Genera_Pagos Then
                            StrSql = StrSql & " AND pago_dto = 3" 'pagos
                        End If
                        If Genera_Descuentos Then
                            StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
                        End If
                    End If
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
            
            
            ' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
            ' si es jornal
            ' TipDiaPago tiene el tprocnro de la primer quincena y
            ' TipDiaDescuento tiene el tprocnro de la segunda quincena
            Aux_TipDiaPago = TipDiaPago
            Aux_TipDiaDescuento = TipDiaDescuento
            If Jornal Then
                'reviso a que quincena corresponde
                If Day(Aux_Generar_Fecha_Desde) > 15 Then
                    'es segunda quincena
                    Aux_TipDiaPago = TipDiaDescuento
                Else
                    'Es primera quincena
                    Aux_TipDiaDescuento = TipDiaPago
                End If
            End If
            
            
            If Pliq_Nro <> 0 Then
                Mes_Inicio = Pliq_Mes
                Ano_Inicio = Pliq_Anio
            End If
            Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
            'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
            Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
            
            Genera_Dto_DiasCorr = False
            Call Politica(1506)
            If Not PoliticaOK Then
                Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
            End If
            
            If Genera_Dto_DiasCorr Then
                
                'ojo que el dto si es jornal ==> se parte
                
                If Pliq_Nro <> 0 Then
                    Mes_Inicio = Pliq_Mes
                    Ano_Inicio = Pliq_Anio
                End If
                Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            
            Else
                Flog.writeline "No se generan los descuentos."
            End If
            continuar = False 'llegue hasta el final y pude generar pagos y descuentos
            
            
avanzar:
       rs_Periodos_Vac.MoveNext
   Loop

' Cierro todo

rs_Periodos_Vac.Close

If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
End Sub



Private Sub Generar_PagoDescuento(ByVal mes_aplicar As Integer, ByVal ano_aplicar As Integer, ByVal tipdia As Integer, ByVal Dias_Afecta As Integer, ByVal anti_vac As Integer)
Dim rs As New Recordset
Dim StrSql As String
Dim TipoDia As Integer
Dim PliqNro As Long

Select Case anti_vac
Case 1, 3:
    Flog.writeline "Pago Modelo " & tipdia & " por " & Dias_Afecta & " dias, para año " & ano_aplicar & " mes " & mes_aplicar
Case 2, 4:
    Flog.writeline "Descuento Modelo " & tipdia & " por " & Dias_Afecta & " dias, para año " & ano_aplicar & " mes " & mes_aplicar
Case Else
End Select

If ((anti_vac = 1 Or anti_vac = 3) And Genera_Pagos) Or ((anti_vac = 2 Or anti_vac = 4) And Genera_Descuentos) Then

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
                 tipdia & "," & _
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
                 tipdia & "," & _
                 anti_vac & "," & _
                 NroVac & "," & _
                 Dias_Afecta & "," & _
                 "0" & _
                 ")"
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
End If

End Sub

Public Sub InsertarSituacionRevista(Ternro, fechadesde, FechaHasta)
    'FechaDesde,Ternro, FechaHasta
    Dim rs As New Recordset
    Dim rs_Est As New Recordset
    Dim Estrnro_SitRev As String
    
    StrSql = "SELECT estrnro, tdnro FROM csijp_srtd "
    StrSql = StrSql & " WHERE tdnro = 2"
    If rs.State = adStateOpen Then rs.Close
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Estrnro_SitRev = rs!estrnro
    End If
    If rs.State = adStateOpen Then rs.Close
    
    If Trim(Estrnro_SitRev) <> "" Then
    
        'Busco el tipo de la situacion de revista anterior
        StrSql = "SELECT * FROM his_estructura "
        StrSql = StrSql & " WHERE tenro   = 30 "
        StrSql = StrSql & " AND   ternro  = " & Ternro
        StrSql = StrSql & " AND   htetdesde <= " & ConvFecha(fechadesde)
        StrSql = StrSql & " AND   (htethasta >= " & ConvFecha(fechadesde)
        StrSql = StrSql & " OR   htethasta  is null) "
        If rs_Est.State = adStateOpen Then rs_Est.Close
        OpenRecordset StrSql, rs_Est
        If Not rs_Est.EOF Then
            'la cierro un dia antes
            If EsNulo(rs_Est!htethasta) Then
                If Not (rs_Est!htetdesde = fechadesde) Then
                    StrSql = " UPDATE his_estructura SET "
                    StrSql = StrSql & " htethasta = " & ConvFecha(CDate(fechadesde - 1))
                    StrSql = StrSql & " WHERE tenro   = 30 "
                    StrSql = StrSql & " AND   estrnro  = " & rs_Est!estrnro
                    StrSql = StrSql & " AND   ternro  = " & Ternro
                    StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                    StrSql = StrSql & " AND   htethasta  is null "
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    'la borro porque se va superponer con la licencia
                    StrSql = " DELETE his_estructura "
                    StrSql = StrSql & " WHERE tenro   = 30 "
                    StrSql = StrSql & " AND   estrnro  = " & rs_Est!estrnro
                    StrSql = StrSql & " AND   ternro  = " & Ternro
                    StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                    StrSql = StrSql & " AND   htethasta  is null "
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                'Inserto la misma situacion despues de la nueva situacion (la de la licencia)
                StrSql = "INSERT INTO his_estructura "
                StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde) "
                StrSql = StrSql & " VALUES (30, " & Ternro & ", "
                StrSql = StrSql & rs_Est!estrnro & ", "
                StrSql = StrSql & ConvFecha(CDate(FechaHasta + 1)) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                If rs_Est!htethasta > FechaHasta Then
                    If rs_Est!htetdesde > fechadesde Then
                        StrSql = " UPDATE his_estructura SET "
                        StrSql = StrSql & " htethasta = " & ConvFecha(CDate(fechadesde - 1))
                        StrSql = StrSql & " WHERE tenro   = 30 "
                        StrSql = StrSql & " AND   ternro  = " & Ternro
                        StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                        StrSql = StrSql & " AND   htethasta  = " & ConvFecha(rs_Est!htethasta)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'la borro porque se va superponer con la licencia
                        StrSql = " DELETE his_estructura "
                        StrSql = StrSql & " WHERE tenro   = 30 "
                        StrSql = StrSql & " AND   estrnro  = " & rs_Est!estrnro
                        StrSql = StrSql & " AND   ternro  = " & Ternro
                        StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                        StrSql = StrSql & " AND   htethasta  = " & ConvFecha(rs_Est!htethasta)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    'Inserto la misma situacion despues de la nueva situacion (la de la licencia)
                    StrSql = "INSERT INTO his_estructura "
                    StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
                    StrSql = StrSql & " VALUES (30, " & Ternro & ", "
                    StrSql = StrSql & rs_Est!estrnro & ", "
                    StrSql = StrSql & ConvFecha(CDate(FechaHasta + 1)) & ", "
                    StrSql = StrSql & ConvFecha(rs_Est!htethasta) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    If rs_Est!htetdesde > fechadesde Then
                        StrSql = " UPDATE his_estructura SET "
                        StrSql = StrSql & " htethasta = " & ConvFecha(CDate(fechadesde - 1))
                        StrSql = StrSql & " WHERE tenro   = 30 "
                        StrSql = StrSql & " AND   ternro  = " & Ternro
                        StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                        StrSql = StrSql & " AND   htethasta  is null "
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'la borro porque se va superponer con la licencia
                        StrSql = " DELETE his_estructura "
                        StrSql = StrSql & " WHERE tenro   = 30 "
                        StrSql = StrSql & " AND   estrnro  = " & rs_Est!estrnro
                        StrSql = StrSql & " AND   ternro  = " & Ternro
                        StrSql = StrSql & " AND   htetdesde = " & ConvFecha(rs_Est!htetdesde)
                        StrSql = StrSql & " AND   htethasta  = " & ConvFecha(rs_Est!htethasta)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
    
        StrSql = "INSERT INTO his_estructura "
        StrSql = StrSql & " (tenro, ternro, estrnro, htetdesde,htethasta) "
        StrSql = StrSql & " VALUES (30, " & Ternro & ", "
        StrSql = StrSql & Estrnro_SitRev & ", "
        StrSql = StrSql & ConvFecha(fechadesde) & ", "
        StrSql = StrSql & ConvFecha(FechaHasta) & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        Flog.writeline "Las Licencias por vacaciones no tienen Situacion de Revista asociado"
    End If

End Sub


Public Sub Politica1500v_7_1(ByVal Tipo As String)
'***************************************************************
' En cuanto se toma al menos un día de vacaciones, se le paga todos los días correspondientes y descuenta los dias correspondientes
' a partir de la fecha de inicio de la licencia.
'***************************************************************
Dim rs As New Recordset
Dim rs2 As New Recordset
Dim StrSql As String

Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Restantes As Integer
Dim fechaAux As Date

Dim Jornal As Boolean
Dim NroTer As Long

Dim ya_se_pago As Boolean
Dim cantdias As Integer
Dim blv As New ADODB.Recordset
Dim hay_licencia As Boolean

On Error GoTo CE

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If



    'Busco la licencia dentro del intervalo especificado
    StrSql = "SELECT Empleado,elfechadesde,elfechahasta,vacnro FROM emp_lic " & _
            " INNER JOIN lic_vacacion ON emp_lic.emp_licnro= lic_vacacion.emp_licnro" & _
            " WHERE lic_vacacion.emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    
    ' Si no hay licencia me voy
    If rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Licencia inexistente." & nrolicencia
        GoTo CE
    Else
        NroTer = rs!Empleado
        Mes_Inicio = Month(rs!elfechadesde)
        Ano_Inicio = Year(rs!elfechadesde)
        Mes_Fin = Month(rs!elfechahasta)
        Ano_Fin = Year(rs!elfechahasta)
        NroVac = rs!vacnro
    End If
    
    
    Genera_Pagos = False
    Genera_Descuentos = False
    Call Politica(1507)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
        Genera_Pagos = True
        Genera_Descuentos = True
    End If
    
    'VERIFICAR el reproceso y manejar la depuracion
    If Not Reproceso Then
        'si no es reproceso y existe el desglose de pago/descuento, salir
        StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
            GoTo CE
        End If
    Else
        'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
        StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
            GoTo CE
        Else
            'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
            StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    
    'EAM- Verifica si existe el tipo de licencia
    If Not ExisteTipoLicencia(2) Then
        Flog.writeline Espacios(Tabulador * 1) & "Tipo de dia de Vacaciones (2) inexistente."
        GoTo CE
    End If
    
    
    ya_se_pago = False
    'Busco todas las licencias del empleado en el periodo
    StrSql = "SELECT * FROM lic_vacacion" & _
            " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro" & _
            " WHERE emp_lic.empleado = " & NroTer & " AND vacnro= " & NroVac
    OpenRecordset StrSql, blv
    
    Do While Not blv.EOF
        StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro= " & blv!emp_licnro '& " Or (Ternro = " & NroTer & " And vacnro = " & NroVac & " and pago_dto=3)"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            ya_se_pago = True
            Exit Do
        End If
        blv.MoveNext
    Loop
    
    'EAM- Si cantidad es uno quiere decir que todavía no genero el pago. Se verifica porque cicla tantas veces como licencias encontradas
    'If CantidadLicenciasProcesadas = 1 Then
    If Not (ya_se_pago) Then
        'EAM- Obtiene la cantidad de días correspondientes
        StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & NroVac & " AND ternro = " & NroTer & " AND (venc is null OR venc = 0)"
        OpenRecordset StrSql, rs
        cantdias = rs!vdiascorcant
       
        Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, cantdias, 3) 'Pago
    
        StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
        OpenRecordset StrSql, rs
        
        Fin_Licencia = False
        Mes_Afecta = Mes_Inicio
        Ano_Afecta = Ano_Inicio
        'Dias_Restantes = cantdias 'comento este dias restantes seba 19/09/2013
        fechaAux = rs!elfechadesde
        
        'pongo este dias restantes seba 19/09/2013
        Dias_Restantes = rs!elcantdias
        
    
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
        DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
    
        If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
            Dias_Afecta = Dias_Restantes
            fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
        Else
            'Continua en el mes siguiente
            If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
                Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
                fechaAux = Ultimo_Mes
            Else
                Dias_Afecta = 30
                fechaAux = Ultimo_Mes
            End If
        End If
    
        Do While Not Fin_Licencia
            Dias_Restantes = Dias_Restantes - Dias_Afecta
            'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
            'Dias_Afecta = rs!elcantdias
            
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            
            '/* determinar a que continua en el proximo mes */
            If (Mes_Afecta = 12) Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                Mes_Afecta = Mes_Afecta + 1
            End If
    
            If (Dias_Restantes <= 0) Then
                Fin_Licencia = True
            End If
    
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
            DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
    
    
            If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
                Dias_Afecta = Dias_Restantes
                fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
            Else
                'Continua en el mes siguiente
                If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
                    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
                    fechaAux = Ultimo_Mes
                Else
                    Dias_Afecta = 30
                    fechaAux = Ultimo_Mes
                End If
                
            End If
    
    
        Loop
    
    Else
        
        'genero el dto
        StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & NroVac & " AND ternro = " & NroTer & " AND (venc is null OR venc = 0)"
        OpenRecordset StrSql, rs
        cantdias = rs!vdiascorcant
       
        'Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, cantdias, 3) 'Pago
    
        StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
        OpenRecordset StrSql, rs
        
        Fin_Licencia = False
        Mes_Afecta = Mes_Inicio
        Ano_Afecta = Ano_Inicio
        'Dias_Restantes = cantdias
        fechaAux = rs!elfechadesde
        
        StrSql = " SELECT sum(cantdias) valor FROM vacpagdesc "
        StrSql = StrSql & " WHERE vacnro=" & NroVac & " AND ternro=" & NroTer & " AND pago_dto=4 "
        OpenRecordset StrSql, rs2
        If Not rs.EOF Then
            Dias_Restantes = rs2!valor
        End If
        
        Dias_Restantes = cantdias - Dias_Restantes
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
        DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
    
        If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
            Dias_Afecta = Dias_Restantes
            fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
        Else
            'Continua en el mes siguiente
            If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
                Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
                fechaAux = Ultimo_Mes
            Else
                Dias_Afecta = 30
                fechaAux = Ultimo_Mes
            End If
        End If
    
        
        Do While Not Fin_Licencia
            Dias_Restantes = Dias_Restantes - Dias_Afecta
            Dias_Afecta = rs!elcantdias
            'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
            
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            
            '/* determinar a que continua en el proximo mes */
            If (Mes_Afecta = 12) Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                Mes_Afecta = Mes_Afecta + 1
            End If
    
            If (Dias_Restantes <= 0) Then
                Fin_Licencia = True
            End If
    
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
            DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
    
    
            If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
                Dias_Afecta = Dias_Restantes
                fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
            Else
                'Continua en el mes siguiente
                If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
                    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
                    fechaAux = Ultimo_Mes
                Else
                    Dias_Afecta = 30
                    fechaAux = Ultimo_Mes
                End If
                
            End If
    
    
        Loop
        'hasta aca
        
    End If



GoTo ProcesadoOK

CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing
    Set blv = Nothing

End Sub
Public Sub Politica1500v_TMK()
'-----------------------------------------------------------------------
' Customizacion para Temaiken
' paga adelantado todo (en el periodo de liq anterior al que corresponde la licencia.
' descuenta por mes lo que corresponde generar todos los dias de pago
'-----------------------------------------------------------------------
Dim rs As New ADODB.Recordset
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

Dim Aux_mes_inicio As Integer
Dim Aux_ano_inicio As Integer

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
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
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If
rs.Close

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If
'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE FROM vacpagdesc"
        StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND (pago_dto = 3" 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
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
Fin_Licencia = False
Mes_Afecta = Mes_Inicio
Ano_Afecta = Ano_Inicio
Dias_Pendientes = 0
Dias_Restantes = Dias_Afecta

'/* determinar los dias que afecta para el primer mes de decuento */
'/* Genera 30 dias, para todos los meses */
Anio_bisiesto = EsBisiesto(Ano_Afecta)

If Mes_Afecta = 2 Then
' Date en progress es una función
' la sintaxis es DATE(month,day,year)

    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
Else
    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    'ultimo_mes = AFecha(mes_afecta, 30, ano_afecta)
    If Mes_Afecta = 12 Then
        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
    Else
        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    End If
End If

If (rs!elfechahasta <= Ultimo_Mes) Then
'/* termina en el mes */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
Else
'/* continua en el mes siguiente */
    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
End If

Do While Not Fin_Licencia
    Dias_Restantes = Dias_Restantes - Dias_Afecta
    'Revisar
    'RUN generar_descuento(mes_afecta, ano_afecta, dias_afecta, 4)
    'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
    Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
    'IF RETURN-VALUE <> ""
    '         THEN UNDO main, RETURN.
    
    '/* determinar a que continua en el proximo mes */
    If (Mes_Afecta = 12) Then
        Mes_Afecta = 1
        Ano_Afecta = Ano_Afecta + 1
    Else
        Mes_Afecta = Mes_Afecta + 1
    End If
    If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
        Fin_Licencia = True
    End If
    
    '/* determinar los d¡as que afecta */

    Anio_bisiesto = EsBisiesto(Ano_Afecta)

    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    Else
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        'ultimo_mes = AFecha(mes_afecta, 30, ano_afecta)
        If Mes_Afecta = 12 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
        Else
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    End If

    If (rs!elfechahasta <= Ultimo_Mes) Then
    '/* termina en el mes */
        If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
            Dias_Afecta = Dias_Restantes
        Else
            Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
        End If
    Else
    '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
    End If
Loop

rs.Close
Set rs = Nothing

End Sub


'Public Sub Politica1500PagayDescuenta_Mes_a_Mes()
'' corresponde a vacpdo02
''***************************************************************
''PAGA Y DESCUENTA POR MES SIN TOPE DE 30 DIAS FIJOS.
''***************************************************************
'Dim rs As New Recordset
'Dim StrSql As String
'
'Dim Mes_Inicio As Integer
'Dim Ano_Inicio As Integer
'Dim Mes_Fin    As Integer
'Dim Ano_Fin    As Integer
'Dim Fin_Licencia  As Boolean
'Dim Mes_Afecta    As Integer
'Dim Ano_Afecta    As Integer
'Dim Primero_Mes   As Date
'Dim Ultimo_Mes    As Date
'Dim Dias_Afecta   As Integer
'Dim Dias_Pendientes As Integer
'Dim Dias_Restantes As Integer
'
'Dim Dias_ya_tomados As Integer
'Dim Fecha_limite    As Date
'Dim Anio_bisiesto   As Boolean
'Dim Jornal As Boolean
'Dim Legajo As Long
'Dim NroTer As Long
'Dim Nombre As String
'
'Call Politica(1503)
'If Not PoliticaOK Then
'    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
'    TipDiaPago = 3
'    TipDiaDescuento = 3
'    'Exit Sub
'End If
'
'
''Busco la licencia dentro del intervalo especificado
'StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
'rs.Open StrSql, objConn
'' Si no hay licencia me voy
'If rs.EOF Then
'    'MsgBox "Licencia inexistente.", vbCritical
'    rs.Close
'    Set rs = Nothing
'    Exit Sub
'Else
'    'NroTer = emp_lic.empleado
'    NroTer = rs!Empleado
'    Mes_Inicio = Month(rs!elfechadesde)
'    Ano_Inicio = Year(rs!elfechadesde)
'    Mes_Fin = Month(rs!elfechahasta)
'    Ano_Fin = Year(rs!elfechahasta)
'End If
'rs.Close
'
'Genera_Pagos = False
'Genera_Descuentos = False
'Call Politica(1507)
'If Not PoliticaOK Then
'    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
'    Genera_Pagos = True
'    Genera_Descuentos = True
'End If
''VERIFICAR el reproceso y manejar la depuracion
'If Not Reproceso Then
'    'si no es reproceso y existe el desglose de pago/descuento, salir
'    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
'    rs.Open StrSql, objConn
'    If Not rs.EOF Then
'        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
'        rs.Close
'        Set rs = Nothing
'        Exit Sub
'    End If
'    rs.Close
'Else
'    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
'    StrSql = "SELECT * FROM vacpagdesc "
'    StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
'    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
'    rs.Open StrSql, objConn
'    If Not rs.EOF Then
'        rs.Close
'        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
'    Else
'        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
'        rs.Close
'        StrSql = "DELETE FROM vacpagdesc"
'        StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
'        If Genera_Pagos And Genera_Descuentos Then
'            StrSql = StrSql & " AND (pago_dto = 3" 'pagos
'            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
'        Else
'            If Genera_Pagos Then
'                StrSql = StrSql & " AND pago_dto = 3" 'pagos
'            End If
'            If Genera_Descuentos Then
'                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
'            End If
'        End If
'        objConn.Execute StrSql, , adExecuteNoRecords
'    End If
'End If
'
'
'StrSql = "SELECT * FROM tipdia WHERE tdnro = " & TipoLicencia
'rs.Open StrSql, objConn
'If rs.EOF Then
'    'MsgBox " Tipo de dia de Vacaciones (2) inexistente."
'    rs.Close
'    Set rs = Nothing
'    Exit Sub
'Else
'    rs.Close
'End If
'
'StrSql = "SELECT * FROM empleado WHERE ternro = " & NroTer
'rs.Open StrSql, objConn
''jornal = IIf(rs!folinro = 2, True, False)
'rs.Close
'
''/* POLITICA  - ANALISIS */
'
'Fin_Licencia = False
'Mes_Afecta = Mes_Inicio
'Ano_Afecta = Ano_Inicio
'Dias_Pendientes = 0
'
''/* determinar los dias que afecta para el primer mes de descuento */
'Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
'DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
'
'StrSql = "SELECT * FROM emp_lic "
'If TipoLicencia = 2 Then
'    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
'End If
'StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
'OpenRecordset StrSql, rs
'If TipoLicencia = 2 Then
'    NroVac = rs!vacnro
'End If
'
''StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
''rs.Open StrSql, objConn
'
'If (rs!elfechahasta <= Ultimo_Mes) Then
''/* termina en el mes */
'    Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
'Else
''/* continua en el mes siguiente */
'    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
'End If
'
'Do While Not Fin_Licencia
'    'RUN generar_pago  (mes_afecta,ano_afecta,dias_afecta,3)
'    'Call generar_pago(mes_inicio, ano_inicio, Dias_Afecta, 3, Jornal, nrolicencia)
'    'Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
'    'FGZ - 15/12/2004
'    Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaPago, Dias_Afecta, 3) 'Pago
'    'RUN generar_descuento  (mes_afecta,ano_afecta,dias_afecta,4)
'    'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
'    'Call Generar_PagoDescuento(mes_inicio, ano_inicio, TipDiaDescuento, Dias_Afecta, 4) 'Pago
'    'FGZ - 15/12/2004
'    Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Pago
'
'    If (Mes_Afecta = 12) Then
'        Mes_Afecta = 1
'        Ano_Afecta = Ano_Afecta + 1
'    Else
'        Mes_Afecta = Mes_Afecta + 1
'    End If
'
'    If (Ano_Afecta = Ano_Fin) Then
'        If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
'            Fin_Licencia = True
'        End If
'    Else
'        If (Ano_Afecta > Ano_Fin) Then
'            Fin_Licencia = True
'        End If
'
'    End If
'
'    '/* determinar los d­as que afecta */
'    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'    Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
'    DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
'
'    If (rs!elfechahasta <= Ultimo_Mes) Then
'    '/* termina en el mes */
'        Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
'    Else
'    '/* continua en el mes siguiente */
'        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
'    End If
'
'Loop
'
'End Sub



Public Sub Politica1500_v2AdelantaDescuenta_Jornal()
'-----------------------------------------------------------------------
'?????.p    'Glencore
'Genera Pago/dto por dias correspondientes
' Jornales: paga adelantado todo.
' No Jornales: paga adelantado todo y descuenta por mes lo que corresponde generar todos los dias de pago
'Autor: FGZ
'Fecha: 27/07/2005
'Ultima Modificacion: FGZ - 14/10/2005
'-----------------------------------------------------------------------
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset

Dim Jornal As Boolean
Dim Dias_Afecta As Integer
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin As Integer
Dim Ano_Fin As Integer
Dim Aux_Fecha As Date
Dim Aux_TotalDiasCorrespondientes As Long

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No Existe el periodo de vacaciones " & NroVac
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro & _
         " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If

Mes_Inicio = Month(rs_vacacion!vacfecdesde)
Ano_Inicio = Year(rs_vacacion!vacfecdesde)
Mes_Fin = Month(rs_vacacion!vacfechasta)
Ano_Fin = Year(rs_vacacion!vacfechasta)

'Dias_Afecta = rs_vacdiascor!vdiascorcant
'FGZ - 14/10/2005
'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta

'FGZ - 27/07/2005
'Busco la forma de liquidacion
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
        Else
            Jornal = False
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
    End If
Else
    'Flog.writeline "No se encuentra estructura de forma de Liquidacion del empleado"
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
    'Exit Sub
End If

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac & " AND vacpagdesc.pronro is not null "
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos Then
            StrSql = StrSql & " AND pago_dto = 3" 'pagos
        End If
        If Genera_Descuentos Then
            StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Ojo que esta generando tipos 1 y 2 y deberia generar 3 y 4
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
If Not Jornal Then
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
End If

' Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
End Sub

Public Sub Politica1500_AdelantaDescuenta_Schering()
'-----------------------------------------------------------------------
'Customizacion para SCHERING.
' Anticipa y
' descuenta por mes lo que corresponde segun la licencia.
'Fecha: 10/08/2005
'Autor: FGZ
'-----------------------------------------------------------------------
Dim Mes_Inicio      As Integer
Dim Ano_Inicio      As Integer
Dim Mes_Fin         As Integer
Dim Ano_Fin         As Integer
Dim Fin_Licencia    As Boolean
Dim Mes_Afecta      As Integer
Dim Ano_Afecta      As Integer
Dim Primero_Mes     As Date
Dim Ultimo_Mes      As Date
Dim dias            As Integer
Dim Diasanti        As Integer
Dim Dias_Afecta     As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes  As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal          As Boolean
Dim Legajo          As Long
Dim NroTer          As Long
Dim Nombre          As String
Dim objFeriado      As New Feriado
Dim EsFeriado       As Boolean
Dim Dia_Analisis    As Date
Dim Dia_Hasta       As Date
Dim Aux_Fecha       As Date
Dim Topeanti        As Integer 'INITIAL 25. /* Dias de tope para anticipo vac. */
Dim Primero         As Boolean
Dim rs              As New ADODB.Recordset
Dim rs_Estructura   As New ADODB.Recordset

Topeanti = 25 'Dias de tope para anticipo vac.
Primero = True

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
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
    Aux_Fecha = rs!elfechahasta
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If
rs.Close

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If
'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE FROM vacpagdesc"
        StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND (pago_dto = 3" 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
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

'Busco la forma de liquidacion
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
        Else
            Jornal = False
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQl: " & StrSql
    Jornal = False
End If

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

Mes_Inicio = Month(rs!elfechadesde)
Ano_Inicio = Year(rs!elfechadesde)
Mes_Fin = Month(rs!elfechahasta)
Ano_Fin = Year(rs!elfechahasta)
dias = IIf(Month(rs!elfechadesde) = 12, 31, Day(CDate("01/" & Month(rs!elfechadesde) + 1 & "/" & Year(rs!elfechadesde)) - 1))
Dias_Restantes = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
'Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1

'-------------------------------------------------------------------------
'ANTICIPO
Set objFeriado.Conexion = objConn
Set objFeriado.ConexionTraza = objConn

If Month(rs!elfechadesde) <> Month(rs!elfechahasta) Then
    Dia_Analisis = rs!elfechadesde
    Dia_Hasta = CDate(dias & "/" & Month(rs!elfechadesde) & "/" & Year(rs!elfechadesde))
    Do While Dia_Analisis <= Dia_Hasta
        EsFeriado = objFeriado.Feriado(Dia_Analisis, Ternro, False)
        If Not EsFeriado And Weekday(Dia_Analisis) <> 7 And Weekday(Dia_Analisis) <> 1 Then
            Dias_Afecta = Dias_Afecta + 1
        End If
        Dia_Analisis = Dia_Analisis + 1
        Dias_Restantes = Dias_Restantes - 1
    Loop
    Diasanti = Dias_Afecta
    If Diasanti > Topeanti Then
        Diasanti = Topeanti
    End If
    Fin_Licencia = False
Else
    Dia_Analisis = rs!elfechadesde
    Dia_Hasta = rs!elfechahasta
    Do While Dia_Analisis <= Dia_Hasta
        EsFeriado = objFeriado.Feriado(Dia_Analisis, Ternro, False)
        If Not EsFeriado And Weekday(Dia_Analisis) <> 7 And Weekday(Dia_Analisis) <> 1 Then
            Dias_Afecta = Dias_Afecta + 1
        End If
        Dia_Analisis = Dia_Analisis + 1
        Dias_Restantes = Dias_Restantes - 1
    Loop
    Diasanti = Dias_Afecta
    If Diasanti > Topeanti Then
        Diasanti = Topeanti
    End If
End If
'RUN generar-pago (mes-inicio,ano-inicio,diasanti,1).
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Diasanti, 1) 'Pago
'FIN ANTICIPO
'-------------------------------------------------------------------------

'-------------------------------------------------------------------------
'PAGO Y DESCUENTO
'RUN generar-pago (mes-inicio,ano-inicio,dias-afecta,3).
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
'RUN generar-descuento (mes-inicio,ano-inicio,dias-afecta,4).
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 4) 'Descuento

Do While Not Fin_Licencia
    Primero = False
    Dias_Afecta = 0
    
    If Dias_Restantes > 0 Then
        Dia_Hasta = Dia_Analisis + Dias_Restantes - 1
        Do While Dia_Analisis <= Dia_Hasta
            EsFeriado = objFeriado.Feriado(Dia_Analisis, NroTer, False)
            If Not EsFeriado And Weekday(Dia_Analisis) <> 7 And Weekday(Dia_Analisis) <> 1 Then
                Dias_Afecta = Dias_Afecta + 1
            End If
            Dia_Analisis = Dia_Analisis + 1
            Dias_Restantes = Dias_Restantes - 1
        Loop
        'determinar a que continua en el proximo mes
        If (Mes_Inicio = 12) Then
            Mes_Inicio = 1
            Ano_Inicio = Ano_Inicio + 1
        Else
            Mes_Inicio = Mes_Inicio + 1
        End If
        'RUN generar-pago (mes-inicio,ano-inicio,dias-afecta,1).
        Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 1) 'Pago
        'RUN generar-pago (mes-inicio,ano-inicio,dias-afecta,3).
        Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
        'RUN generar-descuento (mes-inicio,ano-inicio,dias-afecta,4).
        Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 4) 'Pago
    Else
        Fin_Licencia = True
    End If
Loop

rs.Close
Set rs = Nothing
End Sub


Public Sub Politica1500_V2AdelantaDescuenta_Nueva()
'-----------------------------------------------------------------------
'Descripcion: paga adelantado todo y descuenta por mes lo que corresponde.
'               Genera a partir de dias correspondientes.
' Autor     : FGZ
'Ult Mod    : FGZ - 12-12-2005
'             Cuando el legajo es jornal realiza los cortes y topeos por quincena
'Ult Mod    : FGZ - 07/04/2006
            'Aux_TipDiaDescuento = TipDiaDescuento
            'Ojo que cuando cae el fin en un dia 31 hace macanas dado que topea a 30
            'especialmente en los jornales
            ' Gustavo Ring .- 16/07/2008 .- calculaba mal los descuentos
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Corrio_Dia As Boolean
Dim acarreo As Integer
Dim Aux_TotalDiasCorrespondientes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim rs As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones"
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If

Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If

'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
Aux_Generar_Fecha_Hasta = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal. SQL : " & StrSql
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If

Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaDescuento = st_ModeloDto2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = st_ModeloDto
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio


If Genera_Pagos Then
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
End If


Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

If Genera_Dto_DiasCorr Then



    'tantos descuentos como meses afecte
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Afecta
    
    'determinar los dias que afecta para el primer mes de decuento
    'Genera 30 dias, para todos los meses
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else

        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If
    
    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta)
    Else
        'continua en el mes siguiente
        'FGZ - 27/02/2006 agregué el +1
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
    End If
    
    Do While Dias_Restantes > 0
     
        If Jornal Then
            If Dias_Afecta > 15 Then
                Dias_Afecta = 15
            End If
        Else
            If Dias_Afecta > 30 Then
                Dias_Afecta = 30
            End If
        End If
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        If Genera_Descuentos And (Dias_Afecta <> 0) Then
            If Not Jornal Then
                   Aux_TipDiaDescuento = 7
            End If
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
            Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
            
            'FGZ - 07/04/2006
            If (Day(Aux_Generar_Fecha_Desde) = 31) And ((Dias_Afecta = 30) Or (Dias_Afecta = 15 And Jornal)) Then
                Aux_Generar_Fecha_Desde = Aux_Generar_Fecha_Desde + 1
                Corrio_Dia = True
            End If
        End If
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                Fin_Licencia = True
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        Fin_Licencia = True
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                Fin_Licencia = True
            End If
        End If
        
        'determinar los dias que afecta
        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
    
        If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
            End If
        Else
            'continua en el mes siguiente
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If
        
        'FGZ - 07/04/2006
        If Corrio_Dia Then
         '   Dias_Afecta = Dias_Afecta - 1
         '   Dias_Restantes = Dias_Restantes - 1
        End If
        
        
        If Not Jornal Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        Else
            If Dias_Afecta > 15 Then
               Dias_Afecta = 15
            End If
        End If
        
        'FGZ - Revisar esta inicializacion (creo que no debe ir)
        'Aux_TipDiaPago = TipDiaPago
        'Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                'Aux_TipDiaPago = TipDiaDescuento
                'FGZ - 07/04/2006
                Aux_TipDiaDescuento = st_ModeloDto2
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = st_ModeloDto
            End If
        End If
    Loop
Else
    Flog.writeline "No se generan los descuentos."
End If


'Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close
If rs_vacacion.State = adStateOpen Then rs_vacacion.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
Set rs_vacacion = Nothing
Set rs_Estructura = Nothing
End Sub

Public Sub Politica1500_v20_PorDiasCorresp()
'-----------------------------------------------------------------------
'Descripcion: paga adelantado todo y descuenta por mes lo que corresponde.
'               Genera a partir de dias correspondientes.
' Autor     : FGZ
'Ult Mod    : FGZ - 12-12-2005
'             Cuando el legajo es jornal realiza los cortes y topeos por quincena
'Ult Mod    : FGZ - 07/04/2006
            'Aux_TipDiaDescuento = TipDiaDescuento
            'Ojo que cuando cae el fin en un dia 31 hace macanas dado que topea a 30
            'especialmente en los jornales
            ' Gustavo Ring .- 16/07/2008 .- calculaba mal los descuentos
            ' FAF - Version nueva a partir de la 12 - Politica1500_V2AdelantaDescuenta_Nueva.
            ' Paga en el modelo definido en la politica 1503. La version anterior paga en el modelo 7
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Corrio_Dia As Boolean
Dim acarreo As Integer
Dim Aux_TotalDiasCorrespondientes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim rs As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones"
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If

Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If

'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
Aux_Generar_Fecha_Hasta = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal. SQL : " & StrSql
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If

Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaDescuento = st_ModeloDto2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = st_ModeloDto
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
End If
Flog.writeline "Modelo Pago " & Aux_TipDiaPago
Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
Quincena_Siguiente = Quincena_Inicio


If Genera_Pagos Then
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
End If


Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

If Genera_Dto_DiasCorr Then



    'tantos descuentos como meses afecte
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Afecta
    
    'determinar los dias que afecta para el primer mes de decuento
    'Genera 30 dias, para todos los meses
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else

        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If
    
    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta)
    Else
        'continua en el mes siguiente
        'FGZ - 27/02/2006 agregué el +1
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
    End If
    
    Do While Dias_Restantes > 0
     
        If Jornal Then
            If Dias_Afecta > 15 Then
                Dias_Afecta = 15
            End If
        Else
            If Dias_Afecta > 30 Then
                Dias_Afecta = 30
            End If
        End If
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        If Genera_Descuentos And (Dias_Afecta <> 0) Then
            'If Not Jornal Then
            '       Aux_TipDiaDescuento = 7
            'End If
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
            Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
            
            'FGZ - 07/04/2006
            If (Day(Aux_Generar_Fecha_Desde) = 31) And ((Dias_Afecta = 30) Or (Dias_Afecta = 15 And Jornal)) Then
                Aux_Generar_Fecha_Desde = Aux_Generar_Fecha_Desde + 1
                Corrio_Dia = True
            End If
        End If
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                Fin_Licencia = True
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        Fin_Licencia = True
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                Fin_Licencia = True
            End If
        End If
        
        'determinar los dias que afecta
        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
    
        If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
            End If
        Else
            'continua en el mes siguiente
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If
        
        'FGZ - 07/04/2006
        If Corrio_Dia Then
         '   Dias_Afecta = Dias_Afecta - 1
         '   Dias_Restantes = Dias_Restantes - 1
        End If
        
        
        If Not Jornal Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        Else
            If Dias_Afecta > 15 Then
               Dias_Afecta = 15
            End If
        End If
        
        'FGZ - Revisar esta inicializacion (creo que no debe ir)
        'Aux_TipDiaPago = TipDiaPago
        'Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                'Aux_TipDiaPago = TipDiaDescuento
                'FGZ - 07/04/2006
                Aux_TipDiaDescuento = st_ModeloDto2
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = st_ModeloDto
            End If
        End If
    Loop
Else
    Flog.writeline "No se generan los descuentos."
End If


'Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close
If rs_vacacion.State = adStateOpen Then rs_vacacion.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
Set rs_vacacion = Nothing
Set rs_Estructura = Nothing
End Sub
Public Sub Politica1500_v25_PorDiasCorresp()
'-----------------------------------------------------------------------
'Autor      : Gonzalez Nicolás
'Fecha      : 08/08/2013
'Descripcion: Paga y Descuenta hasta 30 días por mes
'Modificado :12/08/2013 - NG - Corrección para el mes de Febrero.
'            21/08/2013 - NG - Corrección en partición de días.
'            22/10/2013 - NG - Corrección al partir días con los meses de 31 días.
'            24/10/2013 - NG - Corrección al partir días a partir del día 30.
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Corrio_Dia As Boolean
Dim acarreo As Integer
Dim Aux_TotalDiasCorrespondientes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim Aux_Generar_Fecha_Desde2
Dim Dias_Afecta2 As Integer


Dim rs As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones"
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If

Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If

'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
Aux_Generar_Fecha_Hasta = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal. SQL : " & StrSql
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If

Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaDescuento = st_ModeloDto2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = st_ModeloDto
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
End If
Flog.writeline "Modelo Pago " & Aux_TipDiaPago
Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
Quincena_Siguiente = Quincena_Inicio






'--------------------------------------------------------------
If Genera_Pagos Then
    Aux_Generar_Fecha_Desde2 = Aux_Generar_Fecha_Desde
    Dias_Afecta2 = Dias_Afecta
    
    'tantos Pagos como meses afecte
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Afecta2
    
    'determinar los dias que afecta para el primer mes de Pago
    'Genera 30 dias, para todos los meses
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else

        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If
    
'    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
'        'termina en el mes
'        Dias_Afecta2 = DateDiff("d", Aux_Generar_Fecha_Desde2, Aux_Generar_Fecha_Hasta)
'    Else
'        'continua en el mes siguiente
'        Dias_Afecta2 = DateDiff("d", Aux_Generar_Fecha_Desde2, Ultimo_Mes) + 1
'    End If
   Dim esfebrero
   esfebrero = False
    Select Case Day(Aux_Generar_Fecha_Desde2)
        Case 28, 29:
                    If IsDate("31/" & Month(Aux_Generar_Fecha_Desde2) & "/" & Year(Aux_Generar_Fecha_Desde2)) Then
                        'Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde2)) + 2)
                        '22/10/2013
                        Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde2)) + 1)
                    Else
                        Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde2)) + 1)
                        esfebrero = True
                    End If
        Case 30:
                    If IsDate("31/" & Month(Aux_Generar_Fecha_Desde2) & "/" & Year(Aux_Generar_Fecha_Desde2)) Then
                        'Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde2)) + 2)
                        Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde2)) + 1)
                    Else
                        Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde2)) + 1)
                    End If
        Case 31:
                    Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde2)) + 2)
        Case Else
                    Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde2)) + 1)
    End Select
        
    If Dias_Afecta2 > Dias_Afecta Then
        Dias_Afecta2 = Dias_Afecta
    End If


    
    Do While Dias_Restantes > 0
     
        If Jornal Then
            If Dias_Afecta2 > 15 Then
                Dias_Afecta2 = 15
            End If
        Else
            If Dias_Afecta2 > 30 Then
                Dias_Afecta2 = 30
            End If
        End If
        
        'If (Month(Ultimo_Mes)) = 2 And (Dias_Restantes > Dias_Afecta2) And Dias_Afecta2 < 30 And esfebrero = False Then
         '   Dias_Restantes = Dias_Restantes - Dias_Afecta2 - (30 - Day(Ultimo_Mes))
         '   Dias_Afecta2 = Dias_Afecta2 + (30 - Day(Ultimo_Mes))
        'Else
            Dias_Restantes = Dias_Restantes - Dias_Afecta2
        'End If
        
        If Genera_Pagos And (Dias_Afecta2 <> 0) Then
            'Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta2, 3) 'Pago
            
            'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
            Aux_Generar_Fecha_Desde2 = DateAdd("d", Dias_Afecta2, Aux_Generar_Fecha_Desde2)
            
            If (Day(Aux_Generar_Fecha_Desde2) = 31) And ((Dias_Afecta2 = 30) Or (Dias_Afecta2 = 15 And Jornal)) Then
                Aux_Generar_Fecha_Desde2 = Aux_Generar_Fecha_Desde2 + 1
                Corrio_Dia = True
            End If
        End If
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                Fin_Licencia = True
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        Fin_Licencia = True
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                Fin_Licencia = True
            End If
        End If
        
        'determinar los dias que afecta
        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
    
        If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
            'If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
            If (Dias_Restantes <= DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta2 = Dias_Restantes
            Else
                Dias_Afecta2 = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
            End If
        Else
            'continua en el mes siguiente
            'Dias_Afecta2 = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
            If Dias_Restantes <= 30 Then
                Dias_Afecta2 = Dias_Restantes
            Else
                Dias_Afecta2 = 30
            End If
        End If
        
                
        If Not Jornal Then
            If Dias_Afecta2 > 30 Then
               Dias_Afecta2 = 30
            End If
        Else
            If Dias_Afecta2 > 15 Then
               Dias_Afecta2 = 15
            End If
        End If
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                 Aux_TipDiaDescuento = st_ModeloDto2
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = st_ModeloDto
            End If
        End If
    Loop
Else
    Flog.writeline "No se generan los Pagos."
End If
'------------------------------------------------------------


Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

If Genera_Dto_DiasCorr Then

   'tantos descuentos como meses afecte
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Afecta
    'determinar los dias que afecta para el primer mes de decuento
    'Genera 30 dias, para todos los meses
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else

        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If
    
    'If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
    '    Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta)
    'Else
        'continua en el mes siguiente
        'FGZ - 27/02/2006 agregué el +1
    '    Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
    'End If
    esfebrero = False
    Select Case Day(Aux_Generar_Fecha_Desde)
        Case 28, 29:
                    If IsDate("31/" & Month(Aux_Generar_Fecha_Desde) & "/" & Year(Aux_Generar_Fecha_Desde)) Then
                        'Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde)) + 2)
                        'NG 22/10/2013
                        Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde)) + 1)
                    Else
                        Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde)) + 1)
                        esfebrero = True
                    End If
                    
        Case 30:
                    If IsDate("31/" & Month(Aux_Generar_Fecha_Desde) & "/" & Year(Aux_Generar_Fecha_Desde)) Then
                        'Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde)) + 2)
                        Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde)) + 1)
                    Else
                        Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde)) + 1)
                    End If
        Case 31:
                    Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde)) + 2)
        Case Else
                    Dias_Afecta2 = ((30 - Day(Aux_Generar_Fecha_Desde)) + 1)
    End Select
        
    If Dias_Afecta2 < Dias_Afecta Then
        Dias_Afecta = Dias_Afecta2
    End If
    
    
    Do While Dias_Restantes > 0
     
        If Jornal Then
            If Dias_Afecta > 15 Then
                Dias_Afecta = 15
            End If
        Else
            If Dias_Afecta > 30 Then
                Dias_Afecta = 30
            End If
        End If
        
        
'       If (Month(Ultimo_Mes)) = 2 And (Dias_Restantes > Dias_Afecta) And Dias_Afecta2 < 30 And esfebrero = False Then
'            Dias_Restantes = Dias_Restantes - Dias_Afecta - (30 - Day(Ultimo_Mes))
'            Dias_Afecta = Dias_Afecta + (30 - Day(Ultimo_Mes))
'        Else
            Dias_Restantes = Dias_Restantes - Dias_Afecta
'        End If
        
       
        
        If Genera_Descuentos And (Dias_Afecta <> 0) Then
            'If Not Jornal Then
            '       Aux_TipDiaDescuento = 7
            'End If
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
            Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
            
            'FGZ - 07/04/2006
            If (Day(Aux_Generar_Fecha_Desde) = 31) And ((Dias_Afecta = 30) Or (Dias_Afecta = 15 And Jornal)) Then
                Aux_Generar_Fecha_Desde = Aux_Generar_Fecha_Desde + 1
                Corrio_Dia = True
            End If
        End If
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                Fin_Licencia = True
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        Fin_Licencia = True
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                Fin_Licencia = True
            End If
        End If
        
        'determinar los dias que afecta
        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
    
        If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
            If (Dias_Restantes <= DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
            End If
        Else
            'continua en el mes siguiente
            'Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
            If Dias_Restantes <= 30 Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = 30
            End If
        End If
        
        'FGZ - 07/04/2006
        If Corrio_Dia Then
         '   Dias_Afecta = Dias_Afecta - 1
         '   Dias_Restantes = Dias_Restantes - 1
        End If
        
        
        If Not Jornal Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        Else
            If Dias_Afecta > 15 Then
               Dias_Afecta = 15
            End If
        End If
        
        'FGZ - Revisar esta inicializacion (creo que no debe ir)
        'Aux_TipDiaPago = TipDiaPago
        'Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                'Aux_TipDiaPago = TipDiaDescuento
                'FGZ - 07/04/2006
                Aux_TipDiaDescuento = st_ModeloDto2
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = st_ModeloDto
            End If
        End If
    Loop
Else
    Flog.writeline "No se generan los descuentos."
End If


'Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close
If rs_vacacion.State = adStateOpen Then rs_vacacion.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
Set rs_vacacion = Nothing
Set rs_Estructura = Nothing
End Sub
Public Sub Politica1500_V2PagaDescuenta_PorMes()
'-----------------------------------------------------------------------
'Descripcion: paga y descuenta por mes lo que corresponde.
'               Genera a partir de dias correspondientes.
' Autor     : FGZ
'Ult Mod    : FGZ - 25-01-2006
'             Cuando el legajo es jornal realiza los cortes y topeos por quincena
'             FAF - 25-09-2006 - Se corrijio para CAS 2382
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer

Dim Aux_TotalDiasCorrespondientes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim rs As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones " & NroVac
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If


Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If

'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
Aux_Generar_Fecha_Hasta = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If

Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaPago = TipDiaDescuento
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = TipDiaPago
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio


'Esto va dentro del ciclo en esta version
'If Genera_Pagos Then
'    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
'End If


Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

If Genera_Dto_DiasCorr Then

    'tantos descuentos como meses afecte
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Afecta
    
    'determinar los dias que afecta para el primer mes de decuento
    'Genera 30 dias, para todos los meses
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If
    
    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta)
    Else
        'continua en el mes siguiente
'        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes)
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
    End If
    
    Do While Not Fin_Licencia And Dias_Restantes > 0
        
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                Aux_TipDiaPago = TipDiaDescuento
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = TipDiaPago
            End If
        End If
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        If Genera_Pagos Then
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
        End If
        
        If Genera_Descuentos And (Dias_Afecta <> 0) Then
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
            Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
        End If
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                Fin_Licencia = True
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        Fin_Licencia = True
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                Fin_Licencia = True
            End If
        End If
        
        'determinar los dias que afecta
        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
    
        If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
            End If
        Else
            'continua en el mes siguiente
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If
        
        If Not Jornal Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        Else
            If Dias_Afecta > 15 Then
               Dias_Afecta = 15
            End If
        End If
        
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                Aux_TipDiaPago = TipDiaDescuento
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = TipDiaPago
            End If
        End If
    Loop
Else
    Flog.writeline "No se generan los descuentos."
End If


'Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close
If rs_vacacion.State = adStateOpen Then rs_vacacion.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
Set rs_vacacion = Nothing
Set rs_Estructura = Nothing
End Sub

Public Sub Politica1500_V3AdelantaDescuenta_Nueva()
'-----------------------------------------------------------------------
'Descripcion: paga adelantado todo y descuenta por mes lo que corresponde.
'               Genera a partir de dias correspondientes.
' Autor     : FGZ
'Ult Mod    : FGZ - 12-12-2005
'             Cuando el legajo es jornal realiza los cortes y topeos por quincena
'Ult Mod    : FGZ - 07/04/2006
            'Aux_TipDiaDescuento = TipDiaDescuento
            'Ojo que cuando cae el fin en un dia 31 hace macanas dado que topea a 30
            'especialmente en los jornales
            ' Gustavo Ring .- 16/07/2008 .- calculaba mal los descuentos
            'Margiotta, Emanuel (13972) - Se cambio el campo con el que calucla. En vez de tomar vacdiascorcant toma vacdiascorcantcorr
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Corrio_Dia As Boolean
Dim acarreo As Integer
Dim Aux_TotalDiasCorrespondientes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim rs As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones"
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If

Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If

'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
If IsNull(rs_vacdiascor!vdiascorcantcorr) Then
    Aux_TotalDiasCorrespondientes = 0
    Flog.writeline "La cantidad de días correspondientes es 0. No se generarán los pagos y los descuentos."
    Exit Sub
Else
    If rs_vacdiascor!vdiascorcantcorr > Total_Dias_A_Generar Then
        Dias_Afecta = Total_Dias_A_Generar
    Else
        Dias_Afecta = rs_vacdiascor!vdiascorcantcorr
        
    End If

    Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcantcorr
End If


Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
Aux_Generar_Fecha_Hasta = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal. SQL : " & StrSql
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If

Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
' Reutilizo el parametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaDescuento = st_ModeloDto2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = st_ModeloDto
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio


If Genera_Pagos Then
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
End If


Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

If Genera_Dto_DiasCorr Then



    'tantos descuentos como meses afecte
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Afecta
    
    'determinar los dias que afecta para el primer mes de decuento
    'Genera 30 dias, para todos los meses
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else

        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If
    
    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta)
    Else
        'continua en el mes siguiente
        'FGZ - 27/02/2006 agregué el +1
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
    End If
    
    Do While Dias_Restantes > 0
     
        If Jornal Then
            If Dias_Afecta > 15 Then
                Dias_Afecta = 15
            End If
        Else
            If Dias_Afecta > 30 Then
                Dias_Afecta = 30
            End If
        End If
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        If Genera_Descuentos And (Dias_Afecta <> 0) Then
'            If Not Jornal Then
'                   Aux_TipDiaDescuento = 7
'            End If
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
            Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
            
            'FGZ - 07/04/2006
            If (Day(Aux_Generar_Fecha_Desde) = 31) And ((Dias_Afecta = 30) Or (Dias_Afecta = 15 And Jornal)) Then
                Aux_Generar_Fecha_Desde = Aux_Generar_Fecha_Desde + 1
                Corrio_Dia = True
            End If
        End If
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                Fin_Licencia = True
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        Fin_Licencia = True
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                Fin_Licencia = True
            End If
        End If
        
        'determinar los dias que afecta
        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
    
        If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
            End If
        Else
            'continua en el mes siguiente
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If
        
        'FGZ - 07/04/2006
        If Corrio_Dia Then
         '   Dias_Afecta = Dias_Afecta - 1
         '   Dias_Restantes = Dias_Restantes - 1
        End If
        
        
        If Not Jornal Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        Else
            If Dias_Afecta > 15 Then
               Dias_Afecta = 15
            End If
        End If
        
        'FGZ - Revisar esta inicializacion (creo que no debe ir)
        'Aux_TipDiaPago = TipDiaPago
        'Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                'Aux_TipDiaPago = TipDiaDescuento
                'FGZ - 07/04/2006
                Aux_TipDiaDescuento = st_ModeloDto2
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = st_ModeloDto
            End If
        End If
    Loop
Else
    Flog.writeline "No se generan los descuentos."
End If


'Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close
If rs_vacacion.State = adStateOpen Then rs_vacacion.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
Set rs_vacacion = Nothing
Set rs_Estructura = Nothing
End Sub


Public Sub Politica1500_v3AdelantaDescuenta()
'-----------------------------------------------------------------------
'Genera Pago/dto por dias correspondientes del tipo de vacaciones configurado en el parametro st_TipoDia2 de la politica 1501
'paga adelantado todo y descuenta por mes lo que corresponde generar todos los dias de pago
'Fecha:
'Autor: Margiotta, Emanuel
'-----------------------------------------------------------------------
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim Jornal As Boolean
Dim Dias_Afecta As Integer
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin As Integer
Dim Ano_Fin As Integer
Dim Aux_TotalDiasCorrespondientes As Long

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer


StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones " & NroVac
    Exit Sub
End If


StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro & " AND vacdiascor.vacnro = " & NroVac & _
        " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If


Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If



'Reviso que la cantidad de dias correspondientes del periodo no supere la cantidad de dias que quedan por tomar
If Not IsNull(rs_vacdiascor!vdiascorcantcorr) Then
    Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcantcorr
Else
    Aux_TotalDiasCorrespondientes = 0
End If
If Aux_TotalDiasCorrespondientes > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = Aux_TotalDiasCorrespondientes
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
    
'Busco la forma de liquidacion
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
        Else
            Jornal = False
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
    End If
Else
    'Flog.writeline "No se encuentra estructura de forma de Liquidacion del empleado"
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
    'Exit Sub
End If

PoliticaOK = False
Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc  WHERE vacnro = " & NroVac & _
            " AND ternro =" & Ternro & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


' Reutilizo el parametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaPago = TipDiaDescuento
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = TipDiaPago
    End If
End If


If Pliq_Nro <> 0 Then
    Mes_Inicio = Pliq_Mes
    Ano_Inicio = Pliq_Anio
End If
Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

If Genera_Dto_DiasCorr Then
    
    'ojo que el dto si es jornal ==> se parte
    
    If Pliq_Nro <> 0 Then
        Mes_Inicio = Pliq_Mes
        Ano_Inicio = Pliq_Anio
    End If
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento

Else
    Flog.writeline "No se generan los descuentos."
End If


' Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
End Sub
Public Sub Politica1500_V4AdelantaDescuenta_Nueva()
'-----------------------------------------------------------------------
'Descripcion: Si no existe licencia genera el pago por dias correspondientes y no genera dto.
'             Si exsite licencia genera el pago por dias correspondientes y descuenta por licencias cuando corresponde
'Autor      : Sebastian Stremel
'Ult Mod    : 23/09/2013
'
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Corrio_Dia As Boolean
Dim acarreo As Integer
Dim Aux_TotalDiasCorrespondientes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim rs As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

Dim ya_se_pago As Boolean
Dim blv As New ADODB.Recordset
Dim cantdias As Integer
Dim fechaAux
Dim rs2 As New ADODB.Recordset
Dim rs_lic As New ADODB.Recordset


Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones"
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If

Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If

'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
If IsNull(rs_vacdiascor!vdiascorcantcorr) Then
    Aux_TotalDiasCorrespondientes = 0
    Flog.writeline "La cantidad de días correspondientes es 0. No se generarán los pagos y los descuentos."
    Exit Sub
Else
    If rs_vacdiascor!vdiascorcantcorr > Total_Dias_A_Generar Then
        Dias_Afecta = Total_Dias_A_Generar
    Else
        Dias_Afecta = rs_vacdiascor!vdiascorcantcorr
        
    End If

    Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcantcorr
End If


Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
Aux_Generar_Fecha_Hasta = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            GoTo descuentos
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal. SQL : " & StrSql
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If

Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
' Reutilizo el parametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaDescuento = st_ModeloDto2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = st_ModeloDto
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio


If Genera_Pagos Then
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
End If


Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

descuentos:
Genera_Dto_DiasCorr = True
If Genera_Dto_DiasCorr Then
    
    'Busco la licencia dentro del intervalo especificado
    StrSql = " SELECT Empleado,elfechadesde,elfechahasta,vacnro, emp_lic.emp_licnro FROM emp_lic "
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro= lic_vacacion.emp_licnro "
    StrSql = StrSql & " WHERE emp_lic.empleado = " & Ternro & " And  lic_vacacion.vacnro = " & NroVac
    StrSql = StrSql & " AND " & _
    " (elfechadesde >= " & ConvFecha(fecha_desde) & ") AND " & _
    " (elfechadesde <= " & ConvFecha(fecha_hasta) & ") "
    '"  emp_lic.licestnro = 2 " 'Autorizada
    '" WHERE lic_vacacion.emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs_lic
    
    ' Si no hay licencia me voy
    If rs_lic.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Licencia inexistente." & nrolicencia
        GoTo CE
    Else
        Do While Not rs_lic.EOF
            NroTer = rs_lic!Empleado
            Mes_Inicio = Month(rs_lic!elfechadesde)
            Ano_Inicio = Year(rs_lic!elfechadesde)
            Mes_Fin = Month(rs_lic!elfechahasta)
            Ano_Fin = Year(rs_lic!elfechahasta)
            NroVac = rs_lic!vacnro
            nrolicencia = rs_lic!emp_licnro
    'End If
        
        
        Genera_Pagos = False
        Genera_Descuentos = False
        Call Politica(1507)
        If Not PoliticaOK Then
            Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
            Genera_Pagos = True
            Genera_Descuentos = True
        End If
        
        'VERIFICAR el reproceso y manejar la depuracion
        If Not Reproceso Then
            'si no es reproceso y existe el desglose de pago/descuento, salir
            StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
                GoTo CE
            End If
        Else
            'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
            StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
                GoTo CE
            Else
                'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
                StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
        
        
        'EAM- Verifica si existe el tipo de licencia
        If Not ExisteTipoLicencia(2) Then
            Flog.writeline Espacios(Tabulador * 1) & "Tipo de dia de Vacaciones (2) inexistente."
            GoTo CE
        End If
        
        
        ya_se_pago = False
        'Busco todas las licencias del empleado en el periodo
        StrSql = "SELECT * FROM lic_vacacion" & _
                " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro" & _
                " WHERE emp_lic.empleado = " & NroTer & " AND vacnro= " & NroVac
        OpenRecordset StrSql, blv
        
        Do While Not blv.EOF
            StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro= " & blv!emp_licnro
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                ya_se_pago = True
                Exit Do
            End If
            blv.MoveNext
        Loop
        
        'EAM- Si cantidad es uno quiere decir que todavía no genero el pago. Se verifica porque cicla tantas veces como licencias encontradas
        'If CantidadLicenciasProcesadas = 1 Then
        If Not (ya_se_pago) Then
            'EAM- Obtiene la cantidad de días correspondientes
            StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & NroVac & " AND ternro = " & NroTer & " AND (venc is null OR venc = 0)"
            OpenRecordset StrSql, rs
            cantdias = rs!vdiascorcant
           
            'Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, cantdias, 3) 'Pago
        
            StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
            OpenRecordset StrSql, rs
            
            Fin_Licencia = False
            Mes_Afecta = Mes_Inicio
            Ano_Afecta = Ano_Inicio
            'Dias_Restantes = cantdias 'comento este dias restantes seba 19/09/2013
            fechaAux = rs!elfechadesde
            
            'pongo este dias restantes seba 19/09/2013
            Dias_Restantes = rs!elcantdias
            
        
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
            DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
        
            If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
                Dias_Afecta = Dias_Restantes
                fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
            Else
                'Continua en el mes siguiente
                If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
                    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
                    fechaAux = Ultimo_Mes
                Else
                    Dias_Afecta = 30
                    fechaAux = Ultimo_Mes
                End If
            End If
        
            Do While Not Fin_Licencia
                Dias_Restantes = Dias_Restantes - Dias_Afecta
                'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
                'Dias_Afecta = rs!elcantdias
                
                Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
                
                '/* determinar a que continua en el proximo mes */
                If (Mes_Afecta = 12) Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    Mes_Afecta = Mes_Afecta + 1
                End If
        
                If (Dias_Restantes <= 0) Then
                    Fin_Licencia = True
                End If
        
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
                DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
        
        
                If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
                    Dias_Afecta = Dias_Restantes
                    fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
                Else
                    'Continua en el mes siguiente
                    If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
                        Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
                        fechaAux = Ultimo_Mes
                    Else
                        Dias_Afecta = 30
                        fechaAux = Ultimo_Mes
                    End If
                    
                End If
        
        
            Loop
        
        Else
            
            'genero el dto
            StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & NroVac & " AND ternro = " & NroTer & " AND (venc is null OR venc = 0)"
            OpenRecordset StrSql, rs
            cantdias = rs!vdiascorcant
           
            'Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, cantdias, 3) 'Pago
        
            StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
            OpenRecordset StrSql, rs
            
            Fin_Licencia = False
            Mes_Afecta = Mes_Inicio
            Ano_Afecta = Ano_Inicio
            'Dias_Restantes = cantdias
            fechaAux = rs!elfechadesde
            
            StrSql = " SELECT sum(cantdias) valor FROM vacpagdesc "
            StrSql = StrSql & " WHERE vacnro=" & NroVac & " AND ternro=" & NroTer & " AND pago_dto=4 "
            OpenRecordset StrSql, rs2
            If Not rs.EOF Then
                Dias_Restantes = rs2!valor
            End If
            
            Dias_Restantes = cantdias - Dias_Restantes
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
            DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
        
            If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
                Dias_Afecta = Dias_Restantes
                fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
            Else
                'Continua en el mes siguiente
                If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
                    Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
                    fechaAux = Ultimo_Mes
                Else
                    Dias_Afecta = 30
                    fechaAux = Ultimo_Mes
                End If
            End If
        
            
            Do While Not Fin_Licencia
                Dias_Restantes = Dias_Restantes - Dias_Afecta
                Dias_Afecta = rs!elcantdias
                'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
                
                Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Descuento
                
                '/* determinar a que continua en el proximo mes */
                If (Mes_Afecta = 12) Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    Mes_Afecta = Mes_Afecta + 1
                End If
        
                If (Dias_Restantes <= 0) Then
                    Fin_Licencia = True
                End If
        
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
                DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))
        
        
                If (DateAdd("d", Dias_Restantes, fechaAux) < Ultimo_Mes) Then
                    Dias_Afecta = Dias_Restantes
                    fechaAux = DateAdd("d", Dias_Restantes, fechaAux)
                Else
                    'Continua en el mes siguiente
                    If (DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1 <= 30) Then
                        Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
                        fechaAux = Ultimo_Mes
                    Else
                        Dias_Afecta = 30
                        fechaAux = Ultimo_Mes
                    End If
                    
                End If
        
        
            Loop
            'hasta aca
            
        End If
        'hasta aca
    rs_lic.MoveNext
    Loop
    End If
Else
    Flog.writeline "No se generan los descuentos."
End If

CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

'Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close
If rs_vacacion.State = adStateOpen Then rs_vacacion.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
Set rs_vacacion = Nothing
Set rs_Estructura = Nothing
End Sub



Public Sub Politica1500_V2PagaDescuenta_PorQuincena()
'-----------------------------------------------------------------------
'Descripcion: paga y descuenta por Quincena lo que corresponde.
'               Genera a partir de dias Licencias.
' Autor     : EAM
'Ult Mod    :
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Restantes As Integer
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim NroTer As Long
Dim Aux_Generar_Fecha_Hasta As Date
Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer
Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer
Dim rs As New ADODB.Recordset

On Error GoTo CE
Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If


'EAM- Busco los datos de la licencia que se esta analizando
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
'EAM- Si no hay licencia deja de analizar
If rs.EOF Then
    Flog.writeline "Licencia inexistente.", vbCritical
    GoTo CE
Else
    NroTer = rs!Empleado

    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Aux_Generar_Fecha_Desde = rs!elfechadesde
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
    Aux_Generar_Fecha_Hasta = rs!elfechahasta

    'EAM- (v3.28) verifico si esta vacio y lo seteo en 0. Guarda la lista de lic_nro para contabilizar las licencias pagas
    If listEmpLic = "" Then
        listEmpLic = 0
    End If
    listEmpLic = listEmpLic & "," & nrolicencia
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        Flog.writeline "DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA"
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'EAM- Busca la licencia y si es de vacación(2) obtengo el vacnro.
StrSql = "SELECT * FROM emp_lic "
If TipoLicencia = 2 Then
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
End If
StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If (TipoLicencia = 2) And Not (rs.EOF) Then
    NroVac = rs!vacnro
    
    If (NroVac <> NroVacAux) Then
        NroVacAux = NroVac
    End If
End If

'EAM- Busca la forma de liquidacion True-->Jornal | false -> Mensual
Jornal = FormaDeLiquidacion(fecha_desde, 22)

'Se setean los valores aca por si no es jornal
Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento

' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaPago = TipDiaPago2
        Aux_TipDiaDescuento = TipDiaDescuento2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio

'tantos descuentos como meses afecte
Fin_Licencia = False
Mes_Afecta = Mes_Inicio
Ano_Afecta = Ano_Inicio
    
'Determinar los dias que afecta
Anio_bisiesto = EsBisiesto(Ano_Afecta)

If Mes_Afecta = 2 Then
    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    If Not Jornal Then
        'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    Else
        If Quincena_Siguiente = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    End If
Else
    
    If Not Jornal Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Mes_Afecta = 12 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
        Else
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    Else
        If Quincena_Siguiente = 1 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    End If
End If
    
'calcula los dias de la licencia, dependiendo si llega a fin de la quincena o no
If Aux_Generar_Fecha_Hasta > Ultimo_Mes Then
    Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
Else
    Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta) + 1
End If


Dias_Restantes = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta) + 1


Do While Not Fin_Licencia And Dias_Restantes > 0
    Dias_Restantes = Dias_Restantes - Dias_Afecta
    If Genera_Pagos Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
    End If
    
    If Genera_Descuentos And (Dias_Afecta <> 0) Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
        'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
        Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
    End If
    'Continua en la siguiente quincena
    If Quincena_Siguiente = 1 Then
        Quincena_Siguiente = 2
        Aux_TipDiaPago = TipDiaPago2
        Aux_TipDiaDescuento = TipDiaDescuento2
    Else
        Quincena_Siguiente = 1
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
    End If
        
    'determinar a que continua en el proximo mes
    If (Mes_Afecta = 12) Then
        If Not Jornal Then
            Mes_Afecta = 1
            Ano_Afecta = Ano_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    Else
        If Not Jornal Then
            Mes_Afecta = Mes_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    End If
    
    If (Ano_Afecta = Ano_Fin) Then
        If (Mes_Afecta > Mes_Fin) Then
            Fin_Licencia = True
        Else
            If (Mes_Afecta = Mes_Fin) And Jornal Then
                If Quincena_Siguiente > Quincena_Fin Then
                    Fin_Licencia = True
                End If
            End If
        End If
    Else
        If (Ano_Afecta > Ano_Fin) Then
            Fin_Licencia = True
        End If
    End If
    
    'determinar los dias que afecta
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    If Mes_Afecta = 2 Then
        
        If Not Jornal Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If

    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
    'termina en el mes
        If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
            Dias_Afecta = Dias_Restantes
        Else
            Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
        End If
    Else
        'continua en el mes siguiente
        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
    End If
    
    If Not Jornal Then
        If Dias_Afecta > 30 Then
           Dias_Afecta = 30
        End If
    Else
        If Dias_Afecta > 15 Then
           Dias_Afecta = 15
        End If
    End If
    
Loop
GoTo ProcesadoOK
CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing

End Sub



Public Sub Politica1500_V2PagaDescuenta_PorQuincena_versionMDF()  'VERSION MDF
'-----------------------------------------------------------------------
'Descripcion: paga y descuenta por Quincena lo que corresponde.
'               Genera a partir de dias Licencias.
' Autor     : EAM
'Ult Mod    :
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_AfectaAux As Integer       'EAM- (v3.28)
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim cantDiasPagos As Integer    'EAM- (v3.28)



Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim NroTer As Long
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim rs As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset
Dim rs_Pagos  As New ADODB.Recordset

On Error GoTo CE
Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago2 = 3
    TipDiaDescuento2 = 3
End If


'EAM- Busco los datos de la licencia que se esta analizando
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
'EAM- Si no hay licencia deja de analizar
If rs.EOF Then
    Flog.writeline "Licencia inexistente.", vbCritical
    GoTo CE
Else
    NroTer = rs!Empleado
    
    
    'EAM- (v3.31)
    If Not IsNull(rs!elfechacert) Then
        Aux_Generar_Fecha_Desde = rs!elfechacert
        Mes_Inicio = Month(rs!elfechacert)
        Ano_Inicio = Year(rs!elfechacert)
        Mes_Fin = Month(CDate(DateAdd("d", rs!elcantdias, rs!elfechacert)))
        Ano_Fin = Year(CDate(DateAdd("d", rs!elcantdias, rs!elfechacert)))
        Aux_Generar_Fecha_Hasta = CDate(DateAdd("d", rs!elcantdias, rs!elfechacert))
    Else
        Mes_Inicio = Month(rs!elfechadesde)
        Ano_Inicio = Year(rs!elfechadesde)
        Aux_Generar_Fecha_Desde = rs!elfechadesde
        Mes_Fin = Month(rs!elfechahasta)
        Ano_Fin = Year(rs!elfechahasta)
        Aux_Generar_Fecha_Hasta = rs!elfechahasta
    End If
    
    
    'EAM- (v3.28) verifico si esta vacio y lo seteo en 0. Guarda la lista de lic_nro para contabilizar las licencias pagas
    If listEmpLic = "" Then
        listEmpLic = 0
    End If
    listEmpLic = listEmpLic & "," & nrolicencia
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
Else 'MDF
  Aux_TipDiaPago = TipDiaPago
  Aux_TipDiaDescuento = TipDiaDescuento
End If


'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        Flog.writeline "DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA"
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


'EAM- Busca la licencia y si es de vacación(2) obtengo el vacnro.
StrSql = "SELECT * FROM emp_lic "
If TipoLicencia = 2 Then
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
End If
StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If (TipoLicencia = 2) And Not (rs.EOF) Then
    NroVac = rs!vacnro
    
    If (NroVac <> NroVacAux) Then
        NroVacAux = NroVac
    End If
End If



'EAM- Busca la forma de liquidacion True-->Jornal | false -> Mensual
Jornal = FormaDeLiquidacion(fecha_desde, 22)

'--------------
'Aux_TipDiaPago = TipDiaPago2   MDF
'Aux_TipDiaDescuento = TipDiaDescuento2 MDF
'----------------
' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaPago = TipDiaPago2
        Aux_TipDiaDescuento = TipDiaDescuento2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaPago = TipDiaPago 'MDF
        Aux_TipDiaDescuento = TipDiaDescuento
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio



'tantos descuentos como meses afecte
Fin_Licencia = False
Mes_Afecta = Mes_Inicio
Ano_Afecta = Ano_Inicio
    
    
'Determinar los dias que afecta
Anio_bisiesto = EsBisiesto(Ano_Afecta)

If Mes_Afecta = 2 Then
    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    If Not Jornal Then
        'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    Else
        If Quincena_Siguiente = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    End If
Else
    
    If Not Jornal Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Mes_Afecta = 12 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
        Else
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    Else
        If Quincena_Siguiente = 1 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    End If
End If
    
Dim cant_aux 'MDF

'EAM- calcula los dias que afecta
If Not IsNull(rs!elcantdiashab) Then   'mdf 22/01/2015
    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        Dias_Afecta = rs!elcantdiashab
    Else
        Dias_Afecta = rs!elcantdiashab
    End If
Else
    Dias_Afecta = 0
    Flog.writeline "Dias_afecta=0, esto se da porque toma los dias del campo elcantdiashab...."
End If



cant_aux = DateDiff("d", CDate(rs!elfechadesde), CDate(rs!elfechahasta)) + 1 'mdf


'EAM- (v3.28) Acumulo en la variable la cantidad de días a pagar
'EAM- calcula los dias que restan


StrSql = "SELECT sum(cantdias) cantdias FROM vacpagdesc WHERE vacnro= " & NroVac & " AND (pago_dto= 3 or pago_dto= 1) AND emp_licnro in (" & listEmpLic & ")"
OpenRecordset StrSql, rs

If cant_aux <> Dias_Afecta Then 'MDF
    If Not IsNull(rs!cantdias) Then
        cantDiasPagos = rs!cantdias
        Dias_AfectaAux = Dias_Afecta + (cantDiasPagos Mod 5)
        Dias_Afecta = Dias_Afecta + (2 * Int(Dias_AfectaAux / 5))
    Else
       ' cantDiasPagos = 0 MDF
       ' Dias_AfectaAux = Dias_Afecta + (cantDiasPagos Mod 5) MDF
       ' Dias_Afecta = Dias_Afecta + (2 * Int(Dias_AfectaAux / 5)) MDF
        Dias_Afecta = cant_aux
    End If
End If 'MDF


'EAM ESTO ES PARA LA VERSION FINAL
'Dias_Restantes = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta) + 1

Dias_Restantes = Dias_Afecta  'DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta) + 1 '



Do While Not Fin_Licencia And Dias_Restantes > 0
    'Aux_TipDiaPago = TipDiaPago2
    'Aux_TipDiaDescuento = TipDiaDescuento2
    
'    If Jornal Then
'        Aux_TipDiaPago = Quincena_Siguiente
'        Aux_TipDiaDescuento = Quincena_Siguiente
'    End If

    'Dias_Restantes = Dias_Restantes - Dias_Afecta
    
'--------------MDF - Comentado  INICIOOOO
'    StrSql = "SELECT SUM(cantdias) cantdias FROM vacpagdesc " & _
'            " INNER JOIN periodo on periodo.pliqnro= vacpagdesc.pliqnro " & _
'            " WHERE Ternro = " & Ternro & " AND tprocnro = " & Aux_TipDiaPago & " And pliqmes = " & Mes_Afecta & " AND pliqanio= " & Ano_Afecta & " AND pago_dto = 3"
'    OpenRecordset StrSql, rs_Pagos
    
'    If (rs_Pagos!cantdias) < 15 Or IsNull(rs_Pagos!cantdias) Then
        'EAM- Ver como resuelvo esto. Hay que contralar la cantidad y de la quincena y
        
'      If IsNull(rs_Pagos!cantdias) Then
'              Dias_Afecta = Day(Ultimo_Mes) - Day(Aux_Generar_Fecha_Desde)     'MDF
'              Dias_Restantes = Dias_Restantes - Dias_Afecta
'      Else
'        If (rs_Pagos!cantdias + Dias_Afecta) <= 15 Or IsNull(rs_Pagos!cantdias) Then
'            Dias_Restantes = Dias_Restantes - Dias_Afecta
'            Dias_Afecta = Dias_Afecta
'        Else
'            Flog.writeline "____________________________________________________________________________________________________________"
'            Flog.writeline "La licencia " & nrolicencia & " del tercero " & Ternro & "  no se pagara ya que la Quincena esta COMPLETA."
'            Flog.writeline "____________________________________________________________________________________________________________"
'            Exit Sub
            'Dias_Restantes = Dias_Restantes - Dias_Afecta
            'Dias_Afecta = Dias_Afecta
'        End If
'      End If 'mdf
'    Else
'        Flog.writeline "____________________________________________________________________________________________________________"
'        Flog.writeline "La licencia " & nrolicencia & " del tercero " & Ternro & "  no se pagara ya que la Quincena esta COMPLETA."
'        Flog.writeline "____________________________________________________________________________________________________________"
'        Exit Sub
'    End If
 
 'NUEVO
 
    If Day(Aux_Generar_Fecha_Desde) = 15 Then
      'Aux_Generar_Fecha_Desde = DateAdd("d", 1, Aux_Generar_Fecha_Desde) MDF - 27/05/2015
    End If
    If Day(Aux_Generar_Fecha_Desde) = 31 Or Day(Aux_Generar_Fecha_Desde) = 30 Then
      'Aux_Generar_Fecha_Desde = DateAdd("d", 1, Aux_Generar_Fecha_Desde) MDF - 27/05/2015
    End If
     
    Dias_Afecta = (Day(Ultimo_Mes) - Day(Aux_Generar_Fecha_Desde)) + 1
    If Dias_Afecta > Dias_Restantes Then
       Dias_Afecta = Dias_Restantes
    End If
    Dias_Restantes = Dias_Restantes - Dias_Afecta
'fin NUEVO
'---------------------------MDF COMENTADO Fin

 

''EAM ESTO ES PARA LA VERSION FINAL
'    If (DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1) < Dias_Afecta Then
'        Dias_Restantes = Dias_Restantes - (DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1)
'        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
'    Else
'        Dias_Restantes = Dias_Restantes - Dias_Afecta
'        Dias_Afecta = Dias_Afecta
'    End If
    
    
    
    If Genera_Pagos Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
    End If
    
    If Genera_Descuentos And (Dias_Afecta <> 0) Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
        'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
        'Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde) MDF
    End If
   Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde) 'MDF
    'EAM- (3.30) - Se modifico para q tome segun la quincena el modelo de pago y descuento
    'Continua en la siguiente quincena
    If Quincena_Siguiente = 1 Then
        Quincena_Siguiente = 2
        Aux_TipDiaPago = TipDiaPago2
        Aux_TipDiaDescuento = TipDiaDescuento2
    Else
        Quincena_Siguiente = 1
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
    End If
    
    'determinar a que continua en el proximo mes
    If (Mes_Afecta = 12) Then
        If Not Jornal Then
            Mes_Afecta = 1
            Ano_Afecta = Ano_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    Else
        If Not Jornal Then
            Mes_Afecta = Mes_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    End If
    
    If (Ano_Afecta = Ano_Fin) Then
        If (Mes_Afecta > Mes_Fin) Then
            Fin_Licencia = True
        Else
            If (Mes_Afecta = Mes_Fin) And Jornal Then
                If Quincena_Siguiente > Quincena_Fin Then
                    Fin_Licencia = True
                End If
            End If
        End If
    Else
        If (Ano_Afecta > Ano_Fin) Then
            Fin_Licencia = True
        End If
    End If
    
    'determinar los dias que afecta
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If


'    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
'    'termina en el mes
'        If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
'            Dias_Afecta = Dias_Restantes
'        Else
'            Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
'        End If
'    Else
'        'continua en el mes siguiente
'        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
'    End If
    
    If Not Jornal Then
        If Dias_Afecta > 30 Then
          ' Dias_Afecta = 30 MDF
        End If
    Else
        If Dias_Afecta > 15 And Dias_Restantes > 0 Then
          ' Dias_Afecta = 15  MDF
        End If
    End If
    If Not Jornal Then 'MDF
      Aux_TipDiaPago = TipDiaPago 'MDF
      Aux_TipDiaDescuento = TipDiaDescuento 'MDF
    End If  'MDF
Loop
GoTo ProcesadoOK
CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing
    Set rs_Estructura = Nothing

End Sub
Public Sub Politica1500_V2PagaDescuenta_PorQuincena_sykes()
'-----------------------------------------------------------------------
'Descripcion: paga y descuenta por Quincena lo que corresponde.
'               Genera a partir de dias Licencias.
' Autor     : EAM
'Ult Mod    : 'mdf - Politica1500_V2PagaDescuenta_PorQuincena_sykes()
'antes se llamaba Politica1500_V2PagaDescuenta_PorQuincena pero esa quedo para la version 22 de la politica 1500 y se creo
'Politica1500_V2PagaDescuenta_PorQuincena_sykes() para ser usada por la version 27 de la 1500...
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_AfectaAux As Integer       'EAM- (v3.28)
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim cantDiasPagos As Integer    'EAM- (v3.28)



Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim NroTer As Long
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim rs As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset
Dim rs_Pagos  As New ADODB.Recordset

On Error GoTo CE
Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago2 = 3
    TipDiaDescuento2 = 3
End If


'EAM- Busco los datos de la licencia que se esta analizando
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
'EAM- Si no hay licencia deja de analizar
If rs.EOF Then
    Flog.writeline "Licencia inexistente.", vbCritical
    GoTo CE
Else
    NroTer = rs!Empleado
    
    
    'EAM- (v3.31)
    If Not IsNull(rs!elfechacert) Then
        Aux_Generar_Fecha_Desde = rs!elfechacert
        Mes_Inicio = Month(rs!elfechacert)
        Ano_Inicio = Year(rs!elfechacert)
        Mes_Fin = Month(CDate(DateAdd("d", rs!elcantdias, rs!elfechacert)))
        Ano_Fin = Year(CDate(DateAdd("d", rs!elcantdias, rs!elfechacert)))
        Aux_Generar_Fecha_Hasta = CDate(DateAdd("d", rs!elcantdias, rs!elfechacert))
    Else
        Mes_Inicio = Month(rs!elfechadesde)
        Ano_Inicio = Year(rs!elfechadesde)
        Aux_Generar_Fecha_Desde = rs!elfechadesde
        Mes_Fin = Month(rs!elfechahasta)
        Ano_Fin = Year(rs!elfechahasta)
        Aux_Generar_Fecha_Hasta = rs!elfechahasta
    End If
    
    
    'EAM- (v3.28) verifico si esta vacio y lo seteo en 0. Guarda la lista de lic_nro para contabilizar las licencias pagas
    If listEmpLic = "" Then
        listEmpLic = 0
    End If
    listEmpLic = listEmpLic & "," & nrolicencia
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If


'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        Flog.writeline "DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA"
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


'EAM- Busca la licencia y si es de vacación(2) obtengo el vacnro.
StrSql = "SELECT * FROM emp_lic "
If TipoLicencia = 2 Then
    StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
End If
StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs
If (TipoLicencia = 2) And Not (rs.EOF) Then
    NroVac = rs!vacnro
    
    If (NroVac <> NroVacAux) Then
        NroVacAux = NroVac
    End If
End If



'EAM- Busca la forma de liquidacion True-->Jornal | false -> Mensual
'Jornal = FormaDeLiquidacion(fecha_desde, 22) 'MDF
Jornal = FormaDeLiquidacion_sykes(fecha_desde, 22) 'MDF

'Aux_TipDiaPago = TipDiaPago2           MDF
'Aux_TipDiaDescuento = TipDiaDescuento2 MDF

' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaPago = TipDiaPago2
        Aux_TipDiaDescuento = TipDiaDescuento2
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio



'tantos descuentos como meses afecte
Fin_Licencia = False
Mes_Afecta = Mes_Inicio
Ano_Afecta = Ano_Inicio
    
    
'Determinar los dias que afecta
Anio_bisiesto = EsBisiesto(Ano_Afecta)

If Mes_Afecta = 2 Then
    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    If Not Jornal Then
        'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    Else
        If Quincena_Siguiente = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    End If
Else
    
    If Not Jornal Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Mes_Afecta = 12 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
        Else
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    Else
        If Quincena_Siguiente = 1 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    End If
End If
    
'EAM- calcula los dias que afecta
If Not IsNull(rs!elcantdiashab) Then   'mdf 22/01/2015
    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        Dias_Afecta = rs!elcantdiashab
    Else
        Dias_Afecta = rs!elcantdiashab
    End If
Else
    Dias_Afecta = 0
End If

'EAM- (v3.28) Acumulo en la variable la cantidad de días a pagar
'EAM- calcula los dias que restan
StrSql = "SELECT sum(cantdias) cantdias FROM vacpagdesc WHERE vacnro= " & NroVac & " AND (pago_dto= 3 or pago_dto= 1) AND emp_licnro in (" & listEmpLic & ")"
OpenRecordset StrSql, rs


If Not IsNull(rs!cantdias) Then
    cantDiasPagos = rs!cantdias
    Dias_AfectaAux = Dias_Afecta + (cantDiasPagos Mod 5)
    Dias_Afecta = Dias_Afecta + (2 * Int(Dias_AfectaAux / 5))
Else
    cantDiasPagos = 0
    Dias_AfectaAux = Dias_Afecta + (cantDiasPagos Mod 5)
    Dias_Afecta = Dias_Afecta + (2 * Int(Dias_AfectaAux / 5))
End If

'EAM ESTO ES PARA LA VERSION FINAL
'Dias_Restantes = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta) + 1

Dias_Restantes = Dias_Afecta 'DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta) + 1

Dim fecha_loop
fecha_loop = Aux_Generar_Fecha_Desde

Do While Not Fin_Licencia And Dias_Restantes > 0
    'Aux_TipDiaPago = TipDiaPago2
    'Aux_TipDiaDescuento = TipDiaDescuento2
    
'    If Jornal Then
'        Aux_TipDiaPago = Quincena_Siguiente
'        Aux_TipDiaDescuento = Quincena_Siguiente
'    End If

    'Dias_Restantes = Dias_Restantes - Dias_Afecta
    
    StrSql = "SELECT SUM(cantdias) cantdias FROM vacpagdesc " & _
            " INNER JOIN periodo on periodo.pliqnro= vacpagdesc.pliqnro " & _
            " WHERE Ternro = " & Ternro & " AND tprocnro = " & Aux_TipDiaPago & " And pliqmes = " & Mes_Afecta & " AND pliqanio= " & Ano_Afecta & " AND pago_dto = 3"
    OpenRecordset StrSql, rs_Pagos
     If Not Jornal Then
        If (rs_Pagos!cantdias) < 15 Or IsNull(rs_Pagos!cantdias) Then
            'EAM- Ver como resuelvo esto. Hay que contralar la cantidad y de la quincena y
            If (rs_Pagos!cantdias + Dias_Afecta) <= 15 Or IsNull(rs_Pagos!cantdias) Then
                Dias_Restantes = Dias_Restantes - Dias_Afecta
                Dias_Afecta = Dias_Afecta
            Else
                Flog.writeline "____________________________________________________________________________________________________________"
                Flog.writeline "La licencia " & nrolicencia & " del tercero " & Ternro & "  no se pagara ya que la Quincena esta COMPLETA."
                Flog.writeline "____________________________________________________________________________________________________________"
                Exit Sub
                'Dias_Restantes = Dias_Restantes - Dias_Afecta
                'Dias_Afecta = Dias_Afecta
            End If
        Else
            Flog.writeline "____________________________________________________________________________________________________________"
            Flog.writeline "La licencia " & nrolicencia & " del tercero " & Ternro & "  no se pagara ya que la Quincena esta COMPLETA."
            Flog.writeline "____________________________________________________________________________________________________________"
            Exit Sub
        End If
     Else
       
        If IsNull(rs_Pagos!cantdias) Then
           If Dias_Restantes >= (Day(Ultimo_Mes) - Day(fecha_loop)) + 1 Then
            Dias_Afecta = (Day(Ultimo_Mes) - Day(fecha_loop)) + 1
            Dias_Restantes = Dias_Restantes - Dias_Afecta
           Else
              Dias_Afecta = Dias_Restantes
              Dias_Restantes = 0
           End If
        End If
     
     End If
''EAM ESTO ES PARA LA VERSION FINAL
'    If (DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1) < Dias_Afecta Then
'        Dias_Restantes = Dias_Restantes - (DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1)
'        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
'    Else
'        Dias_Restantes = Dias_Restantes - Dias_Afecta
'        Dias_Afecta = Dias_Afecta
'    End If
    
    
    
    If Genera_Pagos Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
    End If
    
    If Genera_Descuentos And (Dias_Afecta <> 0) Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
        'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
        Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
    End If
    
    'EAM- (3.30) - Se modifico para q tome segun la quincena el modelo de pago y descuento
    'Continua en la siguiente quincena
    If Quincena_Siguiente = 1 Then
        Quincena_Siguiente = 2
        Aux_TipDiaPago = TipDiaPago2
        Aux_TipDiaDescuento = TipDiaDescuento2
    Else
        Quincena_Siguiente = 1
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
    End If
    
    'determinar a que continua en el proximo mes
    If (Mes_Afecta = 12) Then
        If Not Jornal Then
            Mes_Afecta = 1
            Ano_Afecta = Ano_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    Else
        If Not Jornal Then
            Mes_Afecta = Mes_Afecta + 1
        Else
            If Quincena_Siguiente = 1 Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                'Queda como esta, en el mismo mes y año
            End If
        End If
    End If
    
    If (Ano_Afecta = Ano_Fin) Then
        If (Mes_Afecta > Mes_Fin) Then
            Fin_Licencia = True
        Else
            If (Mes_Afecta = Mes_Fin) And Jornal Then
                If Quincena_Siguiente > Quincena_Fin Then
                    Fin_Licencia = True
                End If
            End If
        End If
    Else
        If (Ano_Afecta > Ano_Fin) Then
            Fin_Licencia = True
        End If
    End If
    
    'determinar los dias que afecta
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If


'    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
'    'termina en el mes
'        If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
'            Dias_Afecta = Dias_Restantes
'        Else
'            Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
'        End If
'    Else
'        'continua en el mes siguiente
'        Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
'    End If
    
    If Not Jornal Then
        If Dias_Afecta > 30 Then
           Dias_Afecta = 30
        End If
    Else
        If Dias_Afecta > 15 And Dias_Restantes > 0 Then
           Dias_Afecta = 15
        End If
    End If
    
   fecha_loop = Primero_Mes
Loop
GoTo ProcesadoOK
CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing
    Set rs_Estructura = Nothing

End Sub

Public Sub Politica1500v_24()
'***************************************************************
'  EN CUANTO SE TOMA AL MENOS UN DIA DE VACACIONES, SE LE PAGA Y DESCUENTA TODOS LOS DÍAS DE
'  VACACIONES QUE LE CORRESPONDEN PARA EL AÑO. PUEDE SER POR QUINCENA O MENSUAL
'***************************************************************
Dim rs As New Recordset
Dim StrSql As String

Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

Dim ya_se_pago As Boolean
Dim cantdias As Integer
Dim blv As New ADODB.Recordset
Dim bel As New ADODB.Recordset

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
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
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If
rs.Close

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If
'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE FROM vacpagdesc"
        StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND (pago_dto = 3" 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
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

If CantidadLicenciasProcesadas = 1 Then
    StrSql = "SELECT * FROM lic_vacacion WHERE emp_licnro= " & nrolicencia
    rs.Open StrSql, objConn
       
    StrSql = "SELECT * FROM vacacion WHERE vacnro = " & rs!vacnro
    rs.Close
    
    rs.Open StrSql, objConn
          
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & rs!vacnro & _
    " AND ternro = " & NroTer
    StrSql = StrSql & " AND (venc is null OR venc = 0)"
    rs.Close
    
    rs.Open StrSql, objConn
    cantdias = rs!vdiascorcant
    
    NroVac = rs!vacnro
    
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, cantdias, 3) 'PAgo
'    If Mes_Inicio = 1 Then
'        'Call generar_pago(12, ano_inicio - 1, CantDias, 3, Jornal, nrolicencia)
'        Call Generar_PagoDescuento(12, Ano_Inicio - 1, TipDiaPago, Dias_Afecta, 3) 'PAgo
'    Else
'        'Call generar_pago(mes_inicio - 1, ano_inicio, CantDias, 3, Jornal, nrolicencia)
'        Call Generar_PagoDescuento(Mes_Inicio - 1, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'PAgo
'    End If
    
    rs.Close
    
End If
    
    
    StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
    rs.Open StrSql, objConn
    
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = cantdias

    Anio_bisiesto = EsBisiesto(Ano_Afecta)

    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
    DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))

    If (rs!elfechahasta <= Ultimo_Mes) Then
        '/* termina en el mes */
        'Dias_Afecta = cantdias
        Dias_Afecta = rs!elcantdias
    Else
        '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
    End If

    Do While Not Fin_Licencia
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        'Call generar_descuento(mes_afecta, ano_afecta, Dias_Afecta, 4, Jornal, nrolicencia)
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'PAgo
        
        '/* determinar a que continua en el proximo mes */
        If (Mes_Afecta = 12) Then
            Mes_Afecta = 1
            Ano_Afecta = Ano_Afecta + 1
        Else
            Mes_Afecta = Mes_Afecta + 1
        End If

        If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
            Fin_Licencia = True
        End If

        '/* determinar los d¡as que afecta */

        Anio_bisiesto = EsBisiesto(Ano_Afecta)

        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        Ultimo_Mes = IIf(Mes_Afecta = 12, DateAdd("d", -1, AFecha(1, 1, Ano_Afecta + 1)), _
        DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta)))

        If (rs!elfechahasta <= Ultimo_Mes) Then
        '/* termina en el mes */
            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
            End If
        Else
        '/* continua en el mes siguiente */
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If

    Loop

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub


Private Sub politica1505(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
    ' 0 estandar
    ' 1 Glencore
    Call SetearParametrosPolitica(Detalle, ok)
    Select Case subn
        Case 1:
            st_BaseAntiguedad = 1
        Case 2:
            st_BaseAntiguedad = 2
        Case 3:
            st_BaseAntiguedad = 3
        Case 4:
            st_BaseAntiguedad = 4
        Case 5:
            st_BaseAntiguedad = 5
        Case 6:
            st_BaseAntiguedad = 6
        Case 7:
            st_BaseAntiguedad = 7
        Case Else
            st_BaseAntiguedad = 0
    End Select
    PoliticaOK = True
End Sub

Private Sub politica1506(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Activa la generacion de descuentos por dias correspondientes.
'Fecha: 15/10/2005
'Autor: FGZ
'-----------------------------------------------------------------------
    PoliticaOK = True
    Genera_Dto_DiasCorr = True
End Sub


Private Sub politica1507(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Activa los items a generar.
'Fecha: 15/10/2005
'Autor: FGZ
'-----------------------------------------------------------------------
Dim Opcion As Long

    PoliticaOK = False
    Call SetearParametrosPolitica(Detalle, ok)
    If ok Then
        Opcion = st_Opcion
        Select Case Opcion
            Case 1:
                Genera_Pagos = True
                Genera_Descuentos = True
            Case 2:
                Genera_Pagos = True
                Genera_Descuentos = False
            Case 3:
                Genera_Pagos = False
                Genera_Descuentos = True
            Case 4:
                Genera_Pagos = False
                Genera_Descuentos = False
            Case Else
                Genera_Pagos = True
                Genera_Descuentos = True
        End Select
        PoliticaOK = True
    Else
        PoliticaOK = False
    End If
End Sub



Private Sub politica1508(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Setea el parametro Tipo de Dia por Maternidad y un factor de division para
'             para el calculo de antiguedad de uruguay
'Fecha: 29/06/2006
'Autor: Fapitalle N.
'-----------------------------------------------------------------------
    
    Call SetearParametrosPolitica(Detalle, ok)
    If ok Then
        Tipo_Dia_Maternidad = st_TipoDia1
        Factor = st_Factor
        PoliticaOK = True
    Else
        PoliticaOK = False
    End If
End Sub

Private Sub politica1509(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Verifica si se realiza la situacion de revista
'Fecha: 16/06/2009
'Autor: Lisandro Moro
'-----------------------------------------------------------------------
        
    GenerarSituacionRevista = True

End Sub

Private Sub politica1510(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Indica si se como se toman los dias habiles y no habiles
'               Si se saca del tipo de vacaciones o es fijo de Lunes a Viernes.
'               Esta politica nace como custom para Altaplastica
'Fecha: 24/06/2009
'Autor: FGZ
'-----------------------------------------------------------------------
        
    Diashabiles_LV = True
    PoliticaOK = True
End Sub


Private Sub politica1511(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Vacaciones Acordadas
'               Indica si se controla y compara la cantidad de dias correspondientes por escala
'               y la cantidad de dias acordados por el empleado.
'Fecha: 25/06/2009
'Autor: FGZ
'-----------------------------------------------------------------------
        
    DiasAcordados = True
    PoliticaOK = True
End Sub

Public Sub Politica1500v_14()
'Corresponde a vacpdo06.p
'***************************************************************
'  EN CUANTO SE TOMA AL MENOS UN DIA DE VACACIONES, SE LE PAGA Y DESCUENTA TODOS LOS DÍAS DE
'  VACACIONES QUE LE CORRESPONDEN PARA EL AñO
'  TODO ESTO LO HACE MIENTRAS NO PAGO NI DESCONTO NINGUN DIA DE DICHA LICENCIA
'***************************************************************
' FAF - 11/08/2006 - Es copia de Politica1500v_7() con la unica diferencia que topea a 30 dias
Dim rs As New Recordset
Dim StrSql As String

Dim Dia_Inicio As Integer
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Dia_Fin    As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim i As Date
Dim Dias_Afecta   As Integer
'Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Dias_Hab_Afecta As Integer
Dim objFeriado      As New Feriado
Dim EsFeriado       As Boolean
Dim qna_afecta As Integer
Dim estrnro_GrupoLiq As Integer

Dim Anio_bisiesto   As Boolean
Dim Jornal As Integer
Dim NroTer As Long


On Error GoTo CE

Set objFeriado.Conexion = objConn

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs

' Si no hay licencia me voy
If rs.EOF Then
    Flog.wirteline Espacios(Tabulador * 1) & "Licencia inexistente." & nrolicencia
    GoTo CE
Else
    NroTer = rs!Empleado
    Dia_Inicio = Day(rs!elfechadesde)
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Dia_Fin = Day(rs!elfechahasta)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


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
Dias_Afecta = rs!elcantdias

If Genera_Pagos Then
    If Pliq_Nro <> 0 Then
        Mes_Inicio = Pliq_Mes
        Ano_Inicio = Pliq_Anio
    End If
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
End If


    StrSql = "SELECT estructura.estrnro, estructura.estrcodext FROM estructura " & _
            " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro " & _
            " INNER JOIN emp_lic ON his_estructura.ternro = emp_lic.empleado " & _
            " WHERE his_estructura.tenro = 32 AND his_estructura.htethasta IS NULL " & _
            " AND emp_lic.emp_licnro= " & nrolicencia
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Jornal = 1  ' 7 dias
        estrnro_GrupoLiq = 0
        Flog.writeline "    El empleado no posee Grupo de Liquidación. Se considera la semana corrida (7 días)."
    Else
        estrnro_GrupoLiq = rs!estrnro
        Select Case rs!estrcodext
            Case "1", "2":  ' 7 dias
                Jornal = 1
                Flog.writeline "    Según el Cód. Ext. (1 ó 2) del Grupo de Liquidación del empleado, se considera la semana corrida (7 días)."
            Case "3", "5":  ' 6 dias
                Jornal = 3
                Flog.writeline "    Según el Cód. Ext. (3 ó 5) del Grupo de Liquidación del empleado, se excluirán los domingos y feriados de la semana (6 días)."
            Case "4":       ' 5 dias
                Jornal = 4
                Flog.writeline "    Según el Cód. Ext. (4) del Grupo de Liquidación del empleado, se excluirán los domingos, sabados y feriados de la semana (5 días)."
            Case "6":       ' 7 dias
                Jornal = 2
                Flog.writeline "    Según el Cód. Ext. (6) del Grupo de Liquidación del empleado, se considera la semana corrida (7 días)."
            Case Else       ' 7 dias
                Jornal = 1
                Flog.writeline "    No está definido el Cód. Ext. del Grupo de Liquidación del empleado, se considera la semana corrida (7 días)."
        End Select
    End If
    rs.Close
    
    
    StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
    rs.Open StrSql, objConn
    Dias_Hab_Afecta = 0
    If Jornal > 2 Then
        For i = rs!elfechadesde To rs!elfechahasta
            EsFeriado = objFeriado.Feriado(i, rs!Empleado, False)
            Select Case Jornal
                Case 3:
                    If Not (EsFeriado Or Weekday(i) = 1) Then
                        Dias_Hab_Afecta = Dias_Hab_Afecta + 1
                    End If
                Case 4:
                    If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                        Dias_Hab_Afecta = Dias_Hab_Afecta + 1
                    End If
            End Select
        Next
    Else
        Dias_Hab_Afecta = rs!elcantdias
    End If
    
    Flog.writeline "    Días de la licencia: " & Dias_Afecta
    Flog.writeline "    Días que se descontaran en la licencia: " & Dias_Hab_Afecta
    
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    'Dias_Pendientes = 0
    Dias_Restantes = Dias_Hab_Afecta
    
    If Jornal > 2 Then
        If Dia_Inicio >= 16 Then
            qna_afecta = 2
        Else
            qna_afecta = 1
        End If
    Else
        If Jornal = 1 Then
            qna_afecta = 3
        Else
            qna_afecta = 5
        End If
    End If
    Flog.writeline "    Modelo de Liquidación (1.- Primera Quincena, 2.- Segunda Quincena, 3.- Mensuales, 5.- Liq. Final): " & qna_afecta
    
    '/* determinar los días que afecta para el primer mes de decuento */
    '/* Genera 30 dias mínimos, para todos los meses */
    Anio_bisiesto = EsBisiesto(Ano_Afecta)

    If qna_afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
    Else
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    End If
        
    If Mes_Afecta = 2 Then
        If qna_afecta = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            'FGZ - 27/01/2015 -----------------
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            'If Anio_bisiesto Then
            '    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
            'Else
            '    Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
            'End If
            'FGZ - 27/01/2015 -----------------
        End If
    Else
        If qna_afecta = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
            End If
        End If
    End If

    If (rs!elfechahasta <= Ultimo_Mes) Then
        '/* termina en el mes */
        Dias_Afecta = rs!elcantdias
        If Jornal > 2 Then
            Dias_Afecta = Dias_Hab_Afecta
        End If
    Else
        '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
        If Jornal > 2 Then
            Dias_Afecta = 0
            For i = rs!elfechadesde To Ultimo_Mes
                EsFeriado = objFeriado.Feriado(i, rs!Empleado, False)
                Select Case Jornal
                    Case 3:
                        If Not (EsFeriado Or Weekday(i) = 1) Then
                            Dias_Afecta = Dias_Afecta + 1
                        End If
                    Case 4:
                        If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                            Dias_Afecta = Dias_Afecta + 1
                        End If
                End Select
            Next
        End If
    End If
    
    Do While Not Fin_Licencia
        'FGZ - 22/01/2015 ---------------------
        If Dias_Afecta > Dias_Restantes Then
            Dias_Afecta = Dias_Restantes
        End If
        'FGZ - 22/01/2015 ---------------------
        Flog.writeline "    Días entre el " & Primero_Mes & " y el " & Ultimo_Mes & ": " & Dias_Afecta
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, qna_afecta, Dias_Afecta, 4) 'Dto
         
        Flog.writeline "    Días restantes: " & Dias_Restantes
        
        '/* determinar a que continua en el proximo mes */
        If qna_afecta = 1 Then
            qna_afecta = 2
        Else
            If qna_afecta = 2 Then
                qna_afecta = 1
            End If
            If (Mes_Afecta = 12) Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                Mes_Afecta = Mes_Afecta + 1
            End If
        End If
        
'        If (Mes_Afecta > Mes_Fin And Ano_Afecta >= Ano_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
'        If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
        If (Dias_Restantes <= 0) Then
            Fin_Licencia = True
        End If

        '/* determinar los días que afecta */
        Anio_bisiesto = EsBisiesto(Ano_Afecta)

        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        
        If qna_afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
        Else
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        End If
            
        If Mes_Afecta = 2 Then
            If qna_afecta = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                If Anio_bisiesto Then
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
                End If
            End If
        Else
            If qna_afecta = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                End If
            End If
        End If
        
        If (rs!elfechahasta <= Ultimo_Mes) Then
            '/* termina antes del ultimo dia del mes */
            Dias_Afecta = 0
            For i = Primero_Mes To rs!elfechahasta
                If Jornal > 2 Then
                    EsFeriado = objFeriado.Feriado(i, rs!Empleado, False)
                    Select Case Jornal
                        Case 3:
                            If Not (EsFeriado Or Weekday(i) = 1) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                        Case 4:
                            If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                    End Select
                Else
                    Dias_Afecta = Dias_Afecta + 1
                End If
            Next
        Else
            '/* continua en el mes siguiente */
            Dias_Afecta = 0
            For i = Primero_Mes To Ultimo_Mes
                If Jornal > 2 Then
                    EsFeriado = objFeriado.Feriado(i, rs!Empleado, False)
                    Select Case Jornal
                        Case 3:
                            If Not (EsFeriado Or Weekday(i) = 1) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                        Case 4:
                            If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                    End Select
                Else
                    Dias_Afecta = Dias_Afecta + 1
                End If
            Next
        End If
    Loop
GoTo ProcesadoOK

CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing

End Sub

Public Sub Politica1500_v14_DiasCorresp()
'-----------------------------------------------------------------------
'Custom Gorina - Standart
'Genera Pago/dto por dias correspondientes
'paga adelantado todo y descuenta por mes s/Grupo de Liquidacion al que pertenece.
'Particiona los dias y los asigna al modelo de Liquidacion acorde
'Fecha: 14/02/2007
'Autor: FAF
'-----------------------------------------------------------------------
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim Jornal As Integer
Dim Dias_Afecta As Integer
Dim Dia_Inicio As Integer
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin As Integer
Dim Ano_Fin As Integer
Dim Aux_TotalDiasCorrespondientes As Long
Dim fecDesde As Date
Dim fecHasta As Date
Dim estrnro_GrupoLiq As Integer
Dim Dias_Hab_Afecta As Integer
Dim i As Date
Dim objFeriado      As New Feriado
Dim EsFeriado       As Boolean
Dim qna_afecta As Integer
Dim Fin_Licencia As Boolean
Dim Mes_Afecta As Integer
Dim Ano_Afecta As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Anio_bisiesto As Boolean
Dim Primero_Mes As Date
Dim Ultimo_Mes As Date

'Dim Dias_ya_tomados As Integer
'Dim Fecha_limite    As Date
'Dim Anio_bisiesto   As Boolean

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Set objFeriado.Conexion = objConn

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "    No existe el período de vacaciones " & NroVac
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "    No hay días correspondientes a ese período. SQL: " & StrSql
    Exit Sub
End If

'Mes_Inicio = Month(rs_Vacacion!vacfecdesde)
'Ano_Inicio = Year(rs_Vacacion!vacfecdesde)
'Mes_Fin = Month(rs_Vacacion!vacfechasta)
'Ano_Fin = Year(rs_Vacacion!vacfechasta)

Dia_Inicio = Day(Aux_Generar_Fecha_Desde)
Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If


'FGZ - 14/10/2005
'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
    
PoliticaOK = False
Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "  Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "  Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'-----------------------------------------------------------------------------------------
'VERIFICAR el reproceso y manejar la depuracion
'-----------------------------------------------------------------------------------------
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "  No se seteo el Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "  Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "  No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


'-----------------------------------------------------------------------------------------
'Comienzo de generacion del Pago y del Descuento
'-----------------------------------------------------------------------------------------
    StrSql = "SELECT estructura.estrnro, estructura.estrcodext FROM estructura "
    StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.tenro = 32 AND his_estructura.htethasta IS NULL "
    StrSql = StrSql & " AND his_estructura.ternro= " & Ternro
    rs.Open StrSql, objConn
    If rs.EOF Then
        Jornal = 1  ' 7 dias
        estrnro_GrupoLiq = 0
        Flog.writeline "    El empleado no posee Grupo de Liquidación. Se considera la semana corrida (7 días)."
    Else
        estrnro_GrupoLiq = rs!estrnro
        Select Case rs!estrcodext
            Case "1", "2":  ' 7 dias
                Jornal = 1
                Flog.writeline "    Según el Cód. Ext. (1 ó 2) del Grupo de Liquidación del empleado, se considera la semana corrida (7 días)."
            Case "3", "5":  ' 6 dias
                Jornal = 3
                Flog.writeline "    Según el Cód. Ext. (3 ó 5) del Grupo de Liquidación del empleado, se excluirán los domingos y feriados de la semana (6 días)."
            Case "4":       ' 5 dias
                Jornal = 4
                Flog.writeline "    Según el Cód. Ext. (4) del Grupo de Liquidación del empleado, se excluirán los domingos, sabados y feriados de la semana (5 días)."
            Case "6":       ' 7 dias
                Jornal = 2
                Flog.writeline "    Según el Cód. Ext. (6) del Grupo de Liquidación del empleado, se considera la semana corrida (7 días)."
            Case Else       ' 7 dias
                Jornal = 1
                Flog.writeline "    No está definido el Cód. Ext. del Grupo de Liquidación del empleado, se considera la semana corrida (7 días)."
        End Select
    End If
    rs.Close
    
    If Genera_Descuentos Then
        If Pliq_Nro <> 0 Then
            Mes_Inicio = Pliq_Mes
            Ano_Inicio = Pliq_Anio
        End If
        Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
    End If
    
    fecDesde = Aux_Generar_Fecha_Desde
    fecHasta = DateAdd("d", Dias_Afecta, fecDesde)
    
    'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
    Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

    Genera_Dto_DiasCorr = False
    Call Politica(1506)
    If Not PoliticaOK Then
        Flog.writeline "  Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
    End If
    
    If Genera_Dto_DiasCorr Then
        Dias_Hab_Afecta = 0
        If Jornal > 2 Then
            For i = fecDesde To fecHasta
                EsFeriado = objFeriado.Feriado(i, Ternro, False)
                Select Case Jornal
                    Case 3:
                        If Not (EsFeriado Or Weekday(i) = 1) Then
                            Dias_Hab_Afecta = Dias_Hab_Afecta + 1
                        End If
                    Case 4:
                        If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                            Dias_Hab_Afecta = Dias_Hab_Afecta + 1
                        End If
                End Select
            Next
        Else
            Dias_Hab_Afecta = Dias_Afecta
        End If
        
        Flog.writeline "    Días de la licencia: " & Dias_Afecta
        Flog.writeline "    Días que se descontaran en la licencia: " & Dias_Hab_Afecta
        
        Fin_Licencia = False
        Mes_Afecta = Mes_Inicio
        Ano_Afecta = Ano_Inicio
        Dias_Pendientes = 0
        Dias_Restantes = Dias_Hab_Afecta
        
        If Jornal > 2 Then
            If Dia_Inicio >= 16 Then
                qna_afecta = 2
            Else
                qna_afecta = 1
            End If
        Else
            If Jornal = 1 Then
                qna_afecta = 3
            Else
                qna_afecta = 5
            End If
        End If
        Flog.writeline "    Modelo de Liquidación (1.- Primera Quincena, 2.- Segunda Quincena, 3.- Mensuales, 5.- Liq. Final): " & qna_afecta
        
        '/* determinar los días que afecta para el primer mes de decuento */
        '/* Genera 30 dias mínimos, para todos los meses */
        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        
        If qna_afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
        Else
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        End If
            
        If Mes_Afecta = 2 Then
            If qna_afecta = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                If Anio_bisiesto Then
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
                End If
            End If
        Else
            If qna_afecta = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                End If
            End If
        End If
        
        If (fecHasta <= Ultimo_Mes) Then
            '/* termina en el mes */
            If Jornal > 2 Then
                Dias_Afecta = Dias_Hab_Afecta
            End If
        Else
            '/* continua en el mes siguiente */
            Dias_Afecta = DateDiff("d", fecDesde, Ultimo_Mes) + 1
            If Jornal > 2 Then
                Dias_Afecta = 0
                For i = fecDesde To Ultimo_Mes
                    EsFeriado = objFeriado.Feriado(i, Ternro, False)
                    Select Case Jornal
                        Case 3:
                            If Not (EsFeriado Or Weekday(i) = 1) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                        Case 4:
                            If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                    End Select
                Next
            End If
        End If
        
        
        Do While Not Fin_Licencia
            
            Flog.writeline "    Días entre el " & Primero_Mes & " y el " & Ultimo_Mes & ": " & Dias_Afecta
    
            Dias_Restantes = Dias_Restantes - Dias_Afecta
            
    '        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'PAgo
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, qna_afecta, Dias_Afecta, 4) 'PAgo
             
            Flog.writeline "    Días restantes: " & Dias_Restantes
            
            '/* determinar a que continua en el proximo mes */
            If qna_afecta = 1 Then
                qna_afecta = 2
            Else
                If qna_afecta = 2 Then
                    qna_afecta = 1
                End If
                If (Mes_Afecta = 12) Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    Mes_Afecta = Mes_Afecta + 1
                End If
            End If
            
            If (Dias_Restantes <= 0) Then
                Fin_Licencia = True
            End If
    
            '/* determinar los días que afecta */
            Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
            If qna_afecta = 2 Then
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            End If
                
            If Mes_Afecta = 2 Then
                If qna_afecta = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    If Anio_bisiesto Then
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
                    End If
                End If
            Else
                If qna_afecta = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                    End If
                End If
            End If
            
            If (fecHasta <= Ultimo_Mes) Then
                '/* termina antes del ultimo dia del mes */
                Dias_Afecta = 0
                For i = Primero_Mes To fecHasta
                    If Jornal > 2 Then
                        EsFeriado = objFeriado.Feriado(i, Ternro, False)
                        Select Case Jornal
                            Case 3:
                                If Not (EsFeriado Or Weekday(i) = 1) Then
                                    Dias_Afecta = Dias_Afecta + 1
                                End If
                            Case 4:
                                If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                                    Dias_Afecta = Dias_Afecta + 1
                                End If
                        End Select
                    Else
                        Dias_Afecta = Dias_Afecta + 1
                    End If
                Next
            Else
                '/* continua en el mes siguiente */
                Dias_Afecta = 0
                For i = Primero_Mes To Ultimo_Mes
                    If Jornal > 2 Then
                        EsFeriado = objFeriado.Feriado(i, Ternro, False)
                        Select Case Jornal
                            Case 3:
                                If Not (EsFeriado Or Weekday(i) = 1) Then
                                    Dias_Afecta = Dias_Afecta + 1
                                End If
                            Case 4:
                                If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                                    Dias_Afecta = Dias_Afecta + 1
                                End If
                        End Select
                    Else
                        Dias_Afecta = Dias_Afecta + 1
                    End If
                Next
            End If
        Loop
    Else
        Flog.writeline "  No se generan los descuentos (VER Politica 1506)."
    End If
    
    
    
' Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
End Sub

Public Sub Politica1500_PagoDescuento_AGD()

    '------------------------------------------------------------------------------------------------------------------------------------------
    'Descripcion: paga todo y descuenta por mes con tope a 30 dias comenzando a partir del mes marcado como LiquidaVac
    ' Autor     :
    'Ult Mod    : Moro Lisandro - 20-02-2007 -
    '------------------------------------------------------------------------------------------------------------------------------------------
    Dim Mes_Inicio As Integer
    Dim Ano_Inicio As Integer
    Dim Mes_Fin    As Integer
    Dim Ano_Fin    As Integer
    Dim Fin_Licencia  As Boolean
    Dim Mes_Afecta    As Integer
    Dim Ano_Afecta    As Integer
    Dim Primero_Mes   As Date
    Dim Ultimo_Mes    As Date
    Dim Dias_Afecta   As Integer
    Dim Dias_Pendientes As Integer
    Dim Dias_Restantes As Integer
    Dim DiasPrevios As Integer
    Dim elfechadesde As Date
    Dim elfechahasta As Date
    Dim l_vdiapeddesde As Date

    Dim Dias_ya_tomados As Integer
    Dim Fecha_limite    As Date
    Dim Anio_bisiesto   As Boolean
    Dim Jornal As Boolean
    Dim Legajo As Long
    Dim NroTer As Long
    Dim Nombre As String
    Dim Aux_Fecha As Date
    'Dim GenerarSituacionRevista As Boolean
    
    Dim Aux_TipDiaPago As Integer
    Dim Aux_TipDiaDescuento As Integer

    Dim Quincena_Inicio As Integer
    Dim Quincena_Fin As Integer
    Dim Quincena_Siguiente As Integer

    Dim rs As New ADODB.Recordset
    Dim rs_Estructura  As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset

    Call Politica(1503)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
        TipDiaPago = 3
        TipDiaDescuento = 3
        'Exit Sub
    End If

    GenerarSituacionRevista = False
    Call Politica(1509)
    
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
        Mes_Inicio = Month(rs!elfechadesde)
        Ano_Inicio = Year(rs!elfechadesde)

        Flog.writeline "Licencia desde " & rs!elfechadesde & " al " & rs!elfechahasta
    End If
    rs.Close

    Genera_Pagos = False
    Genera_Descuentos = False
    Call Politica(1507)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
        Genera_Pagos = True
        Genera_Descuentos = True
    End If
    
    
    'VERIFICAR el reproceso y manejar la depuracion
    If Not Reproceso Then
        'si no es reproceso y existe el desglose de pago/descuento, salir
        StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        rs.Open StrSql, objConn
        If Not rs.EOF Then
            Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        rs.Close
    Else
        'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
        StrSql = "SELECT * FROM vacpagdesc "
        StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
        StrSql = StrSql & " AND vacpagdesc.pronro is not null"
        rs.Open StrSql, objConn
        If Not rs.EOF Then
            rs.Close
            Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        Else
            'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
            Flog.writeline "DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA"
            rs.Close
            StrSql = "DELETE FROM vacpagdesc"
            StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
            If Genera_Pagos And Genera_Descuentos Then
                StrSql = StrSql & " AND (pago_dto = 3" 'pagos
                StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
            Else
                If Genera_Pagos Then
                    StrSql = StrSql & " AND pago_dto = 3" 'pagos
                End If
                If Genera_Descuentos Then
                    StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
                End If
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If

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


    'POLITICA  - ANALISIS
    StrSql = "SELECT * FROM emp_lic "
    If TipoLicencia = 2 Then
        StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
        StrSql = StrSql & " INNER JOIN vacdiascor ON lic_vacacion.vacnro = vacdiascor.vacnro AND ternro = " & Ternro
    End If
    StrSql = StrSql & " WHERE emp_lic.emp_licnro= " & nrolicencia
    StrSql = StrSql & " AND (venc is null OR venc = 0)"
    OpenRecordset StrSql, rs
    If TipoLicencia = 2 Then
        NroVac = rs!vacnro
    End If
    Dias_Afecta = rs!vdiascorcant
    rs.Close
    Flog.writeline "Dias afecta " & Dias_Afecta

    'Inicializo deacuerdo al periodo de liquidacion asociado a la fecha del pedido de vacaciones
    '    y el campo liquidavac
    Dim l_mes As Integer
    Dim l_anio As Integer
    Dim l_dia As Integer
    Dim Fecha_Fin As Date
    
    StrSql = " SELECT * FROM vacdiasped "
    StrSql = StrSql & " WHERE LiquidaVac = -1 "
    StrSql = StrSql & " AND vdiaspedestado = -1 "
    StrSql = StrSql & " AND ternro = " & Ternro
    StrSql = StrSql & " AND vacnro = " & NroVac
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        l_mes = CInt(Month(rs2!vdiapeddesde))
        l_anio = CInt(Year(rs2!vdiapeddesde))
        l_dia = CInt(Day(rs2!vdiapeddesde))
        l_vdiapeddesde = CDate(rs2!vdiapeddesde)
    Else
        Flog.writeline "No se encontro el Pedido de Vacaciones Marcado como Liquida Vacaciones o no se encuentra aprobado"
        rs2.Close
        Set rs2 = Nothing
        Exit Sub
    End If
    rs2.Close

    Aux_TipDiaPago = TipDiaPago
    Aux_TipDiaDescuento = TipDiaDescuento

    '-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-x-'
    'Busco la cantidad de dias ya descontados
    DiasPrevios = 0
    StrSql = " SELECT cantdias FROM vacpagdesc "
    StrSql = StrSql & " WHERE pago_dto = 4 "
    StrSql = StrSql & " AND ternro = " & Ternro
    StrSql = StrSql & " AND vacnro = " & NroVac
    OpenRecordset StrSql, rs2
    If Not rs2.EOF Then
        Do While Not rs2.EOF
            DiasPrevios = DiasPrevios + CInt(rs2!cantdias)
            rs2.MoveNext
        Loop
    Else
        DiasPrevios = 0
    End If
    rs2.Close

    Flog.writeline "Dias ya descontados " & DiasPrevios
    Dias_Afecta = Dias_Afecta - DiasPrevios
    Flog.writeline "Dias reales que afecta " & Dias_Afecta

    If Genera_Pagos And Dias_Afecta <> 0 Then
        'Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
        Call Generar_PagoDescuento(l_mes, l_anio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
    End If
    
    'Se configura por la politica 1509 si se genera la situacion de revista junto al pago descuento
    If GenerarSituacionRevista And Dias_Afecta <> 0 Then
        Call InsertarSituacionRevista(Ternro, l_vdiapeddesde, (DateAdd("d", Dias_Afecta, l_vdiapeddesde) - 1))
    End If
    
    If Dias_Afecta <> 0 Then
        'tantos descuentos como meses afecte
        Fin_Licencia = False
        'Mes_Afecta = Mes_Inicio
        'Ano_Afecta = Ano_Inicio
        Mes_Afecta = l_mes
        Ano_Afecta = l_anio
        Dias_Pendientes = 0
        Dias_Restantes = Dias_Afecta
        
        'determinar los dias que afecta para el primer mes de decuento
        'Genera 30 dias, para todos los meses
        
        
        Primero_Mes = AFecha(l_mes, l_dia, l_anio)
        Fecha_Fin = DateAdd("d", Dias_Afecta - 1, Primero_Mes)
        Mes_Fin = Month(Fecha_Fin)
        Ano_Fin = Year(Fecha_Fin)
        
        If Mes_Afecta = 2 Then
            Anio_bisiesto = EsBisiesto(Ano_Afecta)
            If Anio_bisiesto Then
                Ultimo_Mes = AFecha(Mes_Afecta, 29, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta, 28, Ano_Afecta)
            End If
        Else
            Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
        End If
    
        If (Fecha_Fin <= Ultimo_Mes) Then
            '/* termina en el mes */
            Dias_Afecta = Dias_Afecta
        Else
            '/* continua en el mes siguiente */
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If
        
        Do While Not (Fin_Licencia) Or Dias_Restantes > 0
            
            Dias_Restantes = Dias_Restantes - Dias_Afecta
            
            If Dias_Afecta > 0 Then
                Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            End If
            
            '/* determinar a que continua en el proximo mes */
            If (Mes_Afecta = 12) Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                Mes_Afecta = Mes_Afecta + 1
            End If
    
            If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
                Fin_Licencia = True
            End If
    
            '/* determinar los d¡as que afecta */
    
            Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            
            If Mes_Afecta = 2 Then
                If Anio_bisiesto Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 29, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta, 28, Ano_Afecta)
                End If
            Else
                Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
            End If
            
            If (Fecha_Fin <= Ultimo_Mes) Then
            '/* termina en el mes */
                If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                    Dias_Afecta = Dias_Restantes
                Else
                    Dias_Afecta = DateDiff("d", Primero_Mes, Fecha_Fin) + 1
                End If
            Else
            '/* continua en el mes siguiente */
                Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
            End If
            
        Loop
        
    End If
    
    
    Set rs = Nothing

End Sub

Public Sub Politica1500_V2PagoDescuento_AGD()
    '------------------------------------------------------------------------------------------------------------------------------
    'Descripcion: Paga todo y descuenta por mes con tope a 30 dias comenzando a partir del mes marcado como LiquidaVac
    ' Autor     : Moro Lisandro - 20-02-2007
    'Ult Mod    : Gustavo Ring  - 22-04-2008 - Si no hay liquida escribe en el flog, borra los p/d cuando reprocesa
    '                              12-06-2008 - Se arreglo para que genere todos los dias de descuentos aunque sean en distintos meses
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim Mes_Inicio As Integer
    Dim Ano_Inicio As Integer
    Dim Dia_Inicio As Integer
    Dim Mes_Fin    As Integer
    Dim Ano_Fin    As Integer
    Dim Fin_Licencia  As Boolean
    Dim Mes_Afecta    As Integer
    Dim Ano_Afecta    As Integer
    Dim Primero_Mes   As Date
    Dim Ultimo_Mes    As Date
    Dim Fecha_Fin As Date
    Dim Dias_Afecta   As Integer
    Dim Dias_Pendientes As Integer
    Dim Dias_Restantes As Integer
    Dim Corrio_Dia As Boolean
    Dim GenerarSituacionRevista As Boolean
    Dim l_vdiapeddesde As Date
    
    Dim Aux_TotalDiasCorrespondientes As Integer
    
    Dim Dias_ya_tomados As Integer
    Dim Fecha_limite    As Date
    Dim Anio_bisiesto   As Boolean
    Dim Jornal As Boolean
    Dim Legajo As Long
    Dim NroTer As Long
    Dim Nombre As String
    Dim Aux_Fecha As Date
    Dim Aux_Generar_Fecha_Hasta As Date
    
    Dim Aux_TipDiaPago As Integer
    Dim Aux_TipDiaDescuento As Integer
    
    Dim Quincena_Inicio As Integer
    Dim Quincena_Fin As Integer
    Dim Quincena_Siguiente As Integer
    
    Dim rs As New ADODB.Recordset
    Dim rs_vacdiascor As New ADODB.Recordset
    Dim rs_vacacion As New ADODB.Recordset
    Dim rs_Estructura  As New ADODB.Recordset
    
    Call Politica(1503)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
        TipDiaPago = 3
        TipDiaDescuento = 3
        'Exit Sub
    End If
    
    GenerarSituacionRevista = False
    Call Politica(1509)
    
    StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
    OpenRecordset StrSql, rs_vacacion
    If rs_vacacion.EOF Then
        Flog.writeline "No existe el periodo de vacaciones"
        Exit Sub
    End If
    
    StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
    StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
    StrSql = StrSql & " AND (venc is null OR venc = 0)"
    OpenRecordset StrSql, rs_vacdiascor
    If rs_vacdiascor.EOF Then
        Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
        Exit Sub
    End If
    
    StrSql = " SELECT vdiapeddesde FROM vacdiasped "
    StrSql = StrSql & " WHERE LiquidaVac = -1 "
    StrSql = StrSql & " AND vdiaspedestado = -1 "
    StrSql = StrSql & " AND ternro = " & Ternro
    StrSql = StrSql & " AND vacnro = " & NroVac
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Primero_Mes = rs!vdiapeddesde
        Mes_Inicio = CInt(Month(rs!vdiapeddesde))
        Ano_Inicio = CInt(Year(rs!vdiapeddesde))
        l_vdiapeddesde = CDate(Day(rs!vdiapeddesde))
    Else
        Flog.writeline "No se encontro el Pedido de Vacaciones Marcado como Liquida Vacaciones o no se encuentra aprobado"
        Exit Sub
    End If
    rs.Close

    'Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
    'Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
    If Not EsNulo(fecha_hasta) Then
        Mes_Fin = Month(fecha_hasta)
        Ano_Fin = Year(fecha_hasta)
    Else
        Mes_Fin = Month(fecha_desde)
        Ano_Fin = Year(fecha_desde)
    End If
    
    'Reviso que la cantidad de dias correspondientes del periodo
    'no supere la cantidad de dias que quedan por tomar
    Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
    If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
        Dias_Afecta = Total_Dias_A_Generar
    Else
        Dias_Afecta = rs_vacdiascor!vdiascorcant
    End If
    Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
    Aux_Generar_Fecha_Hasta = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
    
    Genera_Pagos = False
    Genera_Descuentos = False
    Call Politica(1507)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
        Genera_Pagos = True
        Genera_Descuentos = True
    End If
    
    If Not Reproceso Then
        'si no es reproceso y existe el desglose de pago/descuento, salir
        StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        rs.Open StrSql, objConn
        If Not rs.EOF Then
            If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
                'Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
                'Restauro la cantidad de dias porque no se generaron
                Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            End If
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            Exit Sub
        End If
        rs.Close
    Else
        ' Existe reproceso => hay que borrar los pagos descuentos anteriores
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
                StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
                StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    Aux_TipDiaPago = TipDiaPago
    Aux_TipDiaDescuento = TipDiaDescuento
    
    If Genera_Pagos Then
        Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
    End If
    
   
    'Se configura por la politica 1509 si se genera la situacion de revista junto al pago descuento
    If GenerarSituacionRevista And Dias_Afecta <> 0 Then
        InsertarSituacionRevista Ternro, l_vdiapeddesde, DateAdd("d", Dias_Afecta, l_vdiapeddesde)
    End If
    
   
    Genera_Dto_DiasCorr = False
    Call Politica(1506)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
    End If
    
    If Genera_Dto_DiasCorr Then
    
        Fin_Licencia = False
        Mes_Afecta = Mes_Inicio
        Ano_Afecta = Ano_Inicio
        Dias_Pendientes = 0
        Dias_Restantes = Dias_Afecta
        
        'determinar los dias que afecta para el primer mes de decuento
        'Genera 30 dias, para todos los meses
        
        Fecha_Fin = DateAdd("d", Dias_Afecta - 1, Primero_Mes)
        Mes_Fin = Month(Fecha_Fin)
        Ano_Fin = Year(Fecha_Fin)
        
        If Mes_Afecta = 2 Then
            Anio_bisiesto = EsBisiesto(Ano_Afecta)
            If Anio_bisiesto Then
                Ultimo_Mes = AFecha(Mes_Afecta, 29, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta, 28, Ano_Afecta)
            End If
        Else
            Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
        End If
    
        If (Fecha_Fin <= Ultimo_Mes) Then
            '/* termina en el mes */
            Dias_Afecta = Dias_Afecta
        Else
            '/* continua en el mes siguiente */
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If
        
        Do While Not Fin_Licencia Or (Dias_Restantes > 0)
        
            Dias_Restantes = Dias_Restantes - Dias_Afecta
            
            If Genera_Descuentos And Dias_Afecta > 0 Then
                Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            End If
            
            '/* determinar a que continua en el proximo mes */
            If (Mes_Afecta = 12) Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                Mes_Afecta = Mes_Afecta + 1
            End If
    
            If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
                Fin_Licencia = True
            End If
    
            '/* determinar los d¡as que afecta */
    
            Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            
            If Mes_Afecta = 2 Then
                If Anio_bisiesto Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 29, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta, 28, Ano_Afecta)
                End If
            Else
                Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
            End If
            
            If (Fecha_Fin <= Ultimo_Mes) Then
            '/* termina en el mes */
                If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                    Dias_Afecta = Dias_Restantes
                Else
                    Dias_Afecta = DateDiff("d", Primero_Mes, Fecha_Fin) + 1
                End If
            Else
            '/* continua en el mes siguiente */
                Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
            End If
        
        Loop
        
        'tantos descuentos como meses afecte
'        Fin_Licencia = False
 '       Mes_Afecta = Mes_Inicio
  '      Ano_Afecta = Ano_Inicio
   '     Dias_Pendientes = 0
    '    Dias_Restantes = Dias_Afecta


        'Do While Not Fin_Licencia And (Dias_Restantes > 0)

         '   If Dias_Afecta > 30 Then
          '      Dias_Restantes = Dias_Afecta - 30
           '     Dias_Afecta = 30
          '  Else
           '     Dias_Afecta = Dias_Restantes
            '    Dias_Restantes = 0
         '   End If
'
'            'Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'Pago
          '  If Genera_Descuentos And (Dias_Afecta <> 0) Then
           '     Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            'End If
'
        '    If (Mes_Afecta = 12) Then
        '        Mes_Afecta = 1
        '        Ano_Afecta = Ano_Afecta + 1
        '    Else
        '        Mes_Afecta = Mes_Afecta + 1
        '    End If
'
'            'termina en el mes
        '    If (Dias_Restantes < 30) Then
        '        Dias_Afecta = Dias_Restantes
        '    Else
'                Dias_Afecta = DateDiff("d", Primero_Mes, elfechahasta) + 1
'                Dias_Restantes = Dias_Restantes - 30
'                Dias_Afecta = 30
       '     End If
'
        '    If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
         '       Fin_Licencia = True
          '  End If
'
       ' Loop

    Else
        Flog.writeline "No se generan los descuentos."
    End If
    
    'Cierro todo
    If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
    If rs.State = adStateOpen Then rs.Close
    If rs_vacacion.State = adStateOpen Then rs_vacacion.Close
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    
    Set rs_vacdiascor = Nothing
    Set rs = Nothing
    Set rs_vacacion = Nothing
    Set rs_Estructura = Nothing

End Sub

Public Sub Politica1500v_16()
'***************************************************************
'  EN CUANTO SE TOMA AL MENOS UN DIA DE VACACIONES, SE LE PAGA Y DESCUENTA TODOS LOS D™AS DE
'  VACACIONES QUE LE CORRESPONDEN PARA EL A¾O
'  TODO ESTO LO HACE MIENTRAS NO PAGO NI DESCONTO NINGUN DIA DE DICHA LICENCIA
'***************************************************************
' FAF - 11/08/2006 - Es copia de Politica1500v_7() con la unica diferencia que topea a 30 dias
Dim rs As New Recordset
Dim StrSql As String

Dim Dia_Inicio As Integer
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Dia_Fin    As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim i As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Dias_Hab_Afecta As Integer
Dim objFeriado      As New Feriado
Dim EsFeriado       As Boolean
Dim qna_afecta As Integer
Dim estrnro_GrupoLiq As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Integer
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String

Dim ya_se_pago As Boolean
Dim cantdias As Integer
Dim blv As New ADODB.Recordset
Dim bel As New ADODB.Recordset

Set objFeriado.Conexion = objConn

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
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
    NroTer = rs!Empleado
    Dia_Inicio = Day(rs!elfechadesde)
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Dia_Fin = Day(rs!elfechahasta)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If
rs.Close

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE FROM vacpagdesc"
        StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND (pago_dto = 3" 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


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
Dias_Afecta = rs!elcantdias

If Genera_Pagos Then
    If Pliq_Nro <> 0 Then
        Mes_Inicio = Pliq_Mes
        Ano_Inicio = Pliq_Anio
    End If
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
End If
rs.Close



    StrSql = "SELECT estructura.estrnro, estructura.estrcodext FROM estructura "
    StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " INNER JOIN emp_lic ON his_estructura.ternro = emp_lic.empleado "
    StrSql = StrSql & " WHERE his_estructura.tenro = 22 AND his_estructura.htethasta IS NULL "
    StrSql = StrSql & " AND emp_lic.emp_licnro= " & nrolicencia
    rs.Open StrSql, objConn
    If rs.EOF Then
        Jornal = 1  ' mensual
        estrnro_GrupoLiq = 0
        Flog.writeline "    El empleado no posee Forma de Liquidación. Se considera Mensual."
    Else
        estrnro_GrupoLiq = rs!estrnro
        Select Case rs!estrcodext
            Case "1":  ' Mensual
                Jornal = 1
                Flog.writeline "    Según el Cód. Ext. (1) de la Forma de Liquidación del empleado, se considera Mensual."
            Case "2":  ' Quincenal
                Jornal = 2
                Flog.writeline "    Según el Cód. Ext. (2) de la Forma de Liquidación del empleado, se considera Quincenal."
            Case "3":  ' Semanal
                Jornal = 2
                Flog.writeline "    Según el Cód. Ext. (3) de la Forma de Liquidación del empleado, se considera Semanal. Se procesa como Quincenal."
            Case "4":  ' Quincenal sin Prop. Topes
                Jornal = 2
                Flog.writeline "    Según el Cód. Ext. (4) de la Forma de Liquidación del empleado, se considera Quincenal sin Prop. Topes. Se procesa como Quincenal."
            Case Else  ' Mensual
                Jornal = 1
                Flog.writeline "    No está definido el Cód. Ext. de la Forma de Liquidación del empleado, se considera Mensual."
        End Select
    End If
    rs.Close
    
    
    StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
    rs.Open StrSql, objConn
    
    Flog.writeline "    Días de la licencia: " & Dias_Afecta
    
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Afecta
    
    qna_afecta = 3
    If Jornal = 2 Then
        If Dia_Inicio >= 16 Then
            qna_afecta = 2
        Else
            qna_afecta = 1
        End If
    End If
    Flog.writeline "    Modelo de Liquidación (1.- Primera Quincena, 2.- Segunda Quincena, 3.- Mensuales): " & qna_afecta
    
    '/* determinar los días que afecta para el primer mes de decuento */
    '/* Genera 30 dias mínimos, para todos los meses */
'    Anio_bisiesto = EsBisiesto(Ano_Afecta)

    
    If qna_afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
    Else
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    End If
        
    If Mes_Afecta = 2 Then
        If qna_afecta = 3 Then
            If Anio_bisiesto Then
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
            End If
        Else
            If qna_afecta = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
            End If
        End If
    Else
        If qna_afecta = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            If qna_afecta = 3 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
            Else
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                End If
            End If
        End If
    End If

    If (rs!elfechahasta <= Ultimo_Mes) Then
        '/* termina en el mes */
        Dias_Afecta = rs!elcantdias
    Else
        '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
    End If
    
    Do While Not Fin_Licencia
        
        Flog.writeline "    Días entre el " & Primero_Mes & " y el " & Ultimo_Mes & ": " & Dias_Afecta

        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
'        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'PAgo
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, qna_afecta, Dias_Afecta, 4) 'PAgo
         
        Flog.writeline "    Días restantes: " & Dias_Restantes
        
        '/* determinar a que continua en el proximo mes */
        If qna_afecta = 1 Then
            qna_afecta = 2
        Else
            If qna_afecta = 2 Then
                qna_afecta = 1
            End If
            If (Mes_Afecta = 12) Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                Mes_Afecta = Mes_Afecta + 1
            End If
        End If
        
        If (Dias_Restantes <= 0) Then
            Fin_Licencia = True
        End If

        '/* determinar los días que afecta */
        Anio_bisiesto = EsBisiesto(Ano_Afecta)

        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        
        If qna_afecta = 3 And Mes_Afecta = 3 Then
            If Anio_bisiesto Then
                Primero_Mes = AFecha(Mes_Afecta, 2, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 3, Ano_Afecta)
            End If
        Else
            If qna_afecta = 2 Then
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
            End If
        End If
            
        If Mes_Afecta = 2 Then
            If qna_afecta = 3 Then
                If Anio_bisiesto Then
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
                End If
            Else
                If qna_afecta = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                End If
            End If
        Else
            If qna_afecta = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                If qna_afecta = 3 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
                Else
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                    End If
                End If
            End If
        End If

        
        If (rs!elfechahasta <= Ultimo_Mes) Then
            '/* termina antes del ultimo dia del mes */
            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
            End If
        Else
            '/* continua en el mes siguiente */
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If
    
        If qna_afecta = 3 Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        End If
    Loop

If rs.State = adStateOpen Then rs.Close
Set rs = Nothing

End Sub

Public Sub Politica1500_v16_DiasCorresp()
'-----------------------------------------------------------------------
'Custom SMT - Standart
'Genera Pago/dto por dias correspondientes
'paga adelantado todo y descuenta por mes s/Forma de Liquidacion al que pertenece (quincenal o mensual).
'Particiona los dias y los asigna al modelo de Liquidacion acorde
'Fecha: 15/05/2007
'Autor: FAF
'-----------------------------------------------------------------------
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim Jornal As Integer
Dim Dias_Afecta As Integer
Dim Dia_Inicio As Integer
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin As Integer
Dim Ano_Fin As Integer
Dim Aux_TotalDiasCorrespondientes As Long
Dim fecDesde As Date
Dim fecHasta As Date
Dim estrnro_GrupoLiq As Integer
Dim Dias_Hab_Afecta As Integer
Dim i As Date
Dim objFeriado      As New Feriado
Dim EsFeriado       As Boolean
Dim qna_afecta As Integer
Dim Fin_Licencia As Boolean
Dim Mes_Afecta As Integer
Dim Ano_Afecta As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Anio_bisiesto As Boolean
Dim Primero_Mes As Date
Dim Ultimo_Mes As Date

'Dim Dias_ya_tomados As Integer
'Dim Fecha_limite    As Date
'Dim Anio_bisiesto   As Boolean

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Set objFeriado.Conexion = objConn

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "    No existe el período de vacaciones " & NroVac
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "    No hay días correspondientes a ese período. SQL: " & StrSql
    Exit Sub
End If

'Mes_Inicio = Month(rs_Vacacion!vacfecdesde)
'Ano_Inicio = Year(rs_Vacacion!vacfecdesde)
'Mes_Fin = Month(rs_Vacacion!vacfechasta)
'Ano_Fin = Year(rs_Vacacion!vacfechasta)

Dia_Inicio = Day(Aux_Generar_Fecha_Desde)
Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If


'FGZ - 14/10/2005
'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
    
PoliticaOK = False
Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "  Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "  Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'-----------------------------------------------------------------------------------------
'VERIFICAR el reproceso y manejar la depuracion
'-----------------------------------------------------------------------------------------
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "  No se seteo el Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "  Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "  No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


'-----------------------------------------------------------------------------------------
'Comienzo de generacion del Pago y del Descuento
'-----------------------------------------------------------------------------------------
    StrSql = "SELECT estructura.estrnro, estructura.estrcodext FROM estructura "
    StrSql = StrSql & " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE his_estructura.tenro = 32 AND his_estructura.htethasta IS NULL "
    StrSql = StrSql & " AND his_estructura.ternro= " & Ternro
    rs.Open StrSql, objConn
    If rs.EOF Then
        Jornal = 1  ' mensual
        estrnro_GrupoLiq = 0
        Flog.writeline "    El empleado no posee Forma de Liquidación. Se considera Mensual."
    Else
        estrnro_GrupoLiq = rs!estrnro
        Select Case rs!estrcodext
            Case "1":  ' Mensual
                Jornal = 1
                Flog.writeline "    Según el Cód. Ext. (1) de la Forma de Liquidación del empleado, se considera Mensual."
            Case "2":  ' Quincenal
                Jornal = 2
                Flog.writeline "    Según el Cód. Ext. (2) de la Forma de Liquidación del empleado, se considera Quincenal."
            Case "3":  ' Semanal
                Jornal = 2
                Flog.writeline "    Según el Cód. Ext. (3) de la Forma de Liquidación del empleado, se considera Semanal. Se procesa como Quincenal."
            Case "4":  ' Quincenal sin Prop. Topes
                Jornal = 2
                Flog.writeline "    Según el Cód. Ext. (4) de la Forma de Liquidación del empleado, se considera Quincenal sin Prop. Topes. Se procesa como Quincenal."
            Case Else  ' Mensual
                Jornal = 1
                Flog.writeline "    No está definido el Cód. Ext. de la Forma de Liquidación del empleado, se considera Mensual."
        End Select
    End If
    rs.Close
    
    If Genera_Descuentos Then
        If Pliq_Nro <> 0 Then
            Mes_Inicio = Pliq_Mes
            Ano_Inicio = Pliq_Anio
        End If
        Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
    End If
    
    fecDesde = Aux_Generar_Fecha_Desde
    fecHasta = DateAdd("d", Dias_Afecta, fecDesde)
    
    'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
    Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

    Genera_Dto_DiasCorr = False
    Call Politica(1506)
    If Not PoliticaOK Then
        Flog.writeline "  Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
    End If
    
    If Genera_Dto_DiasCorr Then
        
        Flog.writeline "    Días de la licencia: " & Dias_Afecta
        
        Fin_Licencia = False
        Mes_Afecta = Mes_Inicio
        Ano_Afecta = Ano_Inicio
        Dias_Pendientes = 0
        Dias_Restantes = Dias_Afecta
        
        qna_afecta = 3
        If Jornal = 2 Then
            If Dia_Inicio >= 16 Then
                qna_afecta = 2
            Else
                qna_afecta = 1
            End If
'        Else
'            If Jornal = 1 Then
'                qna_afecta = 3
'            Else
'                qna_afecta = 5
'            End If
        End If
        Flog.writeline "    Modelo de Liquidación (1.- Primera Quincena, 2.- Segunda Quincena, 3.- Mensuales): " & qna_afecta
        
        '/* determinar los días que afecta para el primer mes de decuento */
        '/* Genera 30 dias mínimos, para todos los meses */
'        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        
        If qna_afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
        Else
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        End If
            
        If Mes_Afecta = 2 Then
            If qna_afecta = 3 Then
                If Anio_bisiesto Then
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
                End If
            Else
                If qna_afecta = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                End If
            End If
        Else
            If qna_afecta = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                If qna_afecta = 3 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
                Else
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                    End If
                End If
            End If
        End If
        
        If (fecHasta <= Ultimo_Mes) Then
            '/* termina en el mes */
'            Dias_Afecta = rs!elcantdias
        Else
            '/* continua en el mes siguiente */
            Dias_Afecta = DateDiff("d", fecDesde, Ultimo_Mes) + 1
        End If
        
        
        Do While Not Fin_Licencia
            
            Flog.writeline "    Días entre el " & Primero_Mes & " y el " & Ultimo_Mes & ": " & Dias_Afecta
    
            Dias_Restantes = Dias_Restantes - Dias_Afecta
            
    '        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, TipDiaDescuento, Dias_Afecta, 4) 'PAgo
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, qna_afecta, Dias_Afecta, 4) 'PAgo
             
            Flog.writeline "    Días restantes: " & Dias_Restantes
            
            '/* determinar a que continua en el proximo mes */
            If qna_afecta = 1 Then
                qna_afecta = 2
            Else
                If qna_afecta = 2 Then
                    qna_afecta = 1
                End If
                If (Mes_Afecta = 12) Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    Mes_Afecta = Mes_Afecta + 1
                End If
            End If
            
            If (Dias_Restantes <= 0) Then
                Fin_Licencia = True
            End If
    
            '/* determinar los días que afecta */
'            Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
            Anio_bisiesto = EsBisiesto(Ano_Afecta)

            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        
            If qna_afecta = 3 And Mes_Afecta = 3 Then
                If Anio_bisiesto Then
                    Primero_Mes = AFecha(Mes_Afecta, 2, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 3, Ano_Afecta)
                End If
            Else
                If qna_afecta = 2 Then
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                End If
            End If
            
            If Mes_Afecta = 2 Then
                If qna_afecta = 3 Then
                    If Anio_bisiesto Then
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
                    End If
                Else
                    If qna_afecta = 1 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                    Else
                        Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                    End If
                End If
            Else
                If qna_afecta = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    If qna_afecta = 3 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 30, Ano_Afecta)
                    Else
                        If Mes_Afecta = 12 Then
                            Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                        Else
                            Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                        End If
                    End If
                End If
            End If
            
'            If (rs!elfechahasta <= Ultimo_Mes) Then
                '/* termina antes del ultimo dia del mes */
'                Dias_Afecta = DateDiff("d", Primero_Mes, rs!elfechahasta) + 1
'            Else
                '/* continua en el mes siguiente */
                Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
'            End If
        
            If qna_afecta = 3 Then
                If Dias_Afecta > 30 Then
                   Dias_Afecta = 30
                End If
            End If
        Loop
    Else
        Flog.writeline "  No se generan los descuentos (VER Politica 1506)."
    End If
    
    
    
' Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
End Sub

Public Sub Politica1500_AdelantaDescuenta_ARLEI()
'-----------------------------------------------------------------------
'Customizacion para ARLEI.
'Anticipa completo al inicio de la licencia y descuenta por Quincenas lo que corresponde segun la licencia.
'Fecha: 24/10/2007
'Autor: Lisandro Moro
'st_ModeloPago2
'st_ModeloDto2
'
'-----------------------------------------------------------------------
Dim Mes_Inicio      As Integer
Dim Ano_Inicio      As Integer
Dim Mes_Fin         As Integer
Dim Ano_Fin         As Integer
Dim Fin_Licencia    As Boolean
Dim Mes_Afecta      As Integer
Dim Ano_Afecta      As Integer
Dim Primero_Mes     As Date
Dim Ultimo_Mes      As Date
Dim dias            As Integer
Dim Diasanti        As Integer
Dim Dias_Afecta     As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes  As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal          As Boolean
Dim Legajo          As Long
Dim NroTer          As Long
Dim Nombre          As String
Dim objFeriado      As New Feriado
Dim EsFeriado       As Boolean
Dim Dia_Analisis    As Date
Dim Dia_Hasta       As Date
Dim Aux_Fecha       As Date
Dim Topeanti        As Integer 'INITIAL 25. /* Dias de tope para anticipo vac. */
Dim Primero         As Boolean
Dim rs              As New ADODB.Recordset
Dim rs_Estructura   As New ADODB.Recordset
Dim rs_2            As New ADODB.Recordset
Dim blnPagoRealizado As Boolean
Dim blnPraQuincena     As Boolean
'Dim blnSdaQuincena     As Boolean
Dim blnCambioQuincena     As Boolean

blnPagoRealizado = False
blnPraQuincena = False
'blnSdaQuincena = False
blnCambioQuincena = False

'Dim Dia_inicio As Date
Dim Aux_TotalDiasCorrespondientes As Long

Topeanti = 25 'Dias de tope para anticipo vac.
Primero = True

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If
If st_ModeloDto2 = "" Then
    st_ModeloDto2 = st_ModeloDto
End If
If st_ModeloPago2 = "" Then
    st_ModeloPago2 = st_ModeloPago
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
    Aux_Fecha = rs!elfechadesde
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If
rs.Close

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507) 'Si genera pagos y/o descuentos (Ambos)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If
'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        rs.Close
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE FROM vacpagdesc"
        StrSql = StrSql & " WHERE emp_licnro = " & nrolicencia
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND (pago_dto = 3" 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
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

'Busco la forma de liquidacion
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
        Else
            Jornal = False
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQl: " & StrSql
    Jornal = False
End If


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

Mes_Inicio = Month(rs!elfechadesde)
Ano_Inicio = Year(rs!elfechadesde)
Mes_Fin = Month(rs!elfechahasta)
Ano_Fin = Year(rs!elfechahasta)
dias = IIf(Month(rs!elfechadesde) = 12, 31, Day(CDate("01/" & Month(rs!elfechadesde) + 1 & "/" & Year(rs!elfechadesde)) - 1))
Dias_Restantes = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1
  'Dias_Afecta = DateDiff("d", rs!elfechadesde, rs!elfechahasta) + 1

'Busca la cantidad de dias correspondientes al tercero y las vacaciones
Aux_TotalDiasCorrespondientes = 0
StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_2
If rs_2.EOF Then
    Flog.writeline "No hay dias correspondientes al periodo" & NroVac & " y tercero " & Ternro
    Exit Sub
Else
    Aux_TotalDiasCorrespondientes = CLng(rs_2!vdiascorcant)
End If
rs_2.Close

'Verifico se ya se realizo el pago
StrSql = " SELECT sum(cantdias) cant FROM vacpagdesc "
StrSql = StrSql & " WHERE ternro = " & Ternro
StrSql = StrSql & " AND vacnro = " & NroVac
StrSql = StrSql & " AND pago_dto = 3 "
OpenRecordset StrSql, rs_2
If rs_2.EOF Then
    blnPagoRealizado = False
Else
    Dim l_cantdias As Long
    If IsNull(rs_2("cant")) Then l_cantdias = 0 Else l_cantdias = CLng(rs_2("cant"))
        
    If l_cantdias < Aux_TotalDiasCorrespondientes Then
        blnPagoRealizado = False
    Else
        blnPagoRealizado = True
    End If
End If
rs_2.Close
' dias pendientes

'-------------------------------------------------------------------------
'ANTICIPO
Set objFeriado.Conexion = objConn
Set objFeriado.ConexionTraza = objConn

If Month(rs!elfechadesde) <> Month(rs!elfechahasta) Then
    Dia_Analisis = rs!elfechadesde
    Dia_Hasta = CDate(dias & "/" & Month(rs!elfechadesde) & "/" & Year(rs!elfechadesde))
    Do While Dia_Analisis <= Dia_Hasta
        EsFeriado = objFeriado.Feriado(Dia_Analisis, Ternro, False)
        If Not EsFeriado And Weekday(Dia_Analisis) <> 7 And Weekday(Dia_Analisis) <> 1 Then
            Dias_Afecta = Dias_Afecta + 1
        End If
        Dia_Analisis = Dia_Analisis + 1
        Dias_Restantes = Dias_Restantes - 1
    Loop
    Diasanti = Dias_Afecta
    If Diasanti > Topeanti Then
        Diasanti = Topeanti
    End If
    Fin_Licencia = False
Else
    Dia_Analisis = rs!elfechadesde
    Dia_Hasta = rs!elfechahasta
    Do While Dia_Analisis <= Dia_Hasta
        EsFeriado = objFeriado.Feriado(Dia_Analisis, Ternro, False)
        If Not EsFeriado And Weekday(Dia_Analisis) <> 7 And Weekday(Dia_Analisis) <> 1 Then
            Dias_Afecta = Dias_Afecta + 1
        End If
        Dia_Analisis = Dia_Analisis + 1
        Dias_Restantes = Dias_Restantes - 1
    Loop
    Diasanti = Dias_Afecta
    If Diasanti > Topeanti Then
        Diasanti = Topeanti
    End If
End If


If Jornal Then 'Pago segun la quincena
    If Day(rs!elfechadesde) > 15 Then 'Si es segunda quincena
        blnPraQuincena = False
        TipDiaPago = st_ModeloPago2
        TipDiaDescuento = st_ModeloDto2
    Else 'Si es primera quincena
        blnPraQuincena = True
        TipDiaPago = st_ModeloPago
        TipDiaDescuento = st_ModeloDto
    End If
Else 'Pago al mes
    blnPraQuincena = False
    TipDiaPago = st_ModeloPago
    TipDiaDescuento = st_ModeloDto
End If

'If generar_pago Then
'''''    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Diasanti, 1)  'Pago
'End If


    'RUN generar-pago (mes-inicio,ano-inicio,diasanti,1).
'xxx    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Diasanti, 1)  'Pago

    'FIN ANTICIPO
    '-------------------------------------------------------------------------
    
    '-------------------------------------------------------------------------
    'PAGO Y DESCUENTO
If Not blnPagoRealizado Then
    'RUN generar-pago (mes-inicio,ano-inicio,dias-afecta,3).
    'Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Aux_TotalDiasCorrespondientes, 3) 'Pago    Dias_Restantes
End If

    'RUN generar-descuento (mes-inicio,ano-inicio,dias-afecta,4).
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaDescuento, Dias_Afecta, 4) 'Descuento


Mes_Inicio = Month(Dia_Analisis)
Ano_Inicio = Year(Dia_Analisis)
Do While Not Fin_Licencia
    Primero = False
    Dias_Afecta = 0
    blnCambioQuincena = False
    
    If Dias_Restantes > 0 Then
        If Jornal Then 'Pago segun la quincena
            If Day(Dia_Analisis) > 15 Then 'Si es Segunda quincena
                blnPraQuincena = False
                TipDiaPago = st_ModeloPago2
                TipDiaDescuento = st_ModeloDto2
            Else 'Si es primer quincena
                blnPraQuincena = True
                TipDiaPago = st_ModeloPago
                TipDiaDescuento = st_ModeloDto
            End If
        Else 'Pago al mes
            blnPraQuincena = False
            TipDiaPago = st_ModeloPago
            TipDiaDescuento = st_ModeloDto
        End If
        
        Dia_Hasta = Dia_Analisis + Dias_Restantes - 1
        
        Do While Dia_Analisis <= Dia_Hasta Or blnCambioQuincena
            EsFeriado = objFeriado.Feriado(Dia_Analisis, NroTer, False)
            If Not EsFeriado And Weekday(Dia_Analisis) <> 7 And Weekday(Dia_Analisis) <> 1 Then
                Dias_Afecta = Dias_Afecta + 1
            End If
            Dia_Analisis = Dia_Analisis + 1
            Dias_Restantes = Dias_Restantes - 1
            
            If Jornal Then
                If blnPraQuincena = True Then
                    If Day(Dia_Analisis) > 15 Then
                        blnCambioQuincena = True
                    End If
                Else
                    If Day(Dia_Analisis) <= 15 Then
                        blnCambioQuincena = True
                    End If
                End If
            Else
                blnCambioQuincena = False
            End If
        Loop
        
        'determinar a que continua en el proximo mes
        If blnPraQuincena = False Then
            If (Mes_Inicio = 12) Then
                Mes_Inicio = 1
                Ano_Inicio = Ano_Inicio + 1
            Else
                Mes_Inicio = Mes_Inicio + 1
            End If
        End If

        'RUN generar-pago (mes-inicio,ano-inicio,dias-afecta,1).
'xxx        Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 1) 'Pago
        'RUN generar-pago (mes-inicio,ano-inicio,dias-afecta,3).
        '''''Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
        'RUN generar-descuento (mes-inicio,ano-inicio,dias-afecta,4).
        Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 4) 'Pago
        
    Else
        Fin_Licencia = True
    End If
Loop

rs.Close
Set rs = Nothing
End Sub

Public Sub Politica1500_V2PagaDescuenta_PorMes_Papelbril()
'-----------------------------------------------------------------------
'Descripcion: paga todo en el mes y descuenta por mes lo que corresponde con tope a 30 dias de licencia por mes.
'               Genera a partir de dias correspondientes.
' Autor     : Lisandro Moro
' Fecha     : 14-11-2007
'Ult Mod    :
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer

Dim Aux_TotalDiasCorrespondientes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim rs As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones " & NroVac
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
End If


Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If

'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If
Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta
Aux_Generar_Fecha_Hasta = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            If Genera_Descuentos Then
                StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If

Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaPago = TipDiaDescuento
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = TipDiaPago
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio


'Esto va dentro del ciclo en esta version
'If Genera_Pagos Then
'    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
'End If


Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

If Genera_Dto_DiasCorr Then

    'tantos descuentos como meses afecte
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Afecta
    
    'determinar los dias que afecta para el primer mes de decuento
    'Genera 30 dias, para todos los meses
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If
    
    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta)
    Else
        'continua en el mes siguiente
'        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes)
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
    End If
    
    'Genero el pago completo
    If Genera_Pagos Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Restantes, 3) 'Pago
    End If
    
    
    Do While Not Fin_Licencia And Dias_Restantes > 0
        If Dias_Afecta > 30 Then
            Dias_Afecta = 30
        End If
        
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                Aux_TipDiaPago = TipDiaDescuento
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = TipDiaPago
            End If
        End If
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        'If Genera_Pagos Then
        '    Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
        'End If
        
        If Genera_Descuentos And (Dias_Afecta <> 0) Then
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
            Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
        End If
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                Fin_Licencia = True
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        Fin_Licencia = True
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                Fin_Licencia = True
            End If
        End If
        
        'determinar los dias que afecta
        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
    
        If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
            End If
        Else
            'continua en el mes siguiente
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If
        
        If Not Jornal Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        Else
            If Dias_Afecta > 15 Then
               Dias_Afecta = 15
            End If
        End If
        
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                Aux_TipDiaPago = TipDiaDescuento
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = TipDiaPago
            End If
        End If
    Loop
Else
    Flog.writeline "No se generan los descuentos."
End If


'Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close
If rs_vacacion.State = adStateOpen Then rs_vacacion.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
Set rs_vacacion = Nothing
Set rs_Estructura = Nothing
End Sub


Private Sub politica1512(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Calcula el vencimiento de vacaciones del periodo.
'Fecha: 21/10/2009
'Autor: FGZ
'-----------------------------------------------------------------------

    PoliticaOK = False
    CalculaVencimientos = True
    Call SetearParametrosPolitica(Detalle, ok)
    If ok Then
        PoliticaOK = True
    End If
End Sub



Public Sub DiasVencidos(ByVal Ternro As Long, ByVal NroVac As Long, ByVal NroVacAnterior As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias que vencen en el periodo actual y cuantos se pueden transferir al siguiente periodo.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Opcion As Long
Dim Factor As Double

Dim SaldoPeriodo As Integer
Dim DiasV As Integer
Dim DiasT As Integer
Dim NroTPV As Long
'Dim NroVacAnterior As Long

Dim DiasCorr As Integer
Dim DiasTransf As Integer
Dim DiasTom As Integer
Dim Encontro As Boolean
Dim FechaVencimiento As Date

Dim rs As New ADODB.Recordset
Dim rs_Vac As New ADODB.Recordset
Dim rsDias As New ADODB.Recordset

    Flog.writeline "    Vencimiento de Vacaciones"
    Flog.writeline
    
    Opcion = st_Opcion
    Factor = st_Factor
            
    'Hay una sola version por lo cual por el momento no lo uso ----------
    Select Case Opcion
        Case 1:
        Case Else
    End Select
          
          
    If CalculaVencimientos Then
        Flog.writeline "    Opcion:" & Opcion
        Flog.writeline "    Factor de Division:" & Factor
        Flog.writeline
                
        StrSql = " SELECT * FROM vacacion "
        StrSql = StrSql & " WHERE vacnro = " & NroVacAnterior
        OpenRecordset StrSql, rs_Vac
        If Not rs_Vac.EOF Then
            FechaVencimiento = rs_Vac!vacfechasta
        End If
                        
        If NroVacAnterior <> 0 And FechaVencimiento < Date Then
            'Busco la cantidad de dias que le corresponde
            StrSql = "SELECT vacdiascor.vdiascorcant FROM vacdiascor "
            StrSql = StrSql & " WHERE vacnro = " & NroVacAnterior & " AND Ternro = " & Ternro
            StrSql = StrSql & " AND (venc is null OR venc = 0)"
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                DiasCorr = rs!vdiascorcant
            Else
                DiasCorr = 0
            End If
            'mas los que le fueron trasnferidos del periodo anterior
            StrSql = "SELECT vacdiascor.vdiascorcant FROM vacdiascor "
            StrSql = StrSql & " WHERE vacnro = " & NroVacAnterior & " AND Ternro = " & Ternro
            StrSql = StrSql & " AND (venc = 2)"
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                DiasTransf = rs!vdiascorcant
            Else
                DiasTransf = 0
            End If
            
            Flog.writeline "        Dias Correspondientes del periodo anterior:" & DiasCorr
            
            If DiasCorr > 0 Then
                'Calculo la cantidad de dias que ya gozó
                DiasTom = 0
                StrSql = "SELECT emp_lic.elcantdias FROM lic_vacacion "
                StrSql = StrSql & " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
                StrSql = StrSql & " WHERE lic_vacacion.vacnro = " & NroVacAnterior & " AND emp_lic.empleado = " & Ternro
                OpenRecordset StrSql, rsDias
                Do While Not rsDias.EOF
                    DiasTom = DiasTom + rsDias!elcantdias
                    rsDias.MoveNext
                Loop
                Flog.writeline "        Dias tomados del periodo anterior:" & DiasTom
        
                SaldoPeriodo = DiasCorr + DiasTransf - DiasTom
                
                'Calculo la cantidad de dias que se transfieren a este periodo
                DiasT = Round(DiasCorr * Factor, 0)
                If DiasT > SaldoPeriodo Then
                    DiasT = SaldoPeriodo
                End If
                
                'Calculo los dias que vencen del periodo anterior
                If DiasTom > DiasTransf Then
                    DiasV = SaldoPeriodo - DiasT
                Else
                    DiasV = DiasCorr - DiasT
                End If
                
                If DiasV > SaldoPeriodo Then
                    DiasV = SaldoPeriodo
                End If
            
                Flog.writeline "        Dias que vencen del periodo anterior:" & DiasV
                Flog.writeline "        Dias que se transfieren al periodo actual:" & DiasT
                        
                NroTPV = 0
                If DiasV <> 0 Then
                    StrSql = "SELECT tipvacnro FROM vacdiascor "
                    StrSql = StrSql & " WHERE vacnro = " & NroVacAnterior
                    StrSql = StrSql & " AND ternro = " & Ternro
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        NroTPV = rs!tipvacnro
                    Else
                        NroTPV = 1 ' por default
                    End If
                    
                    'Busco si ya hay dias vencidos
                    StrSql = "SELECT * FROM vacdiascor "
                    StrSql = StrSql & " WHERE vacnro = " & NroVacAnterior & " AND Ternro = " & Ternro
                    StrSql = StrSql & " AND venc = 1 "
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        If Reproceso Then
                            StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & DiasV & ", tipvacnro = " & NroTPV & " WHERE vacnro = " & NroVacAnterior & " AND Ternro = " & Ternro
                            StrSql = StrSql & " AND venc = 1 "
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    Else
                        StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,venc,vdiascormanual,ternro,tipvacnro) VALUES (" & _
                        NroVacAnterior & "," & DiasV & ",1,0," & Ternro & "," & NroTPV & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                Else
                    'por las dudas que se este reprocesando borro si habia venciados
                    StrSql = "DELETE vacdiascor WHERE vacnro = " & NroVacAnterior & " AND Ternro = " & Ternro
                    StrSql = StrSql & " AND venc = 1 "
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                
                If DiasT <> 0 Then
                    If NroTPV = 0 Then
                        StrSql = "SELECT tipvacnro FROM vacdiascor "
                        StrSql = StrSql & " WHERE vacnro = " & NroVac
                        StrSql = StrSql & " AND ternro = " & Ternro
                        OpenRecordset StrSql, rs
                        If Not rs.EOF Then
                            NroTPV = rs!tipvacnro
                        Else
                            NroTPV = 1 ' por default
                        End If
                    End If
                    
                    'Busco si ya hay dias transferidos
                    StrSql = "SELECT * FROM vacdiascor "
                    StrSql = StrSql & " WHERE vacnro = " & NroVac & " AND Ternro = " & Ternro
                    StrSql = StrSql & " AND venc = 2 "
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        If Reproceso Then
                            StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & DiasT & ", tipvacnro = " & NroTPV & " WHERE vacnro = " & NroVac & " AND Ternro = " & Ternro
                            StrSql = StrSql & " AND venc = 2 "
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    Else
                        StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,venc,vdiascormanual,ternro,tipvacnro) VALUES (" & _
                        NroVac & "," & DiasT & ",2,0," & Ternro & "," & NroTPV & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                Else
                    'por las dudas que se este reprocesando borro si habia transferidos
                    StrSql = "DELETE vacdiascor WHERE vacnro = " & NroVac & " AND Ternro = " & Ternro
                    StrSql = StrSql & " AND venc = 2 "
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else
                Flog.writeline "       No se puede calcular, no se encontraron dias correspondientes en el periodo anterior:" & NroVacAnterior
                DiasV = 0
                DiasT = 0
            End If
        Else
            Flog.writeline "       No se puede calcular el vencimiento. La fecha de vencimiento del priodo " & Periodo_Anio & " (" & NroVac & ") no ha vencido."
            'Flog.writeline "       No se puede calcular, no se encontró periodo anterior a " & Periodo_Anio & "(" & NroVac & ")"
        End If
    End If
            
' Cierro todo y libero
If rs.State = adStateOpen Then rs.Close
If rs_Vac.State = adStateOpen Then rs_Vac.Close
If rsDias.State = adStateOpen Then rsDias.Close

Set rs = Nothing
Set rs_Vac = Nothing
Set rsDias = Nothing
End Sub

Public Sub DiasVencidos_Col(ByVal Ternro As Long, ByVal vacanio As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias vencidos en el periodo actual.
' Autor      : Lisandro Moro
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Opcion As Long
Dim Factor As Double

Dim SaldoPeriodo As Integer
Dim DiasV As Integer
Dim DiasT As Integer
Dim NroTPV As Long
'Dim NroVacAnterior As Long

Dim DiasCorr As Integer
Dim DiasTransf As Integer
Dim DiasTom As Integer
Dim Encontro As Boolean
Dim FechaVencimiento As Date
Dim tipoVencimiento As Integer
Dim diasAVencer As Integer
Dim vacnro As Long

Dim rs As New ADODB.Recordset
Dim rs_Vac As New ADODB.Recordset
Dim rsDias As New ADODB.Recordset

    Flog.writeline "    Vencimiento de Vacaciones "
    Flog.writeline
    
    Opcion = st_Opcion
    Factor = st_Factor
    'PolDias = st_Dias
    
    'Hay una sola version por lo cual por el momento no lo uso ----------
    Select Case Opcion
        Case 1:
        Case Else
    End Select
          
    If CalculaVencimientos Then
        Flog.writeline "    Opcion:" & Opcion
        Flog.writeline "    Factor de Division:" & Factor
        Flog.writeline
                
        StrSql = " SELECT * FROM vacacion "
        StrSql = StrSql & " WHERE vacanio = " & vacanio
        StrSql = StrSql & " AND ternro = " & Ternro
        OpenRecordset StrSql, rs_Vac
        If Not rs_Vac.EOF Then
            FechaVencimiento = rs_Vac!vacfechasta
            vacnro = rs_Vac!vacnro
        Else
            vacnro = 0
        End If
        
        'calculo el tipo de vencimientos segun la fecha
        If FechaVencimiento < DateAdd("yyyy", -4, Date) Then
            ' el periodo vencio completamente
            tipoVencimiento = 4
            diasAVencer = 15 ' Configurar por politica
        Else
            If FechaVencimiento < DateAdd("yyyy", -1, Date) Then
                'el priodo vencio parcialmente
                tipoVencimiento = 1
                diasAVencer = 6 ' Configurar por politica
            Else
                'el periodo aun no ha vencido
                tipoVencimiento = 0
                diasAVencer = 0 'Configurar por politica?
            End If
        End If
        
        'FechaAltaEmpleado
        
        'If () Then
        'End If
        
        If vacnro <> 0 Then
        'If NroVacAnterior <> 0 And FechaVencimiento < Date Then
            
            'Busco la cantidad de dias que le corresponde
            StrSql = "SELECT vacdiascor.vdiascorcant FROM vacdiascor "
            StrSql = StrSql & " WHERE vacnro = " & vacnro & " AND Ternro = " & Ternro
            StrSql = StrSql & " AND (venc is null OR venc = 0)"
            OpenRecordset StrSql, rs
            If Not rs.EOF Then
                DiasCorr = rs!vdiascorcant
            Else
                DiasCorr = 0
            End If
            Flog.writeline "        Dias Correspondientes del periodo(" & vacanio & "):" & DiasCorr
            
            ''mas los que le fueron trasnferidos del periodo anterior
            'StrSql = "SELECT vacdiascor.vdiascorcant FROM vacdiascor "
            'StrSql = StrSql & " WHERE vacnro = " & NroVacAnterior & " AND Ternro = " & Ternro
            'StrSql = StrSql & " AND (venc = 2)"
            'OpenRecordset StrSql, rs
            'If Not rs.EOF Then
            '    DiasTransf = rs!vdiascorcant
            'Else
            '    DiasTransf = 0
            'End If
            '
            '
            
            If DiasCorr > 0 Then
                'Calculo la cantidad de dias que ya gozó
                DiasTom = 0
                StrSql = "SELECT emp_lic.elcantdias FROM lic_vacacion "
                StrSql = StrSql & " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro "
                StrSql = StrSql & " WHERE lic_vacacion.vacnro = " & vacnro & " AND emp_lic.empleado = " & Ternro
                'StrSql = StrSql & " WHERE lic_vacacion.vacnro = " & NroVacAnterior & " AND emp_lic.empleado = " & Ternro
                OpenRecordset StrSql, rsDias
                Do While Not rsDias.EOF
                    DiasTom = DiasTom + rsDias!elcantdias
                    rsDias.MoveNext
                Loop
                Flog.writeline "        Dias tomados del periodo (" & vacanio & "):" & DiasTom
        
                SaldoPeriodo = DiasCorr - DiasTom
                'SaldoPeriodo = DiasCorr + DiasTransf - DiasTom
                
                ''Calculo la cantidad de dias que se transfieren a este periodo
                'DiasT = Round(DiasCorr * Factor, 0)
                'If DiasT > SaldoPeriodo Then
                '    DiasT = SaldoPeriodo
                'End If
                
                'Calculo los dias que vencen del periodo -- anterior
                'If DiasTom > DiasTransf Then
                '    DiasV = SaldoPeriodo - DiasT
                'Else
                '    DiasV = DiasCorr - DiasT
                'End If
                
                DiasV = diasAVencer
                If DiasV > SaldoPeriodo Then
                    DiasV = SaldoPeriodo
                End If
            
                Flog.writeline "        Dias que vencen del periodo (" & vacanio & "):" & DiasV
                'Flog.writeline "        Dias que se transfieren al periodo actual:" & DiasT
                        
                NroTPV = 0
                If DiasV <> 0 Then
                    StrSql = "SELECT tipvacnro FROM vacdiascor "
                    StrSql = StrSql & " WHERE vacnro = " & vacnro ' NroVacAnterior
                    StrSql = StrSql & " AND ternro = " & Ternro
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        NroTPV = rs!tipvacnro
                    Else
                        NroTPV = 1 ' por default
                    End If
                    
                    'Busco si ya hay dias vencidos
                    StrSql = "SELECT * FROM vacdiascor "
                    StrSql = StrSql & " WHERE vacnro = " & vacnro & " AND Ternro = " & Ternro
                    StrSql = StrSql & " AND venc = 1 "
                    OpenRecordset StrSql, rs
                    If Not rs.EOF Then
                        If Reproceso Then
                            StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & DiasV & ", tipvacnro = " & NroTPV
                            StrSql = StrSql & " WHERE vacnro = " & vacnro & " AND Ternro = " & Ternro
                            StrSql = StrSql & " AND venc = 1 "
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    Else
                        StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,venc,vdiascormanual,ternro,tipvacnro)"
                        StrSql = StrSql & " VALUES (" & vacnro & "," & DiasV & ",1,0," & Ternro & "," & NroTPV & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                Else
                    'por las dudas que se este reprocesando borro si habia venciados
                    StrSql = "DELETE vacdiascor WHERE vacnro = " & vacnro & " AND Ternro = " & Ternro
                    StrSql = StrSql & " AND venc = 1 "
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                
                'If DiasT <> 0 Then
                '    If NroTPV = 0 Then
                '        StrSql = "SELECT tipvacnro FROM vacdiascor "
                '        StrSql = StrSql & " WHERE vacnro = " & NroVac
                '        StrSql = StrSql & " AND ternro = " & Ternro
                '        OpenRecordset StrSql, rs
                '        If Not rs.EOF Then
                '            NroTPV = rs!tipvacnro
                '        Else
                '            NroTPV = 1 ' por default
                '        End If
                '    End If
                '
                '    'Busco si ya hay dias transferidos
                '    StrSql = "SELECT * FROM vacdiascor "
                '    StrSql = StrSql & " WHERE vacnro = " & NroVac & " AND Ternro = " & Ternro
                '    StrSql = StrSql & " AND venc = 2 "
                '    OpenRecordset StrSql, rs
                '    If Not rs.EOF Then
                '        If Reproceso Then
                '            StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & DiasT & ", tipvacnro = " & NroTPV & " WHERE vacnro = " & NroVac & " AND Ternro = " & Ternro
                '            StrSql = StrSql & " AND venc = 2 "
                '            objConn.Execute StrSql, , adExecuteNoRecords
                '        End If
                '    Else
                '        StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,venc,vdiascormanual,ternro,tipvacnro) VALUES (" & _
                '        NroVac & "," & DiasT & ",2,0," & Ternro & "," & NroTPV & ")"
                '        objConn.Execute StrSql, , adExecuteNoRecords
                '    End If
                'Else
                '    'por las dudas que se este reprocesando borro si habia transferidos
                '    StrSql = "DELETE vacdiascor WHERE vacnro = " & NroVac & " AND Ternro = " & Ternro
                '    StrSql = StrSql & " AND venc = 2 "
                '    objConn.Execute StrSql, , adExecuteNoRecords
                'End If
            Else
                Flog.writeline "       No se puede calcular, no se encontraron dias correspondientes en el periodo anterior:" & vacnro
                DiasV = 0
                DiasT = 0
            End If
        Else
            Flog.writeline "       No se puede calcular el vencimiento. La fecha de vencimiento del priodo " & Periodo_Anio & " (" & NroVac & ") no ha vencido."
            'Flog.writeline "       No se puede calcular, no se encontró periodo anterior a " & Periodo_Anio & "(" & NroVac & ")"
        End If
    End If
            
' Cierro todo y libero
If rs.State = adStateOpen Then rs.Close
If rs_Vac.State = adStateOpen Then rs_Vac.Close
If rsDias.State = adStateOpen Then rsDias.Close

Set rs = Nothing
Set rs_Vac = Nothing
Set rsDias = Nothing
End Sub


Public Function PeriodoCorrespondiente(ByVal Ternro As Long, ByVal Anio As Integer) As Long
Dim l_TienePolAlcance As Boolean
Dim rs As New ADODB.Recordset


'EAM- Verifica si tiene politica de alcance por periodos de vac
StrSql = "SELECT * FROM alcance_testr WHERE tanro= 21"
OpenRecordset StrSql, rs
If Not rs.EOF Then
    l_TienePolAlcance = True
Else
    l_TienePolAlcance = False
End If

If l_TienePolAlcance Then
    'EAM- Busca los periodos de vac con alcance para el empleado
    StrSql = " SELECT DISTINCT vacacion.vacnro, vacdesc, vacfecdesde, vacfechasta" & _
        " FROM  vacacion " & _
        " INNER JOIN vac_estr ON vacacion.vacnro= vac_estr.vacnro " & _
        " INNER JOIN his_estructura ON vac_estr.estrnro = his_estructura.estrnro " & _
        " WHERE  his_estructura.ternro= " & Ternro & _
        " AND vacacion.vacanio= " & Anio & _
        " ORDER BY vacfecdesde ASC "
Else
    StrSql = "SELECT vacacion.vacdesc, vacacion.vacnro, vacacion.vacfecdesde, vacacion.vacfechasta " & _
            "FROM  vacacion " & _
            " WHERE  vacacion.vacanio= " & Anio & _
            " ORDER BY vacfecdesde DESC "

End If
OpenRecordset StrSql, rs
If Not rs.EOF Then
    PeriodoCorrespondiente = rs("vacnro")
Else
    PeriodoCorrespondiente = 0
End If


    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Function
Public Function PeriodoCorrespondienteAlcance(ByVal Ternro As Long, ByVal Anio As Integer, ByVal vac_alcannivel As Integer) As Long
Dim l_TienePolAlcance As Boolean
Dim rs As New ADODB.Recordset
Dim StrSqlAux
' ---------------------------------------------------------------------------------------------
' Descripcion: Valida los períodos para un empleado seguna alcances de política y Alcances de los períodos de Vacaciones
' Autor      : Gonzalez Nicolás
' Fecha      : 07/11/2013
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

'Verifica si tiene politica de alcance por periodos de vac
'StrSql = "SELECT * FROM alcance_testr WHERE tanro= 21"
'OpenRecordset StrSql, rs
'If Not rs.EOF Then
'    l_TienePolAlcance = True
'    StrSqlAux = " INNER JOIN vac_estr ON vacacion.vacnro= vac_estr.vacnro "
'    StrSqlAux = StrSqlAux & " INNER JOIN his_estructura ON vac_estr.estrnro = his_estructura.estrnro "
'    StrSqlAux = StrSqlAux & " WHERE "
'    StrSqlAux = StrSqlAux & "  his_estructura.Ternro = " & Ternro
'    StrSqlAux = StrSqlAux & " AND "
'Else
'    l_TienePolAlcance = False
    StrSqlAux = " WHERE "
'End If

'Alcances (3 Global, 2 Por Estrucutra y 1 Individual) - Configurados por Politica 1515
If vac_alcannivel = 1 Then 'ALCANCE INDIVIDUAL
    StrSql = "SELECT DISTINCT vacacion.vacnro,vacacion.vacanio"
    StrSql = StrSql & ",vacacion.vacfecdesde,vacacion.vacfechasta"
    'StrSql = StrSql &",vac_alcan.vacfecdesde,vac_alcan.vacfechasta "
    StrSql = StrSql & " FROM vacacion"
    'StrSql = StrSql & " INNER JOIN vac_alcan ON vac_alcan.vacnro = vacacion.vacnro"
    '---------------------------
    StrSql = StrSql & StrSqlAux
    '---------------------------
    StrSql = StrSql & "  vacacion.alcannivel = " & vac_alcannivel
    StrSql = StrSql & " AND vacacion.vacanio = " & Anio
    'StrSql = StrSql & " AND vac_alcan.Origen = " & Ternro
    StrSql = StrSql & " ORDER BY vacacion.vacfecdesde ASC"
    'StrSql = StrSql & " ORDER BY vac_alcan.vacfecdesde ASC"
ElseIf vac_alcannivel = 2 Then 'ALCANCE POR ESTRUCTURAS
    'SEGURAMENTE SE DEBA AJUSTAR ESTA QUERY.
    StrSql = "SELECT DISTINCT vacacion.vacnro,vacacion.vacanio,vac_alcan.vacfecdesde,vac_alcan.vacfechasta"
    StrSql = StrSql & " FROM vacacion"
    StrSql = StrSql & " INNER JOIN vac_alcan ON vac_alcan.vacnro = vacacion.vacnro"
    StrSql = StrSql & " INNER JOIN his_estructura H1 ON H1.estrnro = vac_alcan.origen and H1.ternro = " & Ternro
    '---------------------------
    StrSql = StrSql & StrSqlAux
    '---------------------------
    StrSql = StrSql & " vacacion.alcannivel = " & vac_alcannivel & " And vac_alcan.alcannivel = " & vac_alcannivel
    StrSql = StrSql & " And vacacion.vacanio = " & Anio
    StrSql = StrSql & " ORDER BY vac_alcan.vacfecdesde ASC"

Else
    'POR ALCANCE GLOBAL
    StrSql = "SELECT DISTINCT vacacion.vacdesc, vacacion.vacnro,vacacion.vacanio, vacacion.vacfecdesde, vacacion.vacfechasta "
    StrSql = StrSql & " FROM  vacacion "
    
    '---------------------------
    StrSql = StrSql & StrSqlAux
    '---------------------------
    
    StrSql = StrSql & " vacacion.vacanio= " & Anio
    StrSql = StrSql & " AND vacacion.alcannivel = " & vac_alcannivel
    StrSql = StrSql & " ORDER BY vacacion.vacfecdesde DESC "
End If

OpenRecordset StrSql, rs
If Not rs.EOF Then
    fecha_desde = rs!vacfecdesde
    fecha_hasta = rs!vacfechasta
    Periodo_Anio = rs!vacanio
    PeriodoCorrespondienteAlcance = rs("vacnro")
Else
    PeriodoCorrespondienteAlcance = 0
End If
    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Function

Public Function ValidaModeloyVersiones(ByVal Version As String, ByVal TipoProceso As Integer)
    '--------------- VALIDO MODELOS SEGUN POLITICA 1515 | PUEDE TENER ALCANCE POR ESTRUCTURAS --------------
    '*******************************************************************************************************
    '-------------------------- NO AGREGAR MULTILENGUAJE A ESTA FUNCION ------------------------------------
    'TipoProceso --> Codigo del proceso. btprcnro
    'Version     --> Version del proceso
    
    'VALIDO QUE LA POLITICA 1515 ESTE ACTIVA Y CONFIGURADA
    StrSql = " SELECT gti_detpolitica.detpolprograma,gti_alcanpolitica.alcpolnivel"
    StrSql = StrSql & " FROM gti_cabpolitica"
    StrSql = StrSql & " INNER JOIN gti_alcanpolitica ON gti_cabpolitica.cabpolnro = gti_alcanpolitica.cabpolnro"
    StrSql = StrSql & " INNER JOIN gti_detpolitica ON gti_alcanpolitica.detpolnro = gti_detpolitica.detpolnro"
    StrSql = StrSql & " WHERE gti_cabpolitica.cabpolnivel = 1515"
    StrSql = StrSql & " AND gti_alcanpolitica.alcpolestado = -1"
    StrSql = StrSql & " AND gti_cabpolitica.cabpolestado = -1"
    StrSql = StrSql & " ORDER BY gti_alcanpolitica.alcpolnivel ASC"
    OpenRecordset StrSql, rsPolitica
    If Not rsPolitica.EOF Then
        ValidaModeloyVersiones = True
        Do While Not rsPolitica.EOF
            'Valida que no haya inconsistencia de BD con el/los modelo/s a ejectuar.
            If ValidarVBD(Version, TipoProceso, TipoBD, rsPolitica!detpolprograma) = False Then
             '--------- Control de versiones ------
                ValidaModeloyVersiones = False
                Exit Do
            End If
            
            'If rs_Batch_Proceso!alcpolnivel = 1 Then
            'End If
            rsPolitica.MoveNext
        Loop
    Else
        'LA POLITICA 1515 DEBE ESTAR CONFIGURADA PARA QUE CONTINUE EL PROCESO.
        'Flog.writeline "Error cargando configuración de la Política 1515"
        Flog.writeline "Política 1515 NO configurada. (Debe estar activa sólo para Paraguay."
        ValidaModeloyVersiones = False
        HuboErrores = True
    End If
End Function

Public Function PeriodoVenceaFecha(ByVal Ternro As Long, ByVal FechaProc As Date) As Long
Dim l_TienePolAlcance As Boolean
Dim rs As New ADODB.Recordset


'EAM- Verifica si tiene politica de alcance por periodos de vac
StrSql = "SELECT * FROM alcance_testr WHERE tanro= 21"
OpenRecordset StrSql, rs
If Not rs.EOF Then
    l_TienePolAlcance = True
End If

If l_TienePolAlcance Then
    'Busca los periodos de vac con alcance para el empleado
    StrSql = " SELECT DISTINCT vacacion.vacnro, vacdesc, vacfecdesde, vacfechasta"
    StrSql = StrSql & " FROM  vacacion "
    StrSql = StrSql & " INNER JOIN vac_estr ON vacacion.vacnro= vac_estr.vacnro "
    StrSql = StrSql & " INNER JOIN his_estructura ON vac_estr.estrnro = his_estructura.estrnro "
    StrSql = StrSql & " WHERE  his_estructura.ternro= " & Ternro
    StrSql = StrSql & " AND vacacion.vacfechasta < " & ConvFecha(FechaProc)
    StrSql = StrSql & " ORDER BY vacfecdesde DESC "
Else
    StrSql = " SELECT DISTINCT vacacion.vacnro, vacdesc, vacfecdesde, vacfechasta"
    StrSql = StrSql & " FROM  vacacion "
    StrSql = StrSql & " WHERE  vacacion.vacfechasta < " & ConvFecha(FechaProc)
    StrSql = StrSql & " ORDER BY vacfecdesde DESC "
End If
OpenRecordset StrSql, rs
If Not rs.EOF Then
    PeriodoVenceaFecha = rs("vacnro")
Else
    PeriodoVenceaFecha = 0
End If
    
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Function

Private Sub politica1513(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Activa el calculo de días efectivamente trabajados en el ultimo año.
'Fecha: 06/07/2010
'Autor: Margiotta, Emanuel
'-----------------------------------------------------------------------
Dim Opcion As Long

    PoliticaOK = True
    Dias_efect_trab_anio = True
    
    Call SetearParametrosPolitica(Detalle, ok)
    
    If ok Then
        Opcion = st_Opcion
        
        Select Case Opcion
            Case 1:
                Lic_Descuento = st_ListaF1
            Case 2: 'Esta versión es configrable si se contemplan los feriados o no.
                Lic_Descuento = st_ListaF1
            Case Else
        End Select
    Else
        PoliticaOK = False
    End If
    
   
    
    
    
End Sub


Private Sub politica1514(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Activa la Bonificación de días de Vacaciones CR.
'Fecha: 06/06/2011
'Autor: Margiotta, Emanuel
'Modificado: 08/05/2012 - Gonzalez Nicolás - Se versiona para distintos países
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
' 1 - COSTA RICA - SYKES
' 2 - PORTUGAL
    'Flog.Writeline "POL 1514 subn: " & subn
    'Flog.Writeline "POL 1514 cabecera: " & cabecera
    'Flog.Writeline "POL 1514 Det: " & Detalle
    Call SetearParametrosPolitica(Detalle, ok)
        
    If ok Then
        
        'GUARDO LA VERSION DE LA POLÍTICA PARA LUEGO APLICAR EL PLUS
        st_Opcion = subn
       
        Select Case subn
            Case 1:
                '::::: COSTA RICA ::::::
                'Configurar como --> NO CONFIGURABLE
                PoliticaOK = True
                Dias_Bonificacion = True
            Case 2:
                '::::: PORTUGAL :::::
                'Configurar como --> CONFIGURABLE y se agrega parámetro LISTA
                Lic_Descuento = st_ListaF1
              
                If Lic_Descuento = "" Then
                    'No calcula días o no los toma en cuenta
                    Flog.writeline "    Política 1514: Falta configurar las licencias de exclusión para el cálculo de días de PLUS"
                    PoliticaOK = False
                    Dias_Bonificacion = False
                Else
                    'Activa el cálculo de PLUS
                    PoliticaOK = True
                    Dias_Bonificacion = True
                    Flog.writeline "    Licencias consideradas para el cálculo de PLUS: " & Lic_Descuento
                End If
            
            Case 3:
                PoliticaOK = True
                Dias_Bonificacion = True
            Case Else
                PoliticaOK = False
                Dias_Bonificacion = False
                Flog.writeline "Error de configuración de política 1514"
        End Select
    End If

    
End Sub
Private Sub politica1515(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)

'-----------------------------------------------------------------------
'Descripcion: Valida el modelo de vacaciones
'Fecha: 07/11/2013
'Autor: Gonzalez Nicolás
'Modificado: 05/12/2013 - Gonzalez Nicolás - Se setea en false la politica cuando no tiene parametros.
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
    Dim Modelo
    Dim DescAlcannivel As String
    PoliticaOK = False
    Call SetearParametrosPolitica(Detalle, ok)
    st_Opcion = subn
    If ok Then
        Select Case st_ModeloPais
            Case 0:
                Modelo = EscribeLogMI("Argentina")
                DescAlcannivel = "Períodos de Vacaciones Globales"
                alcannivel = 3
            Case 1:
                Modelo = EscribeLogMI("Uruguay")
                DescAlcannivel = "Períodos de Vacaciones Globales"
                alcannivel = 3
            Case 2:
                Modelo = EscribeLogMI("Chile")
                DescAlcannivel = "Períodos de Vacaciones Globales"
                alcannivel = 3
            Case 3:
                Modelo = EscribeLogMI("Colombia")
                DescAlcannivel = "Períodos de Vacaciones Individuales"
                alcannivel = 1
            Case 4:
                Modelo = EscribeLogMI("Costa Rica")
                DescAlcannivel = "Períodos de Vacaciones Individuales"
                alcannivel = 1
            Case 5:
                Modelo = EscribeLogMI("Portugal")
                DescAlcannivel = "Períodos de Vacaciones Globales"
                alcannivel = 3
            Case 6:
                Modelo = EscribeLogMI("Paraguay")
                DescAlcannivel = "Períodos de Vacaciones Individuales"
                alcannivel = 1
            Case 7:
                Modelo = EscribeLogMI("Peru")
                DescAlcannivel = "Períodos de Vacaciones Individuales"
                alcannivel = 1
        End Select
        
        Flog.writeline "************** " & EscribeLogMI("Modelo") & ": " & UCase(Modelo) & " (" & st_ModeloPais & ") **************"
        Flog.writeline "************** " & EscribeLogMI("Versión") & ":     " & st_Opcion & "   **************"
        Flog.writeline "************** " & DescAlcannivel & "   **************"
        
        PoliticaOK = True
    Else
        PoliticaOK = False
    End If
    
End Sub
Private Sub politica1516(ByVal subn As Long, ByVal cabecera As Long, ByVal Detalle As Long)
'-----------------------------------------------------------------------
'Descripcion: Descuenta días de vacaciones según la configuración de proporción de días de licencia de distintos tipo de licencia.
                'Ejemplo de configuración. lista de licencia que se toman en cuenta. proporcion 1 a 1. significa que cada 1 licencia de la lista se decuenta 1 día de vacacion.
'Fecha: 18/04/2013
'Autor: Margiotta, Emanuel
'Modificado:
    
    Call SetearParametrosPolitica(Detalle, ok)
        
    If ok Then
        
        'GUARDO LA VERSION DE LA POLÍTICA PARA LUEGO APLICAR EL PLUS
        st_Opcion = subn
       
        Select Case subn
            Case 1:
                '::::: COSTA RICA ::::::
                'Configurar como --> NO CONFIGURABLE
                PoliticaOK = True
                Lic_Descuento = st_ListaF1
                DiasProporcion = st_CantidadDias
                
            Case Else
                PoliticaOK = False
                Dias_Bonificacion = False
                Flog.writeline "Error de configuración de política 1514"
        End Select
    End If

    
End Sub

Public Sub Politica1500_Uruguay()
'-----------------------------------------------------------------------
'Descripcion: Version para Monresa
' Autor     : FGZ
' Fecha     : 11/08/2010
'-----------------------------------------------------------------------
'Descripcion
'   Se manejan 3 situaciones
'   A) Dias correspondientes = 20
'   B) Dias correspondientes < 20
'   C) Dias correspondientes > 20
'   ----------------------------------

    'A) Dias correspondientes = 20
    'Si es la 1er salida se pagan siempre 10 días independientemente a los días que son gozados.
    
    'A la 2da salida(o sucesivas) del empleado el sistema debe
        'ir a buscar los días gozados en la primer salida y sumarlos a los días informados para la segunda salida.
        'suma de Días salidas anteriores + Días lic actual= "C"
            'Si "C"  es < a 10 el pago es = 0  (porque gozo menos de 10 y ya cobro 10)
            'Si "C" es > a 10 el pago es = 10  (porque gozo mas de 10 y cobro sólo 10,
            '                                   es el segundo pago y se paga de a 10 días)
    '-------------------------------------------------------------------------------------------------------------
    'B) Dias Correspondientes > 20
    '1er salida(3 posibilidades):
        'Si los días de licencia es > días correspondientes/2, pago los días correspondientes.
        'Si los días de licencia < días correspondientes/2 y > a 10, pago los días reales.
        'Si los días de licencia < días correspondientes/2 y <= a 10, pago 10 días.

    '2da salida
        'Se deberían manejar dos acumuladores:
        'Dr= días reales gozados (sumatoria de todos los días gozados reales)
        'Dp= días pagados (sumatoria de todos los días pagados)
    
        'el sistema debe ir a buscar los días pagados en la primer salida y compararlos
        '   con los días correspondientes (dc), siendo (dp) días pagados en primer salida o salidas anteriores:

        'Si  días correspondientes(dc) - dp es = a cero el pago es cero.
        'Si  días correspondientes(dc) - dp es > 0 y dr es > a dp y > a 10 el pago es = a la diferencia
        '                                                       entre los días correspondientes(dc) - dp.
        'Si  días correspondientes - dp es > 0 y dr es <= a dp y <= a 10 el pago es = cero
    '-------------------------------------------------------------------------------------------------------------
    'C) Dias Correspondientes < 20
    'A la 1er salida pago los días correspondientes totales.
' ---------------------------------------------------------------------------------------
Const Dias_Tope = 10

Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Dias_Afecta As Integer
Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer
Dim Jornal As Boolean
Dim Existe_Pago As Boolean
Dim Primer_Pago As Boolean

Dim Dias_Correspondientes As Integer

Dim Dias_Gozados As Integer
Dim Dias_Pagados As Integer
Dim Dias_Acumulados As Integer

Dim Dias_Lic  As Integer
Dim Dias_Pago As Integer
Dim Dias_Dto  As Integer

Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Aux_Fecha_Desde As Date
Dim Aux_Fecha_Hasta As Date

Dim tipoVac As Long
Dim CHabiles As Integer
Dim cNoHabiles As Integer
Dim cFeriados As Integer

Dim rs As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset
'===================================================================

'Busco la licencia dentro del intervalo especificado
Dias_Lic = 0
StrSql = "SELECT * FROM emp_lic "
StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro"
StrSql = StrSql & " WHERE emp_lic.emp_licnro = " & nrolicencia
OpenRecordset StrSql, rs
If Not rs.EOF Then
    NroVac = rs!vacnro
    'Dias_Lic = rs!elcantdiashab
    Dias_Lic = rs!elcantdias
    Aux_Fecha_Desde = rs!elfechadesde
    Aux_Fecha_Hasta = rs!elfechahasta
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If



Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 23
    TipDiaDescuento = 23
End If
If st_ModeloDto2 = 0 Then
    st_ModeloDto2 = st_ModeloDto
End If
If st_ModeloPago2 = 0 Then
    st_ModeloPago2 = st_ModeloPago
End If

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'Reproceso
If Reproceso Then
    If Not Ya_Pago Then
        'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
        StrSql = "SELECT * FROM vacpagdesc "
        StrSql = StrSql & " WHERE ternro = " & Ternro
        StrSql = StrSql & " AND vacpagdesc.vacnro = " & NroVac
        StrSql = StrSql & " AND vacpagdesc.pronro is not null"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        Else
            StrSql = "DELETE FROM vacpagdesc"
            StrSql = StrSql & " WHERE ternro = " & Ternro
            StrSql = StrSql & " AND vacpagdesc.vacnro = " & NroVac
            If Genera_Pagos And Genera_Descuentos Then
                StrSql = StrSql & " AND (pago_dto = 3" 'pagos
                StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
            Else
                If Genera_Pagos Then
                    StrSql = StrSql & " AND pago_dto = 3" 'pagos
                End If
                If Genera_Descuentos Then
                    StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
                End If
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
Else
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
End If


'Busco los dias correspondientes
StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs
If rs.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
Else
    Dias_Correspondientes = rs!vdiascorcant
    'FGZ - 18/11/2010 -------------------------
    tipoVac = rs!tipvacnro
    'FGZ - 18/11/2010 -------------------------
End If


'Determino si es el primer pago del periodo
Existe_Pago = False
StrSql = "SELECT * FROM vacpagdesc "
StrSql = StrSql & " WHERE ternro = " & Ternro
StrSql = StrSql & " AND vacpagdesc.vacnro = " & NroVac
OpenRecordset StrSql, rs
If Not rs.EOF Then
    'Existe_Pago = True
    Primer_Pago = False
Else
    'Existe_Pago = False
    Primer_Pago = True
End If
'Primer_Pago = Not Ya_Pago Or Existe_Pago

'Busco la cantidad de licencias ya gozadas
Dias_Gozados = 0
StrSql = "SELECT sum(elcantdias) dias FROM emp_lic "
StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro"
StrSql = StrSql & " WHERE empleado = " & Ternro
StrSql = StrSql & " AND emp_lic.emp_licnro <> " & nrolicencia
StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(Aux_Fecha_Desde)
'FGZ - 18/11/2010 ------
StrSql = StrSql & " AND lic_vacacion.vacnro = " & NroVac
'FGZ - 18/11/2010 ------
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Dias_Gozados = IIf(EsNulo(rs!dias), 0, rs!dias)
End If

'Busco los dias ya pagados
Dias_Pagados = 0
StrSql = "SELECT sum(cantdias) dias FROM vacpagdesc "
StrSql = StrSql & " WHERE ternro = " & Ternro
StrSql = StrSql & " AND vacnro = " & NroVac
StrSql = StrSql & " AND pago_dto = 3" 'pagos
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Dias_Pagados = IIf(EsNulo(rs!dias), 0, rs!dias)
End If


'Pago segun..
Select Case Dias_Correspondientes
Case Is = 20
    If Primer_Pago Then
        Dias_Pago = 10
    Else
        If (Dias_Gozados + Dias_Lic) > Dias_Tope Then
            If Dias_Pagados >= Dias_Correspondientes Then
                Dias_Pago = 0
            Else
                Dias_Pago = 10
            End If
        Else
            Dias_Pago = 0
        End If
    End If
Case Is > 20
    If Primer_Pago Then
        If Dias_Lic > (Dias_Correspondientes / 2) Then
            Dias_Pago = Dias_Correspondientes
        Else
            If Dias_Lic > 10 Then
                Dias_Pago = Dias_Lic
            Else
                Dias_Pago = 10
            End If
        End If
    Else
        If (Dias_Correspondientes - Dias_Pagados) = 0 Then
            Dias_Pago = 0
        Else
            If (Dias_Correspondientes - Dias_Pagados) > 0 Then
                If (Dias_Gozados + Dias_Lic) > Dias_Pagados And (Dias_Gozados + Dias_Lic) > 10 Then
                    Dias_Pago = Dias_Correspondientes - Dias_Pagados
                Else
                    Dias_Pago = 0
                End If
            Else
                'como se logró esto???
                Dias_Pago = 0
                Flog.writeline "Dias_Correspondientes - Dias_Pagados es menor a 0"
            End If
        End If
    
    End If
Case Else '(< 20)
    'A la primer salida pago los días correspondientes totales.
    If Primer_Pago Then
        Dias_Pago = Dias_Correspondientes
        Dias_Dto = Dias_Lic
    Else
        Dias_Pago = 0
        Dias_Dto = Dias_Lic
    End If
End Select

'Descuento segun Lic
'Dias_Dto = Dias_Lic
'FGZ - 25/08/2010 - Descuento segun dias de Pago
'If Dias_Pago = 0 Then
'    Dias_Dto = Dias_Lic
'Else
'    Dias_Dto = Dias_Pago
'End If
Dias_Dto = Dias_Lic
'------------------------------

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha_Hasta) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Aux_Fecha_Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If


'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'Genero los pagos y/o descuentos
'----------------------------------------------------------------------------------------------------------
'PAGOS
If Genera_Pagos And Dias_Pago > 0 Then
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Aux_TipDiaPago = TipDiaPago
    Dias_Afecta = Dias_Pago

    Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
End If

'Descuentos ---------------------------------------------
If Genera_Descuentos And Dias_Dto > 0 Then
    Dias_Pagados = 0
    Ya_Pago = True
    
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Dto 'Dias_Pago
    Dias_Afecta = Dias_Dto  'Dias_Pago
    
    Aux_TipDiaPago = TipDiaPago
    Aux_TipDiaDescuento = TipDiaDescuento
    
    'FGZ - 18/11/2010 ------------------------------------------------------
    Call DiasHabilesLic(Ternro, tipoVac, Aux_Fecha_Desde, Aux_Fecha_Hasta, CHabiles, cNoHabiles, cFeriados)
    'Dias_Afecta = CHabiles
    'If Dias_Afecta > (DateDiff("d", Aux_Fecha_Desde, Aux_Fecha_Hasta) + 1) Then
    If Dias_Afecta > CHabiles Then
        Dias_Afecta = CHabiles
        Aux_Fecha_Hasta = DateAdd("d", Dias_Afecta - 1, Aux_Fecha_Desde)
        Mes_Fin = Month(Aux_Fecha_Hasta)
        Ano_Fin = Year(Aux_Fecha_Hasta)
    End If
    
    If Jornal Then
        If Day(Aux_Fecha_Desde) > 15 Then
            Aux_TipDiaDescuento = st_ModeloDto2
            Quincena_Inicio = 2
        Else
            Aux_TipDiaDescuento = st_ModeloDto
            Quincena_Inicio = 1
        End If
        If Day(Aux_Fecha_Hasta) > 15 Then
            Quincena_Fin = 2
        Else
            Quincena_Fin = 1
        End If
        Flog.writeline "Quincena Inicio " & Quincena_Inicio
        Flog.writeline "Quincena Fin " & Quincena_Fin
        Flog.writeline "Modelo Pago " & Aux_TipDiaPago
        Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
    Else
        Quincena_Inicio = 1
    End If
    Quincena_Siguiente = Quincena_Inicio

    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If

    'EAM- Preguntar
    If (Aux_Fecha_Hasta <= Ultimo_Mes) Then 'termina en el mes
        'Dias_Afecta = DateDiff("d", Aux_Fecha_Desde, Aux_Fecha_Hasta) + 1
        Call DiasHabilesLic(Ternro, tipoVac, Aux_Fecha_Desde, Aux_Fecha_Hasta, CHabiles, cNoHabiles, cFeriados)
        Dias_Afecta = CHabiles
    Else 'continua en el mes siguiente
        'Dias_Afecta = DateDiff("d", Aux_Fecha_Desde, Ultimo_Mes) + 1
        Call DiasHabilesLic(Ternro, tipoVac, Aux_Fecha_Desde, Ultimo_Mes, CHabiles, cNoHabiles, cFeriados)
        Dias_Afecta = CHabiles
    End If
   
    Primero_Mes = AFecha(Mes_Afecta, Day(Aux_Fecha_Desde), Ano_Afecta)
    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    Aux_TipDiaPago = TipDiaPago
    Aux_TipDiaDescuento = TipDiaDescuento
    
    If Jornal Then 'reviso a que quincena corresponde
        If Day(Primero_Mes) > 15 Then
            Aux_TipDiaDescuento = st_ModeloDto2
            Quincena_Siguiente = 2
        Else
            Aux_TipDiaDescuento = st_ModeloDto
            Quincena_Siguiente = 1
        End If
    End If
    
    
    Do While Not (Fin_Licencia) Or Dias_Acumulados > 0
        If Jornal Then
            If Dias_Afecta > 15 Then
                Dias_Afecta = 15
                Dias_Acumulados = 1
            End If
            If Dias_Acumulados > 0 And Dias_Afecta < 15 Then
                If (Dias_Acumulados + Dias_Afecta) <= 15 Then
                    Dias_Afecta = Dias_Afecta + Dias_Acumulados
                    Dias_Acumulados = 0
                End If
            End If
        Else
            If Dias_Afecta > 30 Then
                Dias_Afecta = 30
                Dias_Acumulados = 1
            End If
            If Dias_Acumulados > 0 And Dias_Afecta < 30 Then
                If (Dias_Acumulados + Dias_Afecta) <= 30 Then
                    Dias_Afecta = Dias_Afecta + Dias_Acumulados
                    Dias_Acumulados = 0
                End If
            End If
            Aux_TipDiaDescuento = 7
        End If
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        'Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
        
        If Jornal Then
            If Quincena_Siguiente = 1 Then
                Quincena_Siguiente = 2
            Else
                Quincena_Siguiente = 1
            End If
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                Fin_Licencia = True
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        Fin_Licencia = True
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                Fin_Licencia = True
            End If
        End If
    
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
        Call DiasHabilesLic(Ternro, tipoVac, Primero_Mes, Ultimo_Mes, CHabiles, cNoHabiles, cFeriados)
        If (Aux_Fecha_Hasta <= Ultimo_Mes) Then
            'If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
            If (Dias_Restantes < CHabiles) Then
                Dias_Afecta = Dias_Restantes
            Else
                'Dias_Afecta = DateDiff("d", Primero_Mes, rs!Aux_Fecha_Hasta) + 1
                'Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Fecha_Hasta) + 1
                Call DiasHabilesLic(Ternro, tipoVac, Primero_Mes, Aux_Fecha_Hasta, CHabiles, cNoHabiles, cFeriados)
                Dias_Afecta = CHabiles
            End If
        Else
            'Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
            Call DiasHabilesLic(Ternro, tipoVac, Primero_Mes, Ultimo_Mes, CHabiles, cNoHabiles, cFeriados)
            Dias_Afecta = CHabiles
        End If
        
        If Not Jornal Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        Else
            If Dias_Afecta > 15 Then
               Dias_Afecta = 15
            End If
        End If
        
        If Jornal Then
            If Day(Primero_Mes) > 15 Then
                Aux_TipDiaDescuento = st_ModeloDto2
            Else
                Aux_TipDiaDescuento = st_ModeloDto
            End If
        End If
    Loop
End If
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------

''Descuentos
'If Genera_Descuentos And Dias_Dto > 0 Then
'    Ya_Pago = True
'
'    Fin_Licencia = False
'    Mes_Afecta = Mes_Inicio
'    Ano_Afecta = Ano_Inicio
'    Dias_Pendientes = 0
'    Dias_Restantes = Dias_Dto
'
'    Aux_TipDiaPago = TipDiaPago
'    Aux_TipDiaDescuento = TipDiaDescuento
'
'    Dias_Afecta = Dias_Dto
'    If Dias_Afecta > (DateDiff("d", Aux_Fecha_Desde, Aux_Fecha_Hasta) + 1) Then
'        Aux_Fecha_Hasta = DateAdd("d", Dias_Afecta - 1, Aux_Fecha_Desde)
'        Mes_Fin = Month(Aux_Fecha_Hasta)
'        Ano_Fin = Year(Aux_Fecha_Hasta)
'    End If
'
'
'    If Jornal Then
'        If Day(Aux_Fecha_Desde) > 15 Then
'            Aux_TipDiaDescuento = st_ModeloDto2
'            Quincena_Inicio = 2
'        Else
'            Aux_TipDiaDescuento = st_ModeloDto
'            Quincena_Inicio = 1
'        End If
'        If Day(Aux_Fecha_Hasta) > 15 Then
'            Quincena_Fin = 2
'        Else
'            Quincena_Fin = 1
'        End If
'        Flog.writeline "Quincena Inicio " & Quincena_Inicio
'        Flog.writeline "Quincena Fin " & Quincena_Fin
'        Flog.writeline "Modelo Pago " & Aux_TipDiaPago
'        Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
'    Else
'        Quincena_Inicio = 1
'    End If
'    Quincena_Siguiente = Quincena_Inicio
'
'    If Mes_Afecta = 2 Then
'        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'        If Not Jornal Then
'            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'        Else
'            If Quincena_Siguiente = 1 Then
'                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
'            Else
'                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'            End If
'        End If
'    Else
'        If Not Jornal Then
'            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'            If Mes_Afecta = 12 Then
'                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
'            Else
'                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'            End If
'        Else
'            If Quincena_Siguiente = 1 Then
'                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
'            Else
'                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
'                If Mes_Afecta = 12 Then
'                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
'                Else
'                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'                End If
'            End If
'        End If
'    End If
'
'    Dias_Afecta = Dias_Dto
'    If (Aux_Fecha_Hasta <= Ultimo_Mes) Then 'termina en el mes
'        Dias_Afecta = DateDiff("d", Aux_Fecha_Desde, Aux_Fecha_Hasta) + 1
'    Else 'continua en el mes siguiente
'        Dias_Afecta = DateDiff("d", Aux_Fecha_Desde, Ultimo_Mes) + 1
'    End If
'
'    Primero_Mes = AFecha(Mes_Afecta, Day(Aux_Fecha_Desde), Ano_Afecta)
'    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'    Aux_TipDiaPago = TipDiaPago
'    Aux_TipDiaDescuento = TipDiaDescuento
'
'    If Jornal Then 'reviso a que quincena corresponde
'        If Day(Primero_Mes) > 15 Then
'            Aux_TipDiaDescuento = st_ModeloDto2
'            Quincena_Siguiente = 2
'        Else
'            Aux_TipDiaDescuento = st_ModeloDto
'            Quincena_Siguiente = 1
'        End If
'    End If
'
'
'    Do While Not (Fin_Licencia) Or Dias_Acumulados > 0
'        If Jornal Then
'            If Dias_Afecta > 15 Then
'                Dias_Afecta = 15
'                Dias_Acumulados = 1
'            End If
'            If Dias_Acumulados > 0 And Dias_Afecta < 15 Then
'                If (Dias_Acumulados + Dias_Afecta) <= 15 Then
'                    Dias_Afecta = Dias_Afecta + Dias_Acumulados
'                    Dias_Acumulados = 0
'                End If
'            End If
'        Else
'            If Dias_Afecta > 30 Then
'                Dias_Afecta = 30
'                Dias_Acumulados = 1
'            End If
'            If Dias_Acumulados > 0 And Dias_Afecta < 30 Then
'                If (Dias_Acumulados + Dias_Afecta) <= 30 Then
'                    Dias_Afecta = Dias_Afecta + Dias_Acumulados
'                    Dias_Acumulados = 0
'                End If
'            End If
'            Aux_TipDiaDescuento = 7
'        End If
'
'        Dias_Restantes = Dias_Restantes - Dias_Afecta
'
'        If Dias_Afecta <> 0 Then
'            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
'        End If
'
'        If Jornal Then
'            If Quincena_Siguiente = 1 Then
'                Quincena_Siguiente = 2
'            Else
'                Quincena_Siguiente = 1
'            End If
'        End If
'
'        'determinar a que continua en el proximo mes
'        If (Mes_Afecta = 12) Then
'            If Not Jornal Then
'                Mes_Afecta = 1
'                Ano_Afecta = Ano_Afecta + 1
'            Else
'                If Quincena_Siguiente = 1 Then
'                    Mes_Afecta = 1
'                    Ano_Afecta = Ano_Afecta + 1
'                Else
'                    'Queda como esta, en el mismo mes y año
'                End If
'            End If
'        Else
'            If Not Jornal Then
'                Mes_Afecta = Mes_Afecta + 1
'            Else
'                If Quincena_Siguiente = 1 Then
'                    Mes_Afecta = Mes_Afecta + 1
'                Else
'                    'Queda como esta, en el mismo mes y año
'                End If
'            End If
'        End If
'
'        If (Ano_Afecta = Ano_Fin) Then
'            If (Mes_Afecta > Mes_Fin) Then
'                Fin_Licencia = True
'            Else
'                If (Mes_Afecta = Mes_Fin) And Jornal Then
'                    If Quincena_Siguiente > Quincena_Fin Then
'                        Fin_Licencia = True
'                    End If
'                End If
'            End If
'        Else
'            If (Ano_Afecta > Ano_Fin) Then
'                Fin_Licencia = True
'            End If
'        End If
'
'        If Mes_Afecta = 2 Then
'            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'            If Not Jornal Then
'                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'            Else
'                If Quincena_Siguiente = 1 Then
'                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
'                Else
'                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'                End If
'            End If
'        Else
'            If Not Jornal Then
'                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'                If Mes_Afecta = 12 Then
'                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
'                Else
'                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'                End If
'            Else
'                If Quincena_Siguiente = 1 Then
'                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
'                Else
'                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
'                    If Mes_Afecta = 12 Then
'                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
'                    Else
'                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'                    End If
'                End If
'            End If
'        End If
'
'        If (Aux_Fecha_Hasta <= Ultimo_Mes) Then
'            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
'                Dias_Afecta = Dias_Restantes
'            Else
'                Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Fecha_Hasta) + 1
'            End If
'        Else
'            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
'        End If
'
'        If Not Jornal Then
'            If Dias_Afecta > 30 Then
'               Dias_Afecta = 30
'            End If
'        Else
'            If Dias_Afecta > 15 Then
'               Dias_Afecta = 15
'            End If
'        End If
'
'        If Jornal Then
'            If Day(Primero_Mes) > 15 Then
'                Aux_TipDiaDescuento = st_ModeloDto2
'            Else
'                Aux_TipDiaDescuento = st_ModeloDto
'            End If
'        End If
'    Loop
'End If

rs.Close
Set rs = Nothing
End Sub



Public Sub Politica1500_Monresa()
'-----------------------------------------------------------------------
'Descripcion: Version para Monresa (Uruguay)
' Autor     : FGZ
' Fecha     : 11/08/2010
'-----------------------------------------------------------------------
'Descripcion
'   Para el PAGO Se manejan 3 situaciones
'       A) Dias correspondientes = 20
'       B) Dias correspondientes < 20
'       C) Dias correspondientes > 20
'   Para el descuento se genera por la misma cantidad que el PAGO con la diferencia que se desglosan
'       cuando pasan de un mes a otro. A diferencia del pago que se genera en un solo mes
'   ----------------------------------------------------------------------------

    'PAGOS ---------------------------------------------------------------------
    'A) Dias correspondientes = 20
    ' PAGO = FIX( (SUM(Licencias del periodo) / 10)*10) - Dias Pagados.
    '---------------------------------------------------------------------------
    'B) Dias Correspondientes > 20
    '   SI SUM(Pagos del periodo) = 0 THEN
    '       SI SUM(Licencias del Periodo) = Dias Correspondientes THEN
    '           PAGO = Dias Correspondientes
    '       ELSE
    '           SI SUM(Licencias del Periodo) >= 10 THEN
    '               PAGO = MINIMO(SUM(Licencias del Periodo), Dias Correspondientes - 10)
    '           ELSE
    '               PAGO = 0
    '           END
    '       END
    '   ELSE
    '       SI SUM(Licencias del Periodo) = Dias Correspondientes THEN
    '           PAGO = Dias Correspondientes - SUM(Pagos del periodo)
    '       ELSE
    '           PAGO = 0
    '       END
    '   END
    '----------------------------------------------------------------------------
    'C) Dias Correspondientes < 20
    '   PAGO = Dias Correspondientes - SUM(Pagos del Periodo).
    'PAGOS ---------------------------------------------------------------------
    
' ---------------------------------------------------------------------------------------
Const Dias_Tope = 10

Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Dias_Afecta As Integer
Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer
Dim Jornal As Boolean
Dim Existe_Pago As Boolean
Dim Primer_Pago As Boolean

Dim Dias_Correspondientes As Integer

Dim Dias_Gozados As Integer
Dim Dias_Pagados As Integer
Dim Dias_Acumulados As Integer

Dim Dias_Lic  As Integer
Dim Dias_Pago As Integer
Dim Dias_Dto  As Integer

Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Aux_Fecha_Desde As Date
Dim Aux_Fecha_Hasta As Date
Dim New_Fecha_Hasta As Date


Dim tipoVac As Long
Dim CHabiles As Integer
Dim cNoHabiles As Integer
Dim cFeriados As Integer

Dim rs As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset
'===================================================================

'Busco la licencia dentro del intervalo especificado
Dias_Lic = 0
StrSql = "SELECT * FROM emp_lic "
StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro"
StrSql = StrSql & " WHERE emp_lic.emp_licnro = " & nrolicencia
OpenRecordset StrSql, rs
If Not rs.EOF Then
    NroVac = rs!vacnro
    'Dias_Lic = rs!elcantdiashab
    Dias_Lic = rs!elcantdias
    Aux_Fecha_Desde = rs!elfechadesde
    Aux_Fecha_Hasta = rs!elfechahasta
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If



Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 23
    TipDiaDescuento = 3
End If
If st_ModeloDto2 = 0 Then
    st_ModeloDto2 = st_ModeloDto
End If
If st_ModeloPago2 = 0 Then
    st_ModeloPago2 = st_ModeloPago
End If

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'Reproceso
If Reproceso Then
    If Not Ya_Pago Then
        'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
        StrSql = "SELECT * FROM vacpagdesc "
        StrSql = StrSql & " WHERE ternro = " & Ternro
        StrSql = StrSql & " AND vacpagdesc.vacnro = " & NroVac
        StrSql = StrSql & " AND vacpagdesc.pronro is not null"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        Else
            StrSql = "DELETE FROM vacpagdesc"
            StrSql = StrSql & " WHERE ternro = " & Ternro
            StrSql = StrSql & " AND vacpagdesc.vacnro = " & NroVac
            If Genera_Pagos And Genera_Descuentos Then
                StrSql = StrSql & " AND (pago_dto = 3" 'pagos
                StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
            Else
                If Genera_Pagos Then
                    StrSql = StrSql & " AND pago_dto = 3" 'pagos
                End If
                If Genera_Descuentos Then
                    StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
                End If
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
Else
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago/descuento, salir"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
End If


'Busco los dias correspondientes
StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs
If rs.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
Else
    Dias_Correspondientes = rs!vdiascorcant
    'FGZ - 18/11/2010 -------------------------
    tipoVac = rs!tipvacnro
    'FGZ - 18/11/2010 -------------------------
End If


''Determino si es el primer pago del periodo
'Existe_Pago = False
'StrSql = "SELECT * FROM vacpagdesc "
'StrSql = StrSql & " WHERE ternro = " & Ternro
'StrSql = StrSql & " AND vacpagdesc.vacnro = " & NroVac
'OpenRecordset StrSql, rs
'If Not rs.EOF Then
'    'Existe_Pago = True
'    Primer_Pago = False
'Else
'    'Existe_Pago = False
'    Primer_Pago = True
'End If
''Primer_Pago = Not Ya_Pago Or Existe_Pago

'Busco la cantidad de licencias ya gozadas
Dias_Gozados = 0
StrSql = "SELECT sum(elcantdias) dias FROM emp_lic "
StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro"
StrSql = StrSql & " WHERE empleado = " & Ternro
StrSql = StrSql & " AND emp_lic.emp_licnro <> " & nrolicencia
StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(Aux_Fecha_Desde)
'FGZ - 18/11/2010 ------
StrSql = StrSql & " AND lic_vacacion.vacnro = " & NroVac
'FGZ - 18/11/2010 ------
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Dias_Gozados = IIf(EsNulo(rs!dias), 0, rs!dias)
End If

'Busco los dias ya pagados
Dias_Pagados = 0
StrSql = "SELECT sum(cantdias) dias FROM vacpagdesc "
StrSql = StrSql & " WHERE ternro = " & Ternro
StrSql = StrSql & " AND vacnro = " & NroVac
StrSql = StrSql & " AND pago_dto = 3" 'pagos
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Dias_Pagados = IIf(EsNulo(rs!dias), 0, rs!dias)
End If


'Pago segun..
Select Case Dias_Correspondientes
Case Is = 20
    Dias_Pago = Fix((Dias_Gozados + Dias_Lic) / 10) * 10 - Dias_Pagados

Case Is > 20
    If Dias_Pagados = 0 Then
        If (Dias_Gozados + Dias_Lic) = Dias_Correspondientes Then
            Dias_Pago = Dias_Correspondientes
        Else
            If (Dias_Gozados + Dias_Lic) >= Dias_Tope Then
                Dias_Pago = Minimo(Dias_Gozados + Dias_Lic, Dias_Correspondientes - Dias_Tope)
            Else
                Dias_Pago = 0
            End If
        End If
    Else
        If (Dias_Gozados + Dias_Lic) = Dias_Correspondientes Then
            Dias_Pago = Dias_Correspondientes - Dias_Pagados
        Else
            Dias_Pago = 0
        End If
    End If
Case Else '(< 20)
    Dias_Pago = Dias_Correspondientes - Dias_Pagados
End Select

Dias_Dto = Dias_Pago
'------------------------------

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha_Hasta) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Aux_Fecha_Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If


'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'Genero los pagos y/o descuentos
'----------------------------------------------------------------------------------------------------------
'PAGOS
If Genera_Pagos And Dias_Pago > 0 Then
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Aux_TipDiaPago = TipDiaPago
    Dias_Afecta = Dias_Pago

    Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
End If

'Descuentos ---------------------------------------------
If Genera_Descuentos And Dias_Dto > 0 Then
    Dias_Pagados = 0
    Ya_Pago = True
    
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Dto 'Dias_Pago
    Dias_Afecta = Dias_Dto  'Dias_Pago
    
    Aux_TipDiaPago = TipDiaPago
    Aux_TipDiaDescuento = TipDiaDescuento
    
    'FGZ - 30/11/2010 ---------------------
    ' Redefino la fecha hasta tal que  la cantidad de dias sean habiles
    New_Fecha_Hasta = CalcFechaHastaLic(Ternro, tipoVac, Aux_Fecha_Desde, Dias_Dto)
    Aux_Fecha_Hasta = New_Fecha_Hasta
    'FGZ - 30/11/2010 ---------------------
    
    
    'FGZ - 18/11/2010 ------------------------------------------------------
    Call DiasHabilesLic(Ternro, tipoVac, Aux_Fecha_Desde, Aux_Fecha_Hasta, CHabiles, cNoHabiles, cFeriados)
    'Dias_Afecta = CHabiles
    'If Dias_Afecta > (DateDiff("d", Aux_Fecha_Desde, Aux_Fecha_Hasta) + 1) Then
    If Dias_Afecta > CHabiles Then
        If CHabiles < Dias_Tope Then
            Dias_Afecta = CHabiles
            Aux_Fecha_Hasta = DateAdd("d", Dias_Afecta - 1, Aux_Fecha_Desde)
            Mes_Fin = Month(Aux_Fecha_Hasta)
            Ano_Fin = Year(Aux_Fecha_Hasta)
        Else
            Dias_Afecta = CHabiles
            Aux_Fecha_Hasta = DateAdd("d", Dias_Afecta - 1, Aux_Fecha_Desde)
            Mes_Fin = Month(Aux_Fecha_Hasta)
            Ano_Fin = Year(Aux_Fecha_Hasta)
        End If
    End If
    
    If Jornal Then
        If Day(Aux_Fecha_Desde) > 15 Then
            Aux_TipDiaDescuento = st_ModeloDto2
            Quincena_Inicio = 2
        Else
            Aux_TipDiaDescuento = st_ModeloDto
            Quincena_Inicio = 1
        End If
        If Day(Aux_Fecha_Hasta) > 15 Then
            Quincena_Fin = 2
        Else
            Quincena_Fin = 1
        End If
        Flog.writeline "Quincena Inicio " & Quincena_Inicio
        Flog.writeline "Quincena Fin " & Quincena_Fin
        Flog.writeline "Modelo Pago " & Aux_TipDiaPago
        Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
    Else
        Quincena_Inicio = 1
    End If
    Quincena_Siguiente = Quincena_Inicio

    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If

    'EAM- Preguntar
    If (Aux_Fecha_Hasta <= Ultimo_Mes) Then 'termina en el mes
        'Dias_Afecta = DateDiff("d", Aux_Fecha_Desde, Aux_Fecha_Hasta) + 1
        Call DiasHabilesLic(Ternro, tipoVac, Aux_Fecha_Desde, Aux_Fecha_Hasta, CHabiles, cNoHabiles, cFeriados)
        Dias_Afecta = CHabiles
    Else 'continua en el mes siguiente
        'Dias_Afecta = DateDiff("d", Aux_Fecha_Desde, Ultimo_Mes) + 1
        Call DiasHabilesLic(Ternro, tipoVac, Aux_Fecha_Desde, Ultimo_Mes, CHabiles, cNoHabiles, cFeriados)
        Dias_Afecta = CHabiles
    End If
   
    Primero_Mes = AFecha(Mes_Afecta, Day(Aux_Fecha_Desde), Ano_Afecta)
    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
    Aux_TipDiaPago = TipDiaPago
    Aux_TipDiaDescuento = TipDiaDescuento
    
    If Jornal Then 'reviso a que quincena corresponde
        If Day(Primero_Mes) > 15 Then
            Aux_TipDiaDescuento = st_ModeloDto2
            Quincena_Siguiente = 2
        Else
            Aux_TipDiaDescuento = st_ModeloDto
            Quincena_Siguiente = 1
        End If
    End If
    
    
    Do While Not (Fin_Licencia) Or Dias_Acumulados > 0
        If Jornal Then
            If Dias_Afecta > 15 Then
                Dias_Afecta = 15
                Dias_Acumulados = 1
            End If
            If Dias_Acumulados > 0 And Dias_Afecta < 15 Then
                If (Dias_Acumulados + Dias_Afecta) <= 15 Then
                    Dias_Afecta = Dias_Afecta + Dias_Acumulados
                    Dias_Acumulados = 0
                End If
            End If
        Else
            If Dias_Afecta > 30 Then
                Dias_Afecta = 30
                Dias_Acumulados = 1
            End If
            If Dias_Acumulados > 0 And Dias_Afecta < 30 Then
                If (Dias_Acumulados + Dias_Afecta) <= 30 Then
                    Dias_Afecta = Dias_Afecta + Dias_Acumulados
                    Dias_Acumulados = 0
                End If
            End If
            'Aux_TipDiaDescuento = 7
        End If
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        'Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
        
        If Jornal Then
            If Quincena_Siguiente = 1 Then
                Quincena_Siguiente = 2
            Else
                Quincena_Siguiente = 1
            End If
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                If Dias_Restantes <= 0 Then
                    Fin_Licencia = True
                End If
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        If Dias_Restantes <= 0 Then
                            Fin_Licencia = True
                        End If
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                If Dias_Restantes <= 0 Then
                    Fin_Licencia = True
                End If
            End If
        End If
    
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
        Call DiasHabilesLic(Ternro, tipoVac, Primero_Mes, Ultimo_Mes, CHabiles, cNoHabiles, cFeriados)
        If (Aux_Fecha_Hasta <= Ultimo_Mes) Then
            'If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
            If (Dias_Restantes < CHabiles) Then
                Dias_Afecta = Dias_Restantes
            Else
                'Dias_Afecta = DateDiff("d", Primero_Mes, rs!Aux_Fecha_Hasta) + 1
                'Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Fecha_Hasta) + 1
                Call DiasHabilesLic(Ternro, tipoVac, Primero_Mes, Aux_Fecha_Hasta, CHabiles, cNoHabiles, cFeriados)
                Dias_Afecta = CHabiles
            End If
        Else
            'Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
            Call DiasHabilesLic(Ternro, tipoVac, Primero_Mes, Ultimo_Mes, CHabiles, cNoHabiles, cFeriados)
            Dias_Afecta = CHabiles
        End If
        
        If Not Jornal Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        Else
            If Dias_Afecta > 15 Then
               Dias_Afecta = 15
            End If
        End If
        
        If Jornal Then
            If Day(Primero_Mes) > 15 Then
                Aux_TipDiaDescuento = st_ModeloDto2
            Else
                Aux_TipDiaDescuento = st_ModeloDto
            End If
        End If
    Loop
End If
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------

''Descuentos
'If Genera_Descuentos And Dias_Dto > 0 Then
'    Ya_Pago = True
'
'    Fin_Licencia = False
'    Mes_Afecta = Mes_Inicio
'    Ano_Afecta = Ano_Inicio
'    Dias_Pendientes = 0
'    Dias_Restantes = Dias_Dto
'
'    Aux_TipDiaPago = TipDiaPago
'    Aux_TipDiaDescuento = TipDiaDescuento
'
'    Dias_Afecta = Dias_Dto
'    If Dias_Afecta > (DateDiff("d", Aux_Fecha_Desde, Aux_Fecha_Hasta) + 1) Then
'        Aux_Fecha_Hasta = DateAdd("d", Dias_Afecta - 1, Aux_Fecha_Desde)
'        Mes_Fin = Month(Aux_Fecha_Hasta)
'        Ano_Fin = Year(Aux_Fecha_Hasta)
'    End If
'
'
'    If Jornal Then
'        If Day(Aux_Fecha_Desde) > 15 Then
'            Aux_TipDiaDescuento = st_ModeloDto2
'            Quincena_Inicio = 2
'        Else
'            Aux_TipDiaDescuento = st_ModeloDto
'            Quincena_Inicio = 1
'        End If
'        If Day(Aux_Fecha_Hasta) > 15 Then
'            Quincena_Fin = 2
'        Else
'            Quincena_Fin = 1
'        End If
'        Flog.writeline "Quincena Inicio " & Quincena_Inicio
'        Flog.writeline "Quincena Fin " & Quincena_Fin
'        Flog.writeline "Modelo Pago " & Aux_TipDiaPago
'        Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
'    Else
'        Quincena_Inicio = 1
'    End If
'    Quincena_Siguiente = Quincena_Inicio
'
'    If Mes_Afecta = 2 Then
'        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'        If Not Jornal Then
'            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'        Else
'            If Quincena_Siguiente = 1 Then
'                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
'            Else
'                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'            End If
'        End If
'    Else
'        If Not Jornal Then
'            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'            If Mes_Afecta = 12 Then
'                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
'            Else
'                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'            End If
'        Else
'            If Quincena_Siguiente = 1 Then
'                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
'            Else
'                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
'                If Mes_Afecta = 12 Then
'                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
'                Else
'                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'                End If
'            End If
'        End If
'    End If
'
'    Dias_Afecta = Dias_Dto
'    If (Aux_Fecha_Hasta <= Ultimo_Mes) Then 'termina en el mes
'        Dias_Afecta = DateDiff("d", Aux_Fecha_Desde, Aux_Fecha_Hasta) + 1
'    Else 'continua en el mes siguiente
'        Dias_Afecta = DateDiff("d", Aux_Fecha_Desde, Ultimo_Mes) + 1
'    End If
'
'    Primero_Mes = AFecha(Mes_Afecta, Day(Aux_Fecha_Desde), Ano_Afecta)
'    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'    Aux_TipDiaPago = TipDiaPago
'    Aux_TipDiaDescuento = TipDiaDescuento
'
'    If Jornal Then 'reviso a que quincena corresponde
'        If Day(Primero_Mes) > 15 Then
'            Aux_TipDiaDescuento = st_ModeloDto2
'            Quincena_Siguiente = 2
'        Else
'            Aux_TipDiaDescuento = st_ModeloDto
'            Quincena_Siguiente = 1
'        End If
'    End If
'
'
'    Do While Not (Fin_Licencia) Or Dias_Acumulados > 0
'        If Jornal Then
'            If Dias_Afecta > 15 Then
'                Dias_Afecta = 15
'                Dias_Acumulados = 1
'            End If
'            If Dias_Acumulados > 0 And Dias_Afecta < 15 Then
'                If (Dias_Acumulados + Dias_Afecta) <= 15 Then
'                    Dias_Afecta = Dias_Afecta + Dias_Acumulados
'                    Dias_Acumulados = 0
'                End If
'            End If
'        Else
'            If Dias_Afecta > 30 Then
'                Dias_Afecta = 30
'                Dias_Acumulados = 1
'            End If
'            If Dias_Acumulados > 0 And Dias_Afecta < 30 Then
'                If (Dias_Acumulados + Dias_Afecta) <= 30 Then
'                    Dias_Afecta = Dias_Afecta + Dias_Acumulados
'                    Dias_Acumulados = 0
'                End If
'            End If
'            Aux_TipDiaDescuento = 7
'        End If
'
'        Dias_Restantes = Dias_Restantes - Dias_Afecta
'
'        If Dias_Afecta <> 0 Then
'            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
'        End If
'
'        If Jornal Then
'            If Quincena_Siguiente = 1 Then
'                Quincena_Siguiente = 2
'            Else
'                Quincena_Siguiente = 1
'            End If
'        End If
'
'        'determinar a que continua en el proximo mes
'        If (Mes_Afecta = 12) Then
'            If Not Jornal Then
'                Mes_Afecta = 1
'                Ano_Afecta = Ano_Afecta + 1
'            Else
'                If Quincena_Siguiente = 1 Then
'                    Mes_Afecta = 1
'                    Ano_Afecta = Ano_Afecta + 1
'                Else
'                    'Queda como esta, en el mismo mes y año
'                End If
'            End If
'        Else
'            If Not Jornal Then
'                Mes_Afecta = Mes_Afecta + 1
'            Else
'                If Quincena_Siguiente = 1 Then
'                    Mes_Afecta = Mes_Afecta + 1
'                Else
'                    'Queda como esta, en el mismo mes y año
'                End If
'            End If
'        End If
'
'        If (Ano_Afecta = Ano_Fin) Then
'            If (Mes_Afecta > Mes_Fin) Then
'                Fin_Licencia = True
'            Else
'                If (Mes_Afecta = Mes_Fin) And Jornal Then
'                    If Quincena_Siguiente > Quincena_Fin Then
'                        Fin_Licencia = True
'                    End If
'                End If
'            End If
'        Else
'            If (Ano_Afecta > Ano_Fin) Then
'                Fin_Licencia = True
'            End If
'        End If
'
'        If Mes_Afecta = 2 Then
'            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'            If Not Jornal Then
'                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'            Else
'                If Quincena_Siguiente = 1 Then
'                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
'                Else
'                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'                End If
'            End If
'        Else
'            If Not Jornal Then
'                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'                If Mes_Afecta = 12 Then
'                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
'                Else
'                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'                End If
'            Else
'                If Quincena_Siguiente = 1 Then
'                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
'                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
'                Else
'                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
'                    If Mes_Afecta = 12 Then
'                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
'                    Else
'                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
'                    End If
'                End If
'            End If
'        End If
'
'        If (Aux_Fecha_Hasta <= Ultimo_Mes) Then
'            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
'                Dias_Afecta = Dias_Restantes
'            Else
'                Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Fecha_Hasta) + 1
'            End If
'        Else
'            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
'        End If
'
'        If Not Jornal Then
'            If Dias_Afecta > 30 Then
'               Dias_Afecta = 30
'            End If
'        Else
'            If Dias_Afecta > 15 Then
'               Dias_Afecta = 15
'            End If
'        End If
'
'        If Jornal Then
'            If Day(Primero_Mes) > 15 Then
'                Aux_TipDiaDescuento = st_ModeloDto2
'            Else
'                Aux_TipDiaDescuento = st_ModeloDto
'            End If
'        End If
'    Loop
'End If

rs.Close
Set rs = Nothing
End Sub

Public Sub Politica1500_PagoLic_PT()
'-----------------------------------------------------------------------
'Descripcion: Version para PORTUGAL
' Autor     : Gonzalez Nicolás
' Fecha     : 15/05/2012
'-----------------------------------------------------------------------
'Descripcion: Calcula el pago sobre licencias tomadas.
'            Se toma 1 día y se calcula la totalidad de días correspondientes, como máximo 22
' ---------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------
Const Dias_Tope = 22

Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Dias_Afecta As Integer
Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer
Dim Jornal As Boolean
Dim Existe_Pago As Boolean
Dim Primer_Pago As Boolean

Dim Dias_Correspondientes As Integer

Dim Dias_Gozados As Integer
Dim Dias_Pagados As Integer
Dim Dias_Acumulados As Integer

Dim Dias_Lic  As Integer
Dim Dias_Pago As Integer
Dim Dias_Dto  As Integer

Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Aux_Fecha_Desde As Date
Dim Aux_Fecha_Hasta As Date
Dim New_Fecha_Hasta As Date

Dim tipoVac As Long
Dim CHabiles As Integer
Dim cNoHabiles As Integer
Dim cFeriados As Integer

Dim rs As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset
'===================================================================

'====================================================
'Busco la licencia dentro del intervalo especificado
'====================================================
Dias_Lic = 0
StrSql = "SELECT * FROM emp_lic "
StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro"
StrSql = StrSql & " WHERE emp_lic.emp_licnro = " & nrolicencia
OpenRecordset StrSql, rs
If Not rs.EOF Then
    NroVac = rs!vacnro
    'Dias_Lic = rs!elcantdiashab
    Dias_Lic = rs!elcantdias
    Aux_Fecha_Desde = rs!elfechadesde
    Aux_Fecha_Hasta = rs!elfechahasta
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If

Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 23
    TipDiaDescuento = 3
End If

If st_ModeloPago2 = 0 Then
    st_ModeloPago2 = st_ModeloPago
End If

Genera_Pagos = True
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos."
    'Genera_Pagos = True
End If

'===========
'REPROCESO
'===========
If Reproceso Then
    If Not Ya_Pago Then
        'verificar si no tiene pagos ya procesados, sino depurarlos
        StrSql = "SELECT * FROM vacpagdesc "
        StrSql = StrSql & " WHERE ternro = " & Ternro
        StrSql = StrSql & " AND vacpagdesc.vacnro = " & NroVac
        StrSql = StrSql & " AND vacpagdesc.pronro is not null"
        OpenRecordset StrSql, rs
        If Not rs.EOF Then
            Flog.writeline "No se puede Reprocesar, hay pagos Liquidados. "
        Else
            StrSql = "DELETE FROM vacpagdesc"
            StrSql = StrSql & " WHERE ternro = " & Ternro
            StrSql = StrSql & " AND vacpagdesc.vacnro = " & NroVac
            If Genera_Pagos And Genera_Descuentos Then
                StrSql = StrSql & " AND (pago_dto = 3)" 'pagos
            Else
                If Genera_Pagos Then
                    StrSql = StrSql & " AND pago_dto = 3" 'pagos
                End If
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
Else
    '________________________________________________________________
    'Si no es reproceso y existe el desglose de pago, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    If Not rs.EOF Then
        Flog.writeline "Reproceso y existe el desglose de pago, salir"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
End If

'================================
'Busco los dias correspondientes
'================================
StrSql = "SELECT * FROM vacdiascor "
StrSql = StrSql & " WHERE vacdiascor.Ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs
If rs.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
Else
    Dias_Correspondientes = rs!vdiascorcant
'    TipoVac = rs!tipvacnro
End If


'=============================================
'Busco la cantidad de licencias ya gozadas
'=============================================
Dias_Gozados = 0
StrSql = "SELECT sum(elcantdias) dias FROM emp_lic "
StrSql = StrSql & " INNER JOIN lic_vacacion ON emp_lic.emp_licnro = lic_vacacion.emp_licnro"
StrSql = StrSql & " WHERE empleado = " & Ternro
StrSql = StrSql & " AND emp_lic.emp_licnro <> " & nrolicencia
StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(Aux_Fecha_Desde)
StrSql = StrSql & " AND lic_vacacion.vacnro = " & NroVac
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Dias_Gozados = IIf(EsNulo(rs!dias), 0, rs!dias)
End If

'==========================
'Busco los dias ya pagados
'==========================
Dias_Pagados = 0
StrSql = "SELECT sum(cantdias) dias FROM vacpagdesc "
StrSql = StrSql & " WHERE ternro = " & Ternro
StrSql = StrSql & " AND vacnro = " & NroVac
StrSql = StrSql & " AND pago_dto = 3" 'pagos
OpenRecordset StrSql, rs
If Not rs.EOF Then
    Dias_Pagados = IIf(EsNulo(rs!dias), 0, rs!dias)
End If

If Dias_Pagados > 0 Then
    Flog.writeline "Ya se ha generado el pago para este período"
    rs.Close
    Set rs = Nothing
    Exit Sub
End If

'==========================
'CALCULO DIAS DE PAGO
'Pago segun..
'==========================
Flog.writeline ""
Flog.writeline "Dias_Pagados: " & Dias_Pagados
Flog.writeline "Dias_Gozados: " & Dias_Gozados
Flog.writeline "Dias_Lic: " & Dias_Lic
Flog.writeline "Dias_Correspondientes: " & Dias_Correspondientes
Flog.writeline ""

Select Case Dias_Correspondientes
    Case Is <= 22
        Dias_Pago = Dias_Correspondientes
    Case Else
        Dias_Pago = Dias_Tope

'Case Is < 10
'    'Dias_Pago = Fix((Dias_Gozados + Dias_Lic) / 10) * 10 - Dias_Pagados
'    Dias_Pago = Dias_Correspondientes
'
'Case Is >= 10
'    If Dias_Pagados = 0 Then
'        If (Dias_Gozados + Dias_Lic) = Dias_Correspondientes Then
'            Dias_Pago = Dias_Correspondientes
'            Flog.writeline "Dias_Pago 1 :" & Dias_Pago
'        Else
'            If (Dias_Gozados + Dias_Lic) >= Dias_Tope Then
'                'VER SI SE DEJA Dias_Tope
'                Dias_Pago = Minimo(Dias_Gozados + Dias_Lic, Dias_Correspondientes - Dias_Tope)
'                Flog.writeline "Dias_Pago 2 :" & Dias_Pago
'            Else
'                Dias_Pago = 0
'                Flog.writeline "Dias_Pago 3 :" & Dias_Pago
'            End If
'        End If
'    Else
'        If (Dias_Gozados + Dias_Lic) = Dias_Correspondientes Then
'            Dias_Pago = Dias_Correspondientes - Dias_Pagados
'            Flog.writeline "Dias_Pago 4 :" & Dias_Pago
'        Else
'            Dias_Pago = 0
'            Flog.writeline "Dias_Pago 5 :" & Dias_Pago
'        End If
'    End If
'Case Else '(< 20)
'    Dias_Pago = Dias_Correspondientes - Dias_Pagados
End Select

'Dias_Dto = Dias_Pago


'=================================
'Busco la forma de liquidacion
'=================================
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha_Hasta) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Aux_Fecha_Hasta) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

'=================================
'Genero los pagos
'=================================
'PAGOS
If Genera_Pagos And Dias_Pago > 0 Then
    Flog.writeline "Genero Pago de Vacaciones"
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Aux_TipDiaPago = TipDiaPago
    Dias_Afecta = Dias_Pago

    Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
Else
    Flog.writeline "El Pago de Vacaciones no ha sido generado"

End If


rs.Close
Set rs = Nothing
End Sub

Public Sub Politica1500_PagoXdiasCorr_PT()
'-----------------------------------------------------------------------
'Descripcion: Genera a partir de dias correspondientes, con tope de 22
'Autor      : Gonzalez Nicolás
'Fecha      : 16/05/2012
'Ult Mod    :
'-----------------------------------------------------------------------
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim Dias_Afecta   As Integer
Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer

Dim Aux_TotalDiasCorrespondientes As Integer

Dim Dias_ya_tomados As Integer
Dim Fecha_limite    As Date
Dim Anio_bisiesto   As Boolean
Dim Jornal As Boolean
Dim Legajo As Long
Dim NroTer As Long
Dim Nombre As String
Dim Aux_Fecha As Date
Dim Aux_Generar_Fecha_Hasta As Date

Dim Aux_TipDiaPago As Integer
Dim Aux_TipDiaDescuento As Integer

Dim Quincena_Inicio As Integer
Dim Quincena_Fin As Integer
Dim Quincena_Siguiente As Integer

Dim Total_Dias_A_Generar_Aux
Dim Topedias
Dim rs As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
Dim rs_vacacion As New ADODB.Recordset
Dim rs_Estructura  As New ADODB.Recordset

'Tope máximo de días a pagar
Topedias = 22
'Guardo los dias correspondientes para reasignarlo al final del proceso
Total_Dias_A_Generar_Aux = Total_Dias_A_Generar


Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
    'Exit Sub
End If

StrSql = "SELECT * FROM vacacion WHERE vacnro = " & NroVac
OpenRecordset StrSql, rs_vacacion
If rs_vacacion.EOF Then
    Flog.writeline "No existe el periodo de vacaciones " & NroVac
    Exit Sub
End If

StrSql = "SELECT * FROM vacdiascor WHERE vacdiascor.ternro = " & Ternro
StrSql = StrSql & " AND vacdiascor.vacnro = " & NroVac
StrSql = StrSql & " AND (venc is null OR venc = 0)"
OpenRecordset StrSql, rs_vacdiascor
If rs_vacdiascor.EOF Then
    Flog.writeline "No hay dias correspondientes a ese periodo. SQL: " & StrSql
    Exit Sub
Else
'    If rs_vacdiascor!vdiascorcant <= Topedias Then
'        Total_Dias_A_Generar = rs_vacdiascor!vdiascorcant
'    End If
End If


Mes_Inicio = Month(Aux_Generar_Fecha_Desde)
Ano_Inicio = Year(Aux_Generar_Fecha_Desde)
If Not EsNulo(fecha_hasta) Then
    Mes_Fin = Month(fecha_hasta)
    Ano_Fin = Year(fecha_hasta)
Else
    Mes_Fin = Month(fecha_desde)
    Ano_Fin = Year(fecha_desde)
End If

'Reviso que la cantidad de dias correspondientes del periodo
'no supere la cantidad de dias que quedan por tomar
Aux_TotalDiasCorrespondientes = rs_vacdiascor!vdiascorcant
If rs_vacdiascor!vdiascorcant > Total_Dias_A_Generar Then
    Dias_Afecta = Total_Dias_A_Generar
Else
    Dias_Afecta = rs_vacdiascor!vdiascorcant
End If

If (Total_Dias_A_Generar - Dias_Afecta) < Dias_Afecta Then
    Total_Dias_A_Generar = 0
End If

'Total_Dias_A_Generar = Total_Dias_A_Generar - Dias_Afecta

Aux_Generar_Fecha_Hasta = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)

Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    'Genera_Descuentos = True
End If

If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        If rs!cantdias >= Aux_TotalDiasCorrespondientes Then
            Flog.writeline "No Reproceso y existe el desglose de pago/descuento, salir"
            'Restauro la cantidad de dias porque no se generaron
            Total_Dias_A_Generar = Total_Dias_A_Generar + Dias_Afecta
            Exit Sub
        Else
            Flog.writeline "Existe desglose de pago/descuento para este periodo, Se generarán " & Dias_Afecta & " dias."
        End If
    End If
    rs.Close
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc "
    StrSql = StrSql & " WHERE vacnro = " & NroVac
    StrSql = StrSql & " AND ternro =" & Ternro
    StrSql = StrSql & " AND vacpagdesc.pronro is not null"
    rs.Open StrSql, objConn
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        rs.Close
        Exit Sub
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        rs.Close
        StrSql = "DELETE vacpagdesc "
        StrSql = StrSql & " WHERE vacnro = " & NroVac
        StrSql = StrSql & " AND ternro =" & Ternro
        If Genera_Pagos And Genera_Descuentos Then
            StrSql = StrSql & " AND ( pago_dto = 3 " 'pagos
            StrSql = StrSql & " OR pago_dto = 4)" 'Descuentos
        Else
            If Genera_Pagos Then
                StrSql = StrSql & " AND pago_dto = 3" 'pagos
            End If
            'If Genera_Descuentos Then
                'StrSql = StrSql & " AND pago_dto = 4" 'Descuentos
            'End If
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

'Busco la forma de liquidacion
Flog.writeline "Busco la forma de liquidacion"
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & Ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (his_estructura.htetdesde <= " & ConvFecha(fecha_desde) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(fecha_desde) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not IsNull(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Jornal = True
            Flog.writeline "Jornal"
        Else
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
        Jornal = False
        Flog.writeline "Mensual"
    End If
Else
    Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
    Jornal = False
    Flog.writeline "Mensual"
End If

Aux_TipDiaPago = TipDiaPago
Aux_TipDiaDescuento = TipDiaDescuento
' Reutilizo elparametro TipDiaDescuento de la politica 1503 porque no se utiliza descuento si es jornal
' si es jornal
' TipDiaPago tiene el tprocnro de la primer quincena y
' TipDiaDescuento tiene el tprocnro de la segunda quincena
If Jornal Then
    'reviso a que quincena corresponde
    If Day(Aux_Generar_Fecha_Desde) > 15 Then
        'es segunda quincena
        Aux_TipDiaPago = TipDiaDescuento
        Quincena_Inicio = 2
    Else
        'Es primera quincena
        Aux_TipDiaDescuento = TipDiaPago
        Quincena_Inicio = 1
    End If
    'Determino la quincena de fin
    If Day(Aux_Generar_Fecha_Hasta) > 15 Then
        Quincena_Fin = 2
    Else
        Quincena_Fin = 1
    End If
    Flog.writeline "Quincena Inicio " & Quincena_Inicio
    Flog.writeline "Quincena Fin " & Quincena_Fin
    Flog.writeline "Modelo Pago " & Aux_TipDiaPago
    Flog.writeline "Modelo Descuento " & Aux_TipDiaDescuento
End If
Quincena_Siguiente = Quincena_Inicio


'Esto va dentro del ciclo en esta version
'If Genera_Pagos Then
'    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
'End If


Genera_Dto_DiasCorr = False
Call Politica(1506)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1506. No se generarán los descuentos."
End If

If Genera_Dto_DiasCorr Then

    'tantos descuentos como meses afecte
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    Dias_Pendientes = 0
    Dias_Restantes = Dias_Afecta
    
    'determinar los dias que afecta para el primer mes de decuento
    'Genera 30 dias, para todos los meses
    Anio_bisiesto = EsBisiesto(Ano_Afecta)
    
    If Mes_Afecta = 2 Then
        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        If Not Jornal Then
            'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        Else
            If Quincena_Siguiente = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        End If
    Else
        
        If Not Jornal Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            End If
        Else
            If Quincena_Siguiente = 1 Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        End If
    End If
    
    If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Aux_Generar_Fecha_Hasta)
    Else
        'continua en el mes siguiente
'        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes)
        Dias_Afecta = DateDiff("d", Aux_Generar_Fecha_Desde, Ultimo_Mes) + 1
    End If
    
    'Genero el pago completo
    If Genera_Pagos Then
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Restantes, 3) 'Pago
    End If
    
    
    Do While Not Fin_Licencia And Dias_Restantes > 0
        If Dias_Afecta > 30 Then
            Dias_Afecta = 30
        End If
        
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                Aux_TipDiaPago = TipDiaDescuento
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = TipDiaPago
            End If
        End If
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        'If Genera_Pagos Then
        '    Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaPago, Dias_Afecta, 3) 'Pago
        'End If
        
        If Genera_Descuentos And (Dias_Afecta <> 0) Then
            Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, Aux_TipDiaDescuento, Dias_Afecta, 4) 'Descuento
            'Actualizo la fecha de generacion con la cantidad generada, por si, quedan dias por generar
            Aux_Generar_Fecha_Desde = DateAdd("d", Dias_Afecta, Aux_Generar_Fecha_Desde)
        End If
        'Continua en la siguiente quincena
        If Quincena_Siguiente = 1 Then
            Quincena_Siguiente = 2
        Else
            Quincena_Siguiente = 1
        End If
        
        'determinar a que continua en el proximo mes
        If (Mes_Afecta = 12) Then
            If Not Jornal Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = 1
                    Ano_Afecta = Ano_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        Else
            If Not Jornal Then
                Mes_Afecta = Mes_Afecta + 1
            Else
                If Quincena_Siguiente = 1 Then
                    Mes_Afecta = Mes_Afecta + 1
                Else
                    'Queda como esta, en el mismo mes y año
                End If
            End If
        End If
        
        If (Ano_Afecta = Ano_Fin) Then
            If (Mes_Afecta > Mes_Fin) Then
                Fin_Licencia = True
            Else
                If (Mes_Afecta = Mes_Fin) And Jornal Then
                    If Quincena_Siguiente > Quincena_Fin Then
                        Fin_Licencia = True
                    End If
                End If
            End If
        Else
            If (Ano_Afecta > Ano_Fin) Then
                Fin_Licencia = True
            End If
        End If
        
        'determinar los dias que afecta
        Anio_bisiesto = EsBisiesto(Ano_Afecta)
        If Mes_Afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
            If Not Jornal Then
                'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta + 1, 1, Ano_Afecta), AFecha(Mes_Afecta + 1, 2, Ano_Afecta))
                Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
            Else
                If Quincena_Siguiente = 1 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    'Ultimo_Mes = IIf(Anio_bisiesto = True, AFecha(Mes_Afecta, 29, Ano_Afecta), AFecha(Mes_Afecta, 28, Ano_Afecta))
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            End If
        Else
            If Not Jornal Then
                Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                End If
            Else
                If Quincena_Siguiente = 1 Then
                    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
                    Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
                Else
                    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
                    If Mes_Afecta = 12 Then
                        Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                    Else
                        Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
                    End If
                End If
            End If
        End If
    
    
        If (Aux_Generar_Fecha_Hasta <= Ultimo_Mes) Then
        'termina en el mes
            If (Dias_Restantes < DateDiff("d", Primero_Mes, Ultimo_Mes)) Then
                Dias_Afecta = Dias_Restantes
            Else
                Dias_Afecta = DateDiff("d", Primero_Mes, Aux_Generar_Fecha_Hasta) + 1
            End If
        Else
            'continua en el mes siguiente
            Dias_Afecta = DateDiff("d", Primero_Mes, Ultimo_Mes) + 1
        End If
        
        If Not Jornal Then
            If Dias_Afecta > 30 Then
               Dias_Afecta = 30
            End If
        Else
            If Dias_Afecta > 15 Then
               Dias_Afecta = 15
            End If
        End If
        
        Aux_TipDiaPago = TipDiaPago
        Aux_TipDiaDescuento = TipDiaDescuento
        
        If Jornal Then
            'reviso a que quincena corresponde
            If Day(Primero_Mes) > 15 Then
                'es segunda quincena
                Aux_TipDiaPago = TipDiaDescuento
            Else
                'Es primera quincena
                Aux_TipDiaDescuento = TipDiaPago
            End If
        End If
    Loop
Else
    Flog.writeline "No se generan los descuentos."
End If



'Cierro todo
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close
If rs.State = adStateOpen Then rs.Close
If rs_vacacion.State = adStateOpen Then rs_vacacion.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_vacdiascor = Nothing
Set rs = Nothing
Set rs_vacacion = Nothing
Set rs_Estructura = Nothing
End Sub



Public Sub Politica1500v_TATA()
'***************************************************************
'  Paga todo y descuenta por mes teniendo en cuenta dias habiles y la fecha de inicio de la licencia
'  Especialmenet creada para TATA
'  Se puede reutilizar en Uruguay
' Creado: FGZ - 27/01/2015
'***************************************************************
Dim rs As New Recordset
Dim StrSql As String

Dim Dia_Inicio As Integer
Dim Mes_Inicio As Integer
Dim Ano_Inicio As Integer
Dim Dia_Fin    As Integer
Dim Mes_Fin    As Integer
Dim Ano_Fin    As Integer
Dim Fin_Licencia  As Boolean
Dim Mes_Afecta    As Integer
Dim Ano_Afecta    As Integer
Dim Primero_Mes   As Date
Dim Ultimo_Mes    As Date
Dim i As Date
Dim Dias_Afecta   As Integer
'Dim Dias_Pendientes As Integer
Dim Dias_Restantes As Integer
Dim Dias_Hab_Afecta As Integer
Dim objFeriado      As New Feriado
Dim EsFeriado       As Boolean
Dim qna_afecta As Integer
Dim estrnro_GrupoLiq As Integer

Dim Anio_bisiesto   As Boolean
Dim Jornal As Integer
Dim NroTer As Long

On Error GoTo CE

Set objFeriado.Conexion = objConn


Call Politica(1503)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1503. Seteo Parametros default."
    TipDiaPago = 3
    TipDiaDescuento = 3
End If

'Busco la licencia dentro del intervalo especificado
StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
OpenRecordset StrSql, rs

' Si no hay licencia me voy
If rs.EOF Then
    Flog.wirteline Espacios(Tabulador * 1) & "Licencia inexistente." & nrolicencia
    GoTo CE
Else
    NroTer = rs!Empleado
    Dia_Inicio = Day(rs!elfechadesde)
    Mes_Inicio = Month(rs!elfechadesde)
    Ano_Inicio = Year(rs!elfechadesde)
    Dia_Fin = Day(rs!elfechahasta)
    Mes_Fin = Month(rs!elfechahasta)
    Ano_Fin = Year(rs!elfechahasta)
End If


Genera_Pagos = False
Genera_Descuentos = False
Call Politica(1507)
If Not PoliticaOK Then
    Flog.writeline "Error cargando configuracion de la Politica 1507. Default. Se generarán los pagos y los descuentos."
    Genera_Pagos = True
    Genera_Descuentos = True
End If

'VERIFICAR el reproceso y manejar la depuracion
If Not Reproceso Then
    'si no es reproceso y existe el desglose de pago/descuento, salir
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline Espacios(Tabulador * 1) & "Reproceso y existe el desglose de pago/descuento, salir"
        GoTo CE
    End If
Else
    'verificar si no tiene pagos/descuentos ya procesados, sino depurarlos
    StrSql = "SELECT * FROM vacpagdesc WHERE emp_licnro = " & nrolicencia & " AND vacpagdesc.pronro is not null"
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        Flog.writeline "No se puede Reprocesar, hay pagos y/o descuentos Liquidados. "
        GoTo CE
    Else
        'DEPURAR LOS PAGOS/DESCUENTOS GENERADOS PARA LA LICENCIA
        StrSql = "DELETE FROM vacpagdesc WHERE emp_licnro = " & nrolicencia
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If


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
Dias_Afecta = rs!elcantdias

If Genera_Pagos Then
    If Pliq_Nro <> 0 Then
        Mes_Inicio = Pliq_Mes
        Ano_Inicio = Pliq_Anio
    End If
    Call Generar_PagoDescuento(Mes_Inicio, Ano_Inicio, TipDiaPago, Dias_Afecta, 3) 'Pago
End If


    StrSql = "SELECT estructura.estrnro, estructura.estrcodext FROM estructura " & _
            " INNER JOIN his_estructura ON estructura.estrnro = his_estructura.estrnro " & _
            " INNER JOIN emp_lic ON his_estructura.ternro = emp_lic.empleado " & _
            " WHERE his_estructura.tenro = 32 AND his_estructura.htethasta IS NULL " & _
            " AND emp_lic.emp_licnro= " & nrolicencia
    OpenRecordset StrSql, rs
    If rs.EOF Then
        Jornal = 1  ' 7 dias
        estrnro_GrupoLiq = 0
        Flog.writeline "    El empleado no posee Grupo de Liquidación. Se considera la semana corrida (7 días)."
    Else
        estrnro_GrupoLiq = rs!estrnro
        Select Case rs!estrcodext
            Case "1", "2":  ' 7 dias
                Jornal = 1
                Flog.writeline "    Según el Cód. Ext. (1 ó 2) del Grupo de Liquidación del empleado, se considera la semana corrida (7 días)."
            Case "3", "5":  ' 6 dias
                Jornal = 3
                Flog.writeline "    Según el Cód. Ext. (3 ó 5) del Grupo de Liquidación del empleado, se excluirán los domingos y feriados de la semana (6 días)."
            Case "4":       ' 5 dias
                Jornal = 4
                Flog.writeline "    Según el Cód. Ext. (4) del Grupo de Liquidación del empleado, se excluirán los domingos, sabados y feriados de la semana (5 días)."
            Case "6":       ' 7 dias
                Jornal = 2
                Flog.writeline "    Según el Cód. Ext. (6) del Grupo de Liquidación del empleado, se considera la semana corrida (7 días)."
            Case Else       ' 7 dias
                Jornal = 1
                Flog.writeline "    No está definido el Cód. Ext. del Grupo de Liquidación del empleado, se considera la semana corrida (7 días)."
        End Select
    End If
    rs.Close
    
    
    StrSql = "SELECT * FROM emp_lic WHERE emp_licnro= " & nrolicencia
    rs.Open StrSql, objConn
    Dias_Hab_Afecta = 0
    If Jornal > 2 Then
        For i = rs!elfechadesde To rs!elfechahasta
            EsFeriado = objFeriado.Feriado(i, rs!Empleado, False)
            Select Case Jornal
                Case 3:
                    If Not (EsFeriado Or Weekday(i) = 1) Then
                        Dias_Hab_Afecta = Dias_Hab_Afecta + 1
                    End If
                Case 4:
                    If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                        Dias_Hab_Afecta = Dias_Hab_Afecta + 1
                    End If
            End Select
        Next
    Else
        Dias_Hab_Afecta = rs!elcantdias
    End If
    
    Flog.writeline "    Días de la licencia: " & Dias_Afecta
    Flog.writeline "    Días que se descontaran en la licencia: " & Dias_Hab_Afecta
    
    Fin_Licencia = False
    Mes_Afecta = Mes_Inicio
    Ano_Afecta = Ano_Inicio
    'Dias_Pendientes = 0
    Dias_Restantes = Dias_Hab_Afecta
    
    If Jornal > 2 Then
        If Dia_Inicio >= 16 Then
            qna_afecta = 2
        Else
            qna_afecta = 1
        End If
    Else
        If Jornal = 1 Then
            qna_afecta = 3
        Else
            qna_afecta = 5
        End If
    End If
    Flog.writeline "    Modelo de Liquidación (1.- Primera Quincena, 2.- Segunda Quincena, 3.- Mensuales, 5.- Liq. Final): " & qna_afecta
    
    '/* determinar los días que afecta para el primer mes de decuento */
    '/* Genera 30 dias mínimos, para todos los meses */
    Anio_bisiesto = EsBisiesto(Ano_Afecta)

    'If qna_afecta = 2 Then
    '    Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
    'Else
    '    Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
    'End If
    
    Primero_Mes = AFecha(Mes_Afecta, Dia_Inicio, Ano_Afecta)
        
    If Mes_Afecta = 2 Then
        If qna_afecta = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta) - 1
        End If
    Else
        If qna_afecta = 1 Then
            Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
        Else
            If Mes_Afecta = 12 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
            Else
                Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
            End If
        End If
    End If

    If (rs!elfechahasta <= Ultimo_Mes) Then
        '/* termina en el mes */
        Dias_Afecta = rs!elcantdias
        If Jornal > 2 Then
            Dias_Afecta = Dias_Hab_Afecta
        End If
    Else
        '/* continua en el mes siguiente */
        Dias_Afecta = DateDiff("d", rs!elfechadesde, Ultimo_Mes) + 1
        If Jornal > 2 Then
            Dias_Afecta = 0
            For i = rs!elfechadesde To Ultimo_Mes
                EsFeriado = objFeriado.Feriado(i, rs!Empleado, False)
                Select Case Jornal
                    Case 3:
                        If Not (EsFeriado Or Weekday(i) = 1) Then
                            Dias_Afecta = Dias_Afecta + 1
                        End If
                    Case 4:
                        If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                            Dias_Afecta = Dias_Afecta + 1
                        End If
                End Select
            Next
        End If
    End If
    
    Do While Not Fin_Licencia
        'FGZ - 22/01/2015 ---------------------
        If Dias_Afecta > Dias_Restantes Then
            Dias_Afecta = Dias_Restantes
        End If
        'FGZ - 22/01/2015 ---------------------
        Flog.writeline "    Días entre el " & Primero_Mes & " y el " & Ultimo_Mes & ": " & Dias_Afecta
        
        Dias_Restantes = Dias_Restantes - Dias_Afecta
        
        Call Generar_PagoDescuento(Mes_Afecta, Ano_Afecta, qna_afecta, Dias_Afecta, 4) 'Dto
         
        Flog.writeline "    Días restantes: " & Dias_Restantes
        
        '/* determinar a que continua en el proximo mes */
        If qna_afecta = 1 Then
            qna_afecta = 2
        Else
            If qna_afecta = 2 Then
                qna_afecta = 1
            End If
            If (Mes_Afecta = 12) Then
                Mes_Afecta = 1
                Ano_Afecta = Ano_Afecta + 1
            Else
                Mes_Afecta = Mes_Afecta + 1
            End If
        End If
        
'        If (Mes_Afecta > Mes_Fin And Ano_Afecta >= Ano_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
'        If (Mes_Afecta > Mes_Fin) Or (Mes_Afecta = 1 And Mes_Fin = 12) Then
        If (Dias_Restantes <= 0) Then
            Fin_Licencia = True
        End If

        '/* determinar los días que afecta */
        Anio_bisiesto = EsBisiesto(Ano_Afecta)

        Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        
        If qna_afecta = 2 Then
            Primero_Mes = AFecha(Mes_Afecta, 16, Ano_Afecta)
        Else
            Primero_Mes = AFecha(Mes_Afecta, 1, Ano_Afecta)
        End If
            
        If Mes_Afecta = 2 Then
            If qna_afecta = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                If Anio_bisiesto Then
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 1, Ano_Afecta)
                Else
                    Ultimo_Mes = AFecha(Mes_Afecta + 1, 2, Ano_Afecta)
                End If
            End If
        Else
            If qna_afecta = 1 Then
                Ultimo_Mes = AFecha(Mes_Afecta, 15, Ano_Afecta)
            Else
                If Mes_Afecta = 12 Then
                    Ultimo_Mes = AFecha(Mes_Afecta, 31, Ano_Afecta)
                Else
                    Ultimo_Mes = DateAdd("d", -1, AFecha(Mes_Afecta + 1, 1, Ano_Afecta))
                End If
            End If
        End If
        
        If (rs!elfechahasta <= Ultimo_Mes) Then
            '/* termina antes del ultimo dia del mes */
            Dias_Afecta = 0
            For i = Primero_Mes To rs!elfechahasta
                If Jornal > 2 Then
                    EsFeriado = objFeriado.Feriado(i, rs!Empleado, False)
                    Select Case Jornal
                        Case 3:
                            If Not (EsFeriado Or Weekday(i) = 1) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                        Case 4:
                            If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                    End Select
                Else
                    Dias_Afecta = Dias_Afecta + 1
                End If
            Next
        Else
            '/* continua en el mes siguiente */
            Dias_Afecta = 0
            For i = Primero_Mes To Ultimo_Mes
                If Jornal > 2 Then
                    EsFeriado = objFeriado.Feriado(i, rs!Empleado, False)
                    Select Case Jornal
                        Case 3:
                            If Not (EsFeriado Or Weekday(i) = 1) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                        Case 4:
                            If Not (EsFeriado Or Weekday(i) = 1 Or Weekday(i) = 7) Then
                                Dias_Afecta = Dias_Afecta + 1
                            End If
                    End Select
                Else
                    Dias_Afecta = Dias_Afecta + 1
                End If
            Next
        End If
    Loop
GoTo ProcesadoOK

CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

ProcesadoOK:
    rs.Close
    Set rs = Nothing

End Sub


Private Sub DiasHabilesLic(ByVal Ternro As Long, ByVal tipoVac As Long, ByVal FechaInicial As Date, ByVal Fechafinal As Date, ByRef CHabiles As Integer, ByRef cNoHabiles As Integer, ByRef cFeriados As Integer)
'-----------------------------------------------------------------------
' Procedimiento
'       Calcula la cantidad de dias entre 2 fechas (habiles, no habiles, feriados)
'Autor: FGZ
'Ultima Mod: FGZ - 18/11/2010
'-----------------------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim objFeriado As New Feriado
Dim DHabiles(1 To 7) As Boolean
Dim EsFeriado As Boolean
Dim objRs As New ADODB.Recordset
Dim ExcluyeFeriadosNoLab As Boolean
Dim ExcluyeFeriadosLab As Boolean
Dim Fecha As Date

Dim AuxFechaDesde As Date
Dim AuxFechaHasta As Date
Dim Quedandias As Boolean

'Dim CHabiles As Integer
'Dim cNoHabiles As Integer
'Dim cFeriados As Integer
Dim CortarLicencias As Boolean

    'Por default deberia cortar las licencias
    CortarLicencias = True
    
    StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & tipoVac
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        DHabiles(1) = objRs!tpvhabiles__1
        DHabiles(2) = objRs!tpvhabiles__2
        DHabiles(3) = objRs!tpvhabiles__3
        DHabiles(4) = objRs!tpvhabiles__4
        DHabiles(5) = objRs!tpvhabiles__5
        DHabiles(6) = objRs!tpvhabiles__6
        DHabiles(7) = objRs!tpvhabiles__7
    
        ExcluyeFeriadosNoLab = CBool(objRs!tpvferiado)
        
        If Not EsNulo(objRs!excferilab) Then
            ExcluyeFeriadosLab = CBool(objRs!excferilab)
        Else
            ExcluyeFeriadosLab = False
        End If
        
        If Not EsNulo(objRs!tpvprog3) Then
            CortarLicencias = CBool(objRs!tpvprog3)
        End If
    Else
        Flog.writeline "No se encontro el tipo de Vacacion " & tipoVac
        Exit Sub
    End If
    
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = objConn
    
    i = 0
    j = 0
    CHabiles = 0
    cNoHabiles = 0
    cFeriados = 0
   
   AuxFechaDesde = FechaInicial
   
   Fecha = FechaInicial
    
    Do While Fecha <= Fechafinal
        Quedandias = True
        EsFeriado = objFeriado.Feriado(Fecha, Ternro, False)
        
        'FGZ - 10/08/2010 -----------------------
        If Not EsFeriado Then
            If DHabiles(Weekday(Fecha)) Then
                i = i + 1
                CHabiles = CHabiles + 1
            Else
                cNoHabiles = cNoHabiles + 1
                If PoliticaOK And Diashabiles_LV Then
                    i = i + 1
                End If
            End If
        Else    'Es feriado
            If Feriado_Laborable Then
                If ExcluyeFeriadosLab Then
                    If DHabiles(Weekday(Fecha)) Then
                        i = i + 1
                        CHabiles = CHabiles + 1
                    Else
                        cNoHabiles = cNoHabiles + 1
                        If PoliticaOK And Diashabiles_LV Then
                            i = i + 1
                        End If
                    End If
                Else
                    cFeriados = cFeriados + 1
                    If PoliticaOK And Diashabiles_LV Then
                        i = i + 1
                    End If
                End If
            Else 'Feriado no laborable
                If ExcluyeFeriadosNoLab Then
                    If DHabiles(Weekday(Fecha)) Then
                        i = i + 1
                        CHabiles = CHabiles + 1
                    Else
                        cNoHabiles = cNoHabiles + 1
                        If PoliticaOK And Diashabiles_LV Then
                            i = i + 1
                        End If
                    End If
                Else
                    cFeriados = cFeriados + 1
                End If
            End If
        End If
        
        Fecha = DateAdd("d", 1, Fecha)
    Loop
    Set objFeriado = Nothing
End Sub

'MDZ - 16/02/2016 - Paso a publica
Public Function CalcFechaHastaLic(ByVal Ternro As Long, ByVal tipoVac As Long, ByVal Aux_Fecha_Desde As Date, ByVal Dias_Dto As Integer)
'-----------------------------------------------------------------------
' Procedimiento
'       Calcula la fecha hasta que debe tener una licencia para que la cantidad de dias habiles sea los que se pasan por parametro
'Autor: FGZ - 30/11/2010
'Ultima Mod:
'-----------------------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim objFeriado As New Feriado
Dim DHabiles(1 To 7) As Boolean
Dim EsFeriado As Boolean
Dim objRs As New ADODB.Recordset
Dim ExcluyeFeriadosNoLab As Boolean
Dim ExcluyeFeriadosLab As Boolean
Dim Fecha As Date

Dim AuxFechaDesde As Date
Dim AuxFechaHasta As Date
Dim Quedandias As Boolean

Dim CHabiles As Integer
Dim cNoHabiles As Integer
Dim cFeriados As Integer
Dim CortarLicencias As Boolean

    'Por default deberia cortar las licencias
    CortarLicencias = True
    
    StrSql = "SELECT * FROM tipovacac WHERE tipvacnro = " & tipoVac
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        DHabiles(1) = objRs!tpvhabiles__1
        DHabiles(2) = objRs!tpvhabiles__2
        DHabiles(3) = objRs!tpvhabiles__3
        DHabiles(4) = objRs!tpvhabiles__4
        DHabiles(5) = objRs!tpvhabiles__5
        DHabiles(6) = objRs!tpvhabiles__6
        DHabiles(7) = objRs!tpvhabiles__7
    
        ExcluyeFeriadosNoLab = CBool(objRs!tpvferiado)
        
        If Not EsNulo(objRs!excferilab) Then
            ExcluyeFeriadosLab = CBool(objRs!excferilab)
        Else
            ExcluyeFeriadosLab = False
        End If
        
        If Not EsNulo(objRs!tpvprog3) Then
            CortarLicencias = CBool(objRs!tpvprog3)
        End If
    Else
        Flog.writeline "No se encontro el tipo de Vacacion " & tipoVac
        Exit Function
    End If
    
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = objConn
    
    i = 0
    j = 0
    CHabiles = 0
    cNoHabiles = 0
    cFeriados = 0
   
   'AuxFechaDesde = Aux_Fecha_Desde
   Fecha = Aux_Fecha_Desde
    
    'Do While Fecha <= Fechafinal
    Do While CHabiles < Dias_Dto
        Quedandias = True
        EsFeriado = objFeriado.Feriado(Fecha, Ternro, False)
        
        'FGZ - 10/08/2010 -----------------------
        If Not EsFeriado Then
            If DHabiles(Weekday(Fecha)) Then
                i = i + 1
                CHabiles = CHabiles + 1
            Else
                cNoHabiles = cNoHabiles + 1
                If PoliticaOK And Diashabiles_LV Then
                    i = i + 1
                End If
            End If
        Else    'Es feriado
            If Feriado_Laborable Then
                If ExcluyeFeriadosLab Then
                    If DHabiles(Weekday(Fecha)) Then
                        i = i + 1
                        CHabiles = CHabiles + 1
                    Else
                        cNoHabiles = cNoHabiles + 1
                        If PoliticaOK And Diashabiles_LV Then
                            i = i + 1
                        End If
                    End If
                Else
                    cFeriados = cFeriados + 1
                    If PoliticaOK And Diashabiles_LV Then
                        i = i + 1
                    End If
                End If
            Else 'Feriado no laborable
                If ExcluyeFeriadosNoLab Then
                    If DHabiles(Weekday(Fecha)) Then
                        i = i + 1
                        CHabiles = CHabiles + 1
                    Else
                        cNoHabiles = cNoHabiles + 1
                        If PoliticaOK And Diashabiles_LV Then
                            i = i + 1
                        End If
                    End If
                Else
                    cFeriados = cFeriados + 1
                End If
            End If
        End If
        
        If CHabiles < Dias_Dto Then
            Fecha = DateAdd("d", 1, Fecha)
        End If
    Loop
    CalcFechaHastaLic = Fecha

Set objFeriado = Nothing
    
End Function



Public Function Minimo(ByVal X, ByVal Y)
    If X <= Y Then
        Minimo = X
    Else
        Minimo = Y
    End If
End Function

'EAM- Obtiene la forma de liquidación de un empleado (Mensual (false)| Jornal (True))
Function FormaDeLiquidacion(ByVal Aux_Fecha As Date, ByVal Tenro As Long)
 Dim rsAux As New ADODB.Recordset
 Dim Jornal As Boolean
 
    Flog.writeline "Busco la forma de liquidacion"
    StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura" & _
            " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " & _
            " WHERE ternro = " & Ternro & " AND his_estructura.tenro = " & Tenro & " AND  (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ")" & _
            " AND  ((" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rsAux
    
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!estrcodext) Then
            If rsAux!estrcodext = "2" Then
                Jornal = True
                Flog.writeline "Jornal"
            Else
                Jornal = False
                Flog.writeline "Mensual"
            End If
        Else
            Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
            Jornal = False
            Flog.writeline "Mensual"
        End If
    Else
        Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
        Jornal = False
        Flog.writeline "Mensual"
    End If
    
    Set rsAux = Nothing
    FormaDeLiquidacion = Jornal
End Function

Function FormaDeLiquidacion_sykes(ByVal Aux_Fecha As Date, ByVal Tenro As Long)
 
'Fecha 18/03/2015
'autor: Matias Fernandez
'causa: no sirve chequear por codigo externo ya que ellos ponen la letra M a todas las estructuras de tipo de estructura 22
'entonces se asocio un codigo, si la estructura tiene ese codigo entonces es jornal, sino es mensual.
 
 Dim rsAux As New ADODB.Recordset
 Dim Jornal As Boolean
 Dim estruct
    Jornal = False
    Flog.writeline "Busco la forma de liquidacion: FormaDeLiquidacion_sykes"
    StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura" & _
            " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " & _
            " WHERE ternro = " & Ternro & " AND his_estructura.tenro = " & Tenro & " AND  (his_estructura.htetdesde <= " & ConvFecha(Aux_Fecha) & ")" & _
            " AND  ((" & ConvFecha(Aux_Fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
    OpenRecordset StrSql, rsAux
    Flog.writeline StrSql
   ' If Not rsAux.EOF Then
   '     If Not IsNull(rsAux!estrcodext) Then
   '         If rsAux!estrcodext = "2" Then
   '             Jornal = True
   '             Flog.writeline "Jornal"
   '         Else
   '             Jornal = False
   '             Flog.writeline "Mensual"
   '         End If
   '     Else
   '         Flog.writeline "Codigo externo de la Forma de liq Nulo. Se asume No Jornal"
   '         Jornal = False
   '         Flog.writeline "Mensual"
   '     End If
   ' Else
   '     Flog.writeline "No existe Forma de liq. Se asume No Jornal. SQL: " & StrSql
   '     Jornal = False
   '     Flog.writeline "Mensual"
   ' End If
    
    If Not rsAux.EOF Then
     estruct = rsAux!estrnro
     StrSql = "select tcodnro from estr_cod inner join estructura on estructura.estrnro = estr_cod.estrnro "
     StrSql = StrSql & " where estructura.estrnro= " & estruct & " and tcodnro = 191"
     OpenRecordset StrSql, rsAux
     Flog.writeline StrSql
        If rsAux.EOF Then
          Flog.writeline "No tiene el codigo asociado, jornal= false"
          Jornal = False
        Else
          Flog.writeline "tiene el codigo asociado, jornal= true"
          Jornal = True
        End If
    End If
    Flog.writeline "fin FormaDeLiquidacion_sykes"
    Set rsAux = Nothing
    FormaDeLiquidacion_sykes = Jornal
End Function
'EAM- Obtiene la forma de liquidación de un empleado (Mensual (false)| Jornal (True))
Function ExisteTipoLicencia(ByVal tdnro As Integer)
 Dim rsAux As New ADODB.Recordset
  
    StrSql = "SELECT * FROM tipdia WHERE tdnro = " & tdnro
    OpenRecordset StrSql, rsAux
    If rsAux.EOF Then
        ExisteTipoLicencia = False
    Else
        ExisteTipoLicencia = True
    End If
End Function
