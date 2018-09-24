Attribute VB_Name = "PrcMasivosPoliticasVac"
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
            ' Call politica1500(det, Cabecera, Detalle)
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
'        TipDiaPago = st_ModeloPago
'        TipDiaDescuento = st_ModeloDto
        PoliticaOK = True
    Else
        PoliticaOK = False
    End If
    
End Sub

Private Sub politica1504(ByVal subn As Long, ByVal Cabecera As Long, ByVal Detalle As Long)

    PoliticaOK = True
End Sub



Public Function AFecha(m As Integer, d As Integer, a As Integer) As Date
' Reemplaza a la función Date de Progress
'ultimo-mes  = DATE (mes-afecta,30,ano-afecta)
Dim auxi
  
  auxi = Str(m) & "/" & Str(d) & "/" & Str(a)
  AFecha = Format(auxi, "mm/dd/yyyy")

End Function




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

