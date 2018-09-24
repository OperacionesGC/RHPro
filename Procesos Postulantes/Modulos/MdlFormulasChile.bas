Attribute VB_Name = "MdlFormulasChile"
' ---------------------------------------------------------
' Modulo de fórmulas conocidas para Chile
' ---------------------------------------------------------
'Tipos
Public Type TregImpunicab
    periodo As Long
    monto1 As Double
    monto2 As Double
    monto3 As Double
    monto4 As Double
    monto5 As Double
    perMes As Integer
    perAnio As Long
    perUTMHist As Double
    fechaRet As Date
    PeriodoDesde As Date
    periodoDesc As String
End Type


Public Function for_ImpuestoUnico(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Impuesto Unico.
' Autor      :
' Fecha      :
' Ultima Mod.: 03/12/2006
'              19/01/2009 - Martin - Busca tb en acu_mes cuando busca los acumuladores de la liquidacion
' ---------------------------------------------------------------------------------------------

Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim p_prorratea As Integer  'Si prorratea o no para liq. finales

'Variables Locales
Dim Devuelve As Double
Dim val_concepto As Double
Dim Neto As Double
Dim prorratea As Double
Dim Retencion As Double
Dim Gan_Imponible As Double
Dim Deducciones As Double
Dim Descuentos As Double
Dim Ded_a23 As Double
Dim Por_Deduccion As Double
Dim Impuesto_Escala As Double
Dim Ret_Ant As Double

Dim Ret_Mes As Integer
Dim Ret_Ano As Integer
Dim fin_mes_ret As Date
Dim ini_anyo_ret As Date
Dim Con_liquid As Integer
Dim I As Integer
Dim j As Integer
Dim Texto As String

'Vectores para manejar el proceso
Dim Items_DDJJ(50) As Double
Dim Items_LIQ(50) As Double
Dim Items_PRORR(50) As Double
Dim Items_OLD_LIQ(50) As Double
Dim Items_TOPE(50) As Double
Dim Items_ART_23(50) As Boolean

'Recorsets Auxiliares
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Item As New ADODB.Recordset
Dim rs_valitem As New ADODB.Recordset
Dim rs_Desmen As New ADODB.Recordset
Dim rs_Desliq As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_escala As New ADODB.Recordset
Dim rs_Ficharet As New ADODB.Recordset
Dim rs_WF_EscalaUTM As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset

Dim Hasta As Integer

' FGZ - 12/02/2004
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Double
Dim val_impdebitos As Double

' FGZ - 12/02/2004

' FGZ - 27/02/2004
Dim Terminar As Boolean
Dim pos1
Dim pos2
' FGZ - 27/02/2004

'Parametro unico
p_concepto = 1005

Bien = False


'FGZ - 19/04/2004
Dim Total_Empresa As Double
Dim Tope As Integer
Dim rs_modelo As New ADODB.Recordset
Dim rs_Traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Double

Total_Empresa = 0
Tope = 10
Descuentos = 0
   
' Creo la tabla temporal para la Escala de UTM
 Call CreateTempTable(TTempWF_EscalaUTM)

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
ini_anyo_ret = C_Date("01/01/" & Ret_Ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).concnro

' Carga la escala de UTM y Multiplica la escala del Imp. Unico por el valor de UTM del Periodo que se esta liquidando */
Call insertar_wf_escalautm(buliq_periodo!pliqmes, buliq_periodo!pliqanio, buliq_periodo!pliqutm)
     
'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case p_concepto:
        val_concepto = rs_wf_tpa!Valor
    End Select
   
    rs_wf_tpa.MoveNext
Loop

' Si el proceso es de gratificai¢n, devuelve el valor sin generar nada
'Busca el proceso de Gratificaci¢n
 StrSql = "SELECT * FROM tipoproc WHERE tprocrecalculo = -1 AND tipoproc.tprocnro = " & buliq_proceso!tprocnro
 OpenRecordset StrSql, rs_modelo

 If Not rs_modelo.EOF Then
         Bien = True
         exito = Bien
         for_ImpuestoUnico = -val_concepto
 End If



If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 3) & "---------Formula-----------------------------"
    Flog.writeline Espacios(Tabulador * 3) & "Fecha del Proceso " & buliq_proceso!profecpago
    Flog.writeline Espacios(Tabulador * 3) & "Mes de Retencion " & Ret_Mes
    Flog.writeline Espacios(Tabulador * 3) & "Año de Retencion " & Ret_Ano
    Flog.writeline Espacios(Tabulador * 3) & "Fin mes de Retencion " & fin_mes_ret
    
End If


' Recorro todos los items de Ganancias
StrSql = "SELECT * FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_Item

Do While Not rs_Item.EOF
 
    Select Case rs_Item!itetipotope
    
    Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
        ' Busco la declaracion jurada
        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!ternro & _
                 " AND desano=" & Ret_Ano & _
                 " AND itenro = " & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
        
        Do While Not rs_Desmen.EOF
            'If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
            If Month(rs_Desmen!desfecdes) = Ret_Mes Then
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                    
            End If
            
            rs_Desmen.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_Item!Itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acuNro)
            
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                   If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
                    Else
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
                    End If
            End If
            
            '19/01/2009 - Martin - Busca tb en acu_mes cuando busca los acumuladores de la liquidacion
            StrSql = "SELECT ammonto, amcant"
            StrSql = StrSql & " FROM acu_mes"
            StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro
            StrSql = StrSql & " AND acunro = " & Acum
            StrSql = StrSql & " AND  amanio = " & buliq_periodo!pliqanio
            StrSql = StrSql & " AND ammes = " & buliq_periodo!pliqmes
            OpenRecordset StrSql, rs_Acu_Mes
            
            If Not rs_Acu_Mes.EOF Then
                Aux_Acu_Monto = rs_Acu_Mes!ammonto
        
                   If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
                    Else
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
                    End If
            End If
            
            rs_itemacum.MoveNext
        Loop
        
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_Item!Itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
    ' End case 2
    ' ------------------------------------------------------------------------
        
'   Case Else:
    End Select
    
            
    'Armo la traza del item
    If CBool(USA_DEBUG) Then
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-DDJJ" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_DDJJ(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Liq" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_LIQ(rs_Item!Itenro)
    End If
    If HACE_TRAZA Then
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_LIQ(rs_Item!Itenro))
    End If
        
   '   /* Calculo Imponible del Imp. Unico */
   ' assign gan-imponible = gan-imponible +
   '                        Items-DDJJ[item.itenro] +
   '                        Items-LIQ[item.itenro].

    Gan_Imponible = Gan_Imponible + Items_DDJJ(rs_Item!Itenro) + Items_LIQ(rs_Item!Itenro)
            
    rs_Item.MoveNext
Loop
            
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        Flog.writeline Espacios(Tabulador * 3) & "9- Imponible del Imp. Unico: " & Gan_Imponible
          End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Imponible del Imp. Unico ", Gan_Imponible)
        End If
            
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        Flog.writeline Espacios(Tabulador * 3) & "9- UTM del Periodo: " & buliq_periodo!pliqutm
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- UTM del Periodo ", buliq_periodo!pliqutm)
     End If
    
    
  'Entra en la escala actualizada por el valor de UTM */
  'run impuni03.p(gan-imponible, output desde-esc,
  '                              output hasta-esc,
  '                              output factor-esc,
  '                              output rebaja-esc).
  
  'Call Entra_EscalaUTM(buliq_periodo!pliqutm)
  
'   find first escala-utm where (valor > escala-utm.desde) and
'                              (valor <= escala-utm.hasta) no-lock no-error.
'  if avail(escala-utm)
'  then
'    assign val - Desde = Escala - utm.Desde
'           val -Hasta = Escala - utm.Hasta
'           val -factor = Escala - utm.factor
'           val-red = escala-utm.rebaja.

  StrSql = "SELECT * FROM " & TTempWF_EscalaUTM & _
           " WHERE desde < " & Gan_Imponible & _
           " AND hasta >= " & Gan_Imponible
            OpenRecordset StrSql, rs_WF_EscalaUTM
  If Not rs_WF_EscalaUTM.EOF Then
  
   Desde_esc = rs_WF_EscalaUTM!Desde
   Hasta_esc = rs_WF_EscalaUTM!Hasta
   factor_esc = rs_WF_EscalaUTM!factor
   rebaja_esc = rs_WF_EscalaUTM!rebaja
 
   If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        Flog.writeline Espacios(Tabulador * 3) & "9- Escala Desde: " & Desde_esc
        Flog.writeline Espacios(Tabulador * 3) & "9- Escala Hasta: " & Hasta_esc
        Flog.writeline Espacios(Tabulador * 3) & "9- Factor Escala: " & factor_esc
        Flog.writeline Espacios(Tabulador * 3) & "9- Rebaja Escala: " & rebaja_esc
        
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Escala Desde: ", Desde_esc)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Escala Hasta: ", Hasta_esc)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Factor Escala: ", factor_esc)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Rebaja Escala: ", rebaja_esc)
        
     End If
   Else
       If CBool(USA_DEBUG) Then
          Flog.writeline Espacios(Tabulador * 3) & "9- No se encontro la escala para el valor " & Gan_Imponible
       End If
       If HACE_TRAZA Then
          Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- No se encontro la escala para el valor: ", Gan_Imponible)
       End If
   End If

                
    ' Calculo las retenciones ya realizadas
    Ret_Ant = 0
        
   
    'como no puede utilizar la funcion month() en sql
    'levanto todas las ficharet del tercero y hago la pregunta dentro del loop
    StrSql = "SELECT * FROM ficharet " & _
             " INNER JOIN proceso ON proceso.pronro = ficharet.pronro " & _
             " AND proceso.pliqnro = " & buliq_periodo!PliqNro & _
             " WHERE empleado = " & buliq_empleado!ternro
             
    OpenRecordset StrSql, rs_Ficharet
    
    Do While Not rs_Ficharet.EOF
        If (Month(rs_Ficharet!Fecha) <= Ret_Mes) And (Year(rs_Ficharet!Fecha) = Ret_Ano) Then
            Ret_Ant = Ret_Ant + rs_Ficharet!importe
        End If
        rs_Ficharet.MoveNext
    Loop
    
    
    'Calcular la retencion
    'assign Monto - calculado = (gan - imponible * factor - esc) - rebaja - esc
    '    retencion = monto-calculado - ret-ant.

    Monto_calculado = (Gan_Imponible * factor_esc) - rebaja_esc
    Retencion = Monto_calculado - Ret_Ant
    
'    StrSql = "SELECT * FROM traza_gan WHERE"
'    StrSql = StrSql & " pliqnro = " & buliq_periodo!PliqNro
'    StrSql = StrSql & " AND pronro = " & buliq_proceso!pronro
'    StrSql = StrSql & " AND concnro = " & Buliq_Concepto(Concepto_Actual).concnro
'    StrSql = StrSql & " AND fecha_pago = " & ConvFecha(buliq_proceso!profecpago)
'    StrSql = StrSql & " AND ternro = " & buliq_empleado!ternro
'    OpenRecordset StrSql, rs_Ficharet
'    If Not rs_Ficharet.EOF Then
'        StrSql = "INSERT INTO traza_gan ("
'        StrSql = StrSql & " ganimpo, porcdecuc, rebaja_esc, retenciones, "
'        StrSql = StrSql & " imp_deter, saldo, pliqnro, pronro, concnro,"
'        StrSql = StrSql & " fecha_pago, ternro)"
'        StrSql = StrSql & " VALUES "
'        StrSql = StrSql & " (" & Gan_Imponible
'        StrSql = StrSql & " ," & factor_esc
'        StrSql = StrSql & " ," & Deducciones
'        StrSql = StrSql & " ," & Ret_Ant
'        StrSql = StrSql & " ," & Monto_calculado
'        StrSql = StrSql & " ," & Retencion
'        StrSql = StrSql & " ," & buliq_periodo!PliqNro
'        StrSql = StrSql & " ," & buliq_proceso!pronro
'        StrSql = StrSql & " ," & Buliq_Concepto(Concepto_Actual).concnro
'        StrSql = StrSql & " ," & ConvFecha(buliq_proceso!profecpago)
'        StrSql = StrSql & " ," & buliq_empleado!ternro
'        StrSql = StrSql & " )"
'        objConn.Execute StrSql, , adExecuteNoRecords
'    Else
'        StrSql = "UPDATE traza_gan SET "
'        StrSql = StrSql & " ganimpo = " & Gan_Imponible
'        StrSql = StrSql & " porcdecuc =" & factor_esc
'        StrSql = StrSql & " rebaja_esc  = " & Deducciones
'        StrSql = StrSql & " retenciones =" & Ret_Ant
'        StrSql = StrSql & " imp_deter =" & Monto_calculado
'        StrSql = StrSql & " saldo =" & Retencion
'        StrSql = StrSql & " WHERE "
'        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
'        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
'        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
'        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
'        objConn.Execute StrSql, , adExecuteNoRecords
'    End If
    
'
'StrSql = "UPDATE traza_gan SET "
'StrSql = StrSql & " ganimpo = " & Gan_Imponible
'StrSql = StrSql & " porcdecuc" & I & "=" & factor_esc
'StrSql = StrSql & " rebaja_esc " & I & "=" & Deducciones
'StrSql = StrSql & " retenciones " & I & "=" & Ret_Ant
'StrSql = StrSql & " imp_deter " & I & "=" & Monto_calculado
'StrSql = StrSql & " saldo " & I & "=" & Retencion
'StrSql = StrSql & " WHERE "
'StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
'StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
'StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
'StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
'StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
'objConn.Execute StrSql, , adExecuteNoRecords

 
    
    
    
    
    ' Si la retenci¢n es negativa, devuelve 0 */
    If Retencion < 0 Then
      Retencion = 0
    End If
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Retenciones anteriores " & Ret_Ant
        Flog.writeline Espacios(Tabulador * 3) & "Monto Calculado " & Monto_calculado
        Flog.writeline Espacios(Tabulador * 3) & "Retencion " & Retencion
    End If
        
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Retenciones anteriores", Ret_Ant)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Monto Calculado ", Monto_calculado)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Retencion ", Retencion)
    End If
   
    
    Bien = True
    
        
    'Retenciones / Devoluciones
    If Retencion <> 0 Then
        Call InsertarFichaRet(buliq_empleado!ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
    End If
    
    ' Grabo todos los items de la liquidacion actual
    I = 1
    Hasta = 50
    Do While I <= Hasta
        If Items_LIQ(I) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!ternro & "," & _
                     ConvFecha(buliq_proceso!profecpago) & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_LIQ(I) & "," & _
                     "0," & _
                     I & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        
        I = I + 1
    Loop

    exito = Bien
    Retencion = Round(Retencion, 0)
    for_ImpuestoUnico = -Retencion
    
      If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Impuesto Unico " & Retencion
    End If
        
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Impuesto Unico ", Retencion)
    End If
    
    Call BorrarTempTable(TTempWF_EscalaUTM)
    
' Cierro todo y libero
  If rs_WF_EscalaUTM.State = adStateOpen Then rs_WF_EscalaUTM.Close
    Set rs_WF_EscalaUTM = Nothing
  If rs_Item.State = adStateOpen Then rs_Item.Close
    Set rs_Item = Nothing
  If rs_valitem.State = adStateOpen Then rs_valitem.Close
    Set rs_valitem = Nothing
  If rs_Desmen.State = adStateOpen Then rs_Desmen.Close
    Set rs_Desmen = Nothing
  If rs_Desliq.State = adStateOpen Then rs_Desliq.Close
    Set rs_Desliq = Nothing
If rs_itemacum.State = adStateOpen Then rs_itemacum.Close
    Set rs_itemacum = Nothing
If rs_itemconc.State = adStateOpen Then rs_itemconc.Close
    Set rs_itemconc = Nothing
If rs_escala.State = adStateOpen Then rs_escala.Close
    Set rs_escala = Nothing
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
    Set rs_wf_tpa = Nothing
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
    Set rs_Acu_Mes = Nothing

    
End Function


Public Sub RecalcImpuestoUnico(ByVal nroPro As Long, ByVal NroPer As Long, ByVal NroCab As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Recalculo de Impuesto Unico.
' Autor      : Martin Ferraro
' Fecha      :
' Ultima Mod.: 16/11/2007
' ---------------------------------------------------------------------------------------------


'Variables
Dim utmAct As Double
Dim nroAcum As Long
Dim montoGratif As Double
Dim mesesTrab As Long
Dim arrInpuniCab(100) As TregImpunicab
Dim I As Long
Dim gratifProp As Double
Dim gratifPropMes As Double
Dim impoSinTope As Double
Dim rtaTribAjustada As Double
Dim existe As Boolean
Dim topeImpGratif As Double
Dim topeMes As Double
Dim difImponible As Double
Dim difImponibleAct As Double
Dim leyesRetenidas As Double
Dim pjeLeyesSociales As Double
Dim leyesSocProp As Double
Dim nuevaRtaImp As Double
Dim desdeEsc As Double
Dim HastaEsc As Double
Dim factorEsc As Double
Dim rebajaEsc As Double
Dim ImpRecalculado As Double
Dim retHist As Double
Dim iuAjustado As Double



'Recorsets
Dim rs_Consulta As New ADODB.Recordset
Dim rs_datos As New ADODB.Recordset


    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "Proceso " & nroPro
        Flog.writeline Espacios(Tabulador * 1) & "Periodo " & NroPer
        Flog.writeline Espacios(Tabulador * 1) & "Cabecera " & NroCab
        Flog.writeline Espacios(Tabulador * 1) & "---------------------------------------------"
        Flog.writeline
    End If
    
    
    '--------------------------------------------------------------------------------------
    'Creo la tabla temporal para la Escala de UTM
    '--------------------------------------------------------------------------------------
    Call CreateTempTable(TTempWF_EscalaUTM)
    
    
    If EsNulo(buliq_periodo!pliqutm) Then utmAct = 0 Else utmAct = buliq_periodo!pliqutm
    
    
    '--------------------------------------------------------------------------------------
    'Busqueda del acumulador imponible (DEBE SER UNICO)
    '--------------------------------------------------------------------------------------
    StrSql = "SELECT acunro, acudesabr"
    StrSql = StrSql & " FROM Acumulador"
    StrSql = StrSql & " WHERE acuimponible = -1"
    StrSql = StrSql & " OR acuimpcont = -1"
    OpenRecordset StrSql, rs_Consulta
    If rs_Consulta.EOF Then
        nroAcum = 0
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "No se encontro acumulador imponible"
        End If
        
'        If HACE_TRAZA Then
'            Call InsertarTraza(NroCab, 0, 0, "No existe ningún Acum. Imponible", 0)
'        End If
        
        Exit Sub
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Acumulador imponible " & rs_Consulta!acuNro & " " & rs_Consulta!acudesabr
        End If
        nroAcum = rs_Consulta!acuNro
        
'        If HACE_TRAZA Then
'            Call InsertarTraza(NroCab, 0, 0, "Nro de Acumulador Imponible " & nroAcum, 0)
'        End If
    End If
    rs_Consulta.Close
    
    
    '--------------------------------------------------------------------------------------
    'Calculo del monto de gratificacion
    '--------------------------------------------------------------------------------------
    montoGratif = 0
    StrSql = "SELECT sum(dlimonto) monto"
    StrSql = StrSql & " FROM detliq"
    StrSql = StrSql & " INNER JOIN concepto ON detliq.concnro = concepto.concnro"
    StrSql = StrSql & " AND concepto.tconnro = 2"
    StrSql = StrSql & " WHERE detliq.cliqnro = " & NroCab
    OpenRecordset StrSql, rs_Consulta
    If Not rs_Consulta.EOF Then
        If Not EsNulo(rs_Consulta!Monto) Then
            montoGratif = rs_Consulta!Monto
        End If
    End If
    rs_Consulta.Close
    
    If montoGratif <= 0 Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Monto de gratif. no (+) " & Round(montoGratif, 0)
        End If
        
'        If HACE_TRAZA Then
'            Call InsertarTraza(NroCab, 0, 0, "Monto de gratif. no (+) ", montoGratif)
'        End If
        
        Exit Sub
    End If
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "Monto de gratif. " & Round(montoGratif, 0)
    End If
'    If HACE_TRAZA Then
'        Call InsertarTraza(NroCab, 0, 0, "Monto de gratif.", montoGratif)
'    End If
    Flog.writeline
    
    '--------------------------------------------------------------------------------------
    'Busco los periodos de recalculo que tiene asociado el proceso
    '--------------------------------------------------------------------------------------
    StrSql = "SELECT impuni_peri.pliqnro, periodo.pliqmes, periodo.pliqanio, periodo.pliqutm"
    StrSql = StrSql & " ,periodo.pliqdesde, periodo.pliqdesc"
    StrSql = StrSql & " FROM impuni_peri"
    StrSql = StrSql & " INNER JOIN periodo ON impuni_peri.pliqnro = periodo.pliqnro"
    StrSql = StrSql & " Where impuni_peri.pronro = " & nroPro
    StrSql = StrSql & " ORDER BY pliqanio, pliqmes"
    OpenRecordset StrSql, rs_Consulta
    If rs_Consulta.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "El proceso no tiene periodos asociados para el recalculo"
        End If
        
        Exit Sub
    Else
        
        mesesTrab = 0
        Do While Not rs_Consulta.EOF
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "Analizando el periodo de recalculo nro " & rs_Consulta!PliqNro & " del " & rs_Consulta!pliqdesc & " - " & rs_Consulta!pliqmes & " " & rs_Consulta!pliqanio
            End If
            
            'Verifico que el empleado trabajo en el mes. Para ello busco que tenga el acumulador imponible encontrado anteriormente
            StrSql = "SELECT proceso.pronro"
            StrSql = StrSql & " FROM Proceso"
            StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro"
            StrSql = StrSql & " AND cabliq.empleado = " & buliq_empleado!ternro
            StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro"
            StrSql = StrSql & " AND acu_liq.acunro = " & nroAcum
            StrSql = StrSql & " WHERE proceso.pliqnro = " & rs_Consulta!PliqNro 'Periodo de recalculo
            OpenRecordset StrSql, rs_datos
            If Not rs_datos.EOF Then
                
                If trabajoEnElMes(rs_Consulta!pliqmes, rs_Consulta!pliqanio) Then
                
                    mesesTrab = mesesTrab + 1
                    arrInpuniCab(mesesTrab).periodo = rs_Consulta!PliqNro
                    arrInpuniCab(mesesTrab).perMes = rs_Consulta!pliqmes
                    arrInpuniCab(mesesTrab).perAnio = rs_Consulta!pliqanio
                    arrInpuniCab(mesesTrab).perUTMHist = IIf(EsNulo(rs_Consulta!pliqutm), 1, rs_Consulta!pliqutm)
                    arrInpuniCab(mesesTrab).periodoDesc = rs_Consulta!pliqdesc
                    
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 2) & rs_Consulta!pliqdesc & " - U.T.M. " & IIf(EsNulo(rs_Consulta!pliqutm), 1, rs_Consulta!pliqutm)
                    End If
'                    If HACE_TRAZA Then
'                        Call InsertarTraza(NroCab, 0, 0, "Monto de gratif.", montoGratif)
'                        Call InsertarTraza(NroCab, 0, 0, rs_Consulta!pliqdesc & " - U.T.M.", IIf(EsNulo(rs_Consulta!pliqutm), 1, rs_Consulta!pliqutm))
'                    End If
                    
                    If rs_Consulta!pliqmes = 12 Then
                        arrInpuniCab(mesesTrab).fechaRet = C_Date("01/01/" & CStr(rs_Consulta!pliqanio + 1)) - 1
                    Else
                        arrInpuniCab(mesesTrab).fechaRet = C_Date("01/" & CStr(rs_Consulta!pliqmes + 1) & "/" & CStr(rs_Consulta!pliqanio)) - 1
                    End If
                    arrInpuniCab(mesesTrab).PeriodoDesde = rs_Consulta!pliqdesde
                
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 2) & rs_Consulta!pliqdesc & " El empleado no trabajo mes completo - No se considera."
                    End If
                End If
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 2) & rs_Consulta!pliqdesc & " El empleado no liquido imponible - No se considera."
                End If
                
            End If
            rs_datos.Close
                    
            rs_Consulta.MoveNext
        Loop
        
    End If
    rs_Consulta.Close
    
    If mesesTrab = 0 Then
        
        If CBool(USA_DEBUG) Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 1) & "En los peíodos elegidos el empleado no trabajo."
        End If
        
'        If HACE_TRAZA Then
'            Call InsertarTraza(NroCab, 0, 0, "En los peíodos elegidos el empleado no trabajo.", 0)
'        End If
        
        Exit Sub
        
    End If
    
    gratifProp = montoGratif / mesesTrab
    
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 1) & "Meses trabajados " & mesesTrab
        Flog.writeline Espacios(Tabulador * 1) & "Gratificación Proporcional " & Round(gratifProp, 0)
        Flog.writeline
    End If
    
'    If HACE_TRAZA Then
'        Call InsertarTraza(NroCab, 0, 0, "Meses Trabajados", mesesTrab)
'        Call InsertarTraza(NroCab, 0, 0, "Gratif. Proporcional", gratifProp)
'    End If
    
    
    '--------------------------------------------------------------------------------------
    'Pre calculo del % del concepto de leyes sociales
    '--------------------------------------------------------------------------------------
    Call liqgra01(6, 35, pjeLeyesSociales)
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "% del concepto de leyes sociales " & pjeLeyesSociales
    End If
     
    
    '--------------------------------------------------------------------------------------
    'Se recorre cada mes trabajado para el calculo
    '--------------------------------------------------------------------------------------
    For I = 1 To mesesTrab
        
        If CBool(USA_DEBUG) Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 1) & "Analizando Meses trabajado " & arrInpuniCab(I).periodoDesc
        End If
        
        arrInpuniCab(I).monto1 = Round(gratifProp, 0)
        
        '--------------------------------------------------------------------------------------
        'Guardo la gratificacion proporcional en el Item 2 de ganancias
        '--------------------------------------------------------------------------------------
        StrSql = "SELECT *"
        StrSql = StrSql & " FROM desliq"
        StrSql = StrSql & " WHERE itenro = 2"
        StrSql = StrSql & " AND empleado = " & buliq_empleado!ternro
        StrSql = StrSql & " AND dlfecha = " & ConvFecha(arrInpuniCab(I).fechaRet)
        StrSql = StrSql & " AND pronro = " & nroPro
        OpenRecordset StrSql, rs_Consulta
        existe = Not rs_Consulta.EOF
        rs_Consulta.Close
        
        'VER CUANDO FALLA Y RE LIQUIDO PORQUE ESTO YA ESTA MODIFICADO Y PERDI EL ULTIMO VALOR
        If existe Then
            StrSql = "UPDATE desliq SET dlmonto = " & Round(gratifProp * arrInpuniCab(I).perUTMHist / utmAct, 0)
            StrSql = StrSql & " WHERE empleado = " & buliq_empleado!ternro
            StrSql = StrSql & " AND itenro = 2"
            StrSql = StrSql & " AND pronro = " & nroPro
            StrSql = StrSql & " AND dlfecha = " & ConvFecha(arrInpuniCab(I).fechaRet)
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "INSERT INTO desliq (empleado,itenro,pronro,dlfecha,dlmonto)"
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & " " & buliq_empleado!ternro
            StrSql = StrSql & " ,2"
            StrSql = StrSql & " ," & nroPro
            StrSql = StrSql & " ," & ConvFecha(arrInpuniCab(I).fechaRet)
            StrSql = StrSql & " ," & Round(gratifProp * arrInpuniCab(I).perUTMHist / utmAct, 0)
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        
        '--------------------------------------------------------------------------------------
        'Gratificacion del mes
        '--------------------------------------------------------------------------------------
        Call calcularImponible(arrInpuniCab(I).periodo, nroAcum, impoSinTope)
        
        gratifPropMes = (gratifProp * arrInpuniCab(I).perUTMHist / utmAct) + impoSinTope
        
        
        '--------------------------------------------------------------------------------------
        'Renta Tributable Ajustada
        '--------------------------------------------------------------------------------------
         rtaTribAjustada = gratifPropMes
         
         
        '--------------------------------------------------------------------------------------
        'Tope imponible de gratificacion (RESULTADO DEL ACUMULADOR TOPEADO)
        '--------------------------------------------------------------------------------------
        topeImpGratif = 0
        StrSql = "SELECT sum(imamonto) monto"
        StrSql = StrSql & " From impmesarg"
        StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND imaanio = " & arrInpuniCab(I).perAnio
        StrSql = StrSql & " AND imames = " & arrInpuniCab(I).perMes
        StrSql = StrSql & " AND acunro = " & nroAcum
        'StrSql = StrSql & " AND (tconnro = 1 OR tconnro = 2)" 'VER QUE LUEGO LO INSERTA (COMO BORRAR??????)
        StrSql = StrSql & " AND (tconnro = 1)" 'VER QUE LUEGO LO INSERTA (COMO BORRAR??????)
        OpenRecordset StrSql, rs_Consulta
        If Not rs_Consulta.EOF Then
            topeImpGratif = IIf(EsNulo(rs_Consulta!Monto), 0, rs_Consulta!Monto)
        End If
        rs_Consulta.Close
        
        
        '--------------------------------------------------------------------------------------
        'Traza de calculados
        '--------------------------------------------------------------------------------------
'        If HACE_TRAZA Then
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Retenido", gratifProp)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Gratif. Mes", gratifPropMes)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Renta Trib. Ajustada", rtaTribAjustada)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Tope Imponible de Gratif", topeImpGratif)
'        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Imponible sin tope " & Round(impoSinTope, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Retenido " & Round(gratifProp, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Gratif. Mes " & Round(gratifPropMes, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Renta Trib. Ajustada " & Round(rtaTribAjustada, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Tope Imponible de Gratif " & Round(topeImpGratif, 0)
        End If
        
        
        '--------------------------------------------------------------------------------------
        'Calculo de la diferencia imponible BUSCA EL TOPE Q APLICO
        '--------------------------------------------------------------------------------------
        topeMes = 0
        StrSql = "SELECT contvalor, ampomax"
        StrSql = StrSql & " FROM AMPO"
        StrSql = StrSql & " WHERE  pliqnro = " & arrInpuniCab(I).periodo
        StrSql = StrSql & " And ampotconnro = 1"
        OpenRecordset StrSql, rs_Consulta
        If Not rs_Consulta.EOF Then
            topeMes = rs_Consulta!contvalor * rs_Consulta!ampomax
        End If
        rs_Consulta.Close
        
        If topeImpGratif >= rtaTribAjustada Then
            'No topeo
            difImponible = 0
        Else
            'Realizo Tope
            If rtaTribAjustada > topeMes Then
                difImponible = topeMes - topeImpGratif
            Else
                difImponible = rtaTribAjustada - topeImpGratif
            End If
        End If
            
            
        '--------------------------------------------------------------------------------------
        'Actualiza la Diferencia Imponible
        '--------------------------------------------------------------------------------------
        difImponibleAct = difImponible / arrInpuniCab(I).perUTMHist * utmAct
        
        
        '--------------------------------------------------------------------------------------
        'Traza de calculados
        '--------------------------------------------------------------------------------------
'        If HACE_TRAZA Then
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Tope Imp", topeMes)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Diferencia Imp", difImponible)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Diferencia Imp. Actual", difImponibleAct)
'        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Tope Imp " & Round(topeMes, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Diferencia Imp " & Round(difImponible, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Diferencia Imp Actual " & Round(difImponibleAct, 0)
        End If
        
        '--------------------------------------------------------------------------------------
        'Modifica el valor del imponible ACTUALIZA CON EN NUEVO TOPE
        '--------------------------------------------------------------------------------------
        StrSql = "SELECT *"
        StrSql = StrSql & " FROM impmesarg"
        StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND imaanio = " & arrInpuniCab(I).perAnio
        StrSql = StrSql & " AND imames = " & arrInpuniCab(I).perMes
        StrSql = StrSql & " AND acunro = " & nroAcum
        StrSql = StrSql & " AND tconnro = 2"
        OpenRecordset StrSql, rs_Consulta
        existe = Not rs_Consulta.EOF
        rs_Consulta.Close
        
        If existe Then
            StrSql = "UPDATE impmesarg SET imamonto = imamonto + " & Round(difImponible, 0)
            StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro
            StrSql = StrSql & " AND imaanio = " & arrInpuniCab(I).perAnio
            StrSql = StrSql & " AND imames = " & arrInpuniCab(I).perMes
            StrSql = StrSql & " AND acunro = " & nroAcum
            StrSql = StrSql & " AND tconnro = 2"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = "INSERT INTO impmesarg (ternro,imaanio,imames,acunro,tconnro,imamonto,imacant)"
            StrSql = StrSql & " VALUES ("
            StrSql = StrSql & " " & buliq_empleado!ternro
            StrSql = StrSql & " ," & arrInpuniCab(I).perAnio
            StrSql = StrSql & " ," & arrInpuniCab(I).perMes
            StrSql = StrSql & " ," & nroAcum
            StrSql = StrSql & " ,2"
            StrSql = StrSql & " ," & Round(difImponible, 0)
            StrSql = StrSql & " ,0"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        
        '--------------------------------------------------------------------------------------
        'Se guarda las Dif. Imponibles
        '--------------------------------------------------------------------------------------
        arrInpuniCab(I).monto2 = Round(difImponibleAct, 0)
        arrInpuniCab(I).monto5 = Round(difImponible, 0)
        
        
        '--------------------------------------------------------------------------------------
        'Leyes retenidas
        '--------------------------------------------------------------------------------------
        leyesRetenidas = calculoLeyRet(arrInpuniCab(I).perMes, arrInpuniCab(I).perAnio, 75)
        
        
        '--------------------------------------------------------------------------------------
        'Nueva Renta Imponible
        '--------------------------------------------------------------------------------------
        leyesSocProp = difImponible * pjeLeyesSociales / 100
        nuevaRtaImp = rtaTribAjustada - (leyesSocProp + leyesRetenidas)
        
        arrInpuniCab(I).monto3 = Round(nuevaRtaImp, 0)
        
        '--------------------------------------------------------------------------------------
        'Traza de calculados
        '--------------------------------------------------------------------------------------
'        If HACE_TRAZA Then
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - % Leyes Soc.", leyesRetenidas)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Leyes Retenidas", leyesRetenidas)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Leyes Soc. Proporcionales", leyesSocProp)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Nueva Renta Imp.", nuevaRtaImp)
'        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Leyes Retenidas " & Round(leyesRetenidas, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Leyes Soc. Proporcionales " & Round(leyesSocProp, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Nueva Renta Imp " & Round(nuevaRtaImp, 0)
        End If
        
        
        'Calcula ganancias con la nueva renta imponible
        '--------------------------------------------------------------------------------------
        'Carga la escala de UTM y Multiplica la escala del Imp. Unico por el valor de UTM del Periodo historico
        '--------------------------------------------------------------------------------------
        Call insertar_wf_escalautm(arrInpuniCab(I).perMes, arrInpuniCab(I).perAnio, arrInpuniCab(I).perUTMHist)
    
        
        '--------------------------------------------------------------------------------------
        'Entra en la escala actualizada por el valor de UTM
        '--------------------------------------------------------------------------------------
        desdeEsc = 0
        HastaEsc = 0
        factorEsc = 0
        rebajaEsc = 0
        
        StrSql = "SELECT * FROM " & TTempWF_EscalaUTM
        StrSql = StrSql & " WHERE desde < " & nuevaRtaImp
        StrSql = StrSql & " AND hasta >= " & nuevaRtaImp
        OpenRecordset StrSql, rs_Consulta
        
        If Not rs_Consulta.EOF Then
            desdeEsc = rs_Consulta!Desde
            HastaEsc = rs_Consulta!Hasta
            factorEsc = rs_Consulta!factor
            rebajaEsc = rs_Consulta!rebaja
        End If
        
        rs_Consulta.Close
        
        ImpRecalculado = (nuevaRtaImp * factorEsc) - rebajaEsc
        
        
        '--------------------------------------------------------------------------------------
        'Traza de calculados
        '--------------------------------------------------------------------------------------
'        If HACE_TRAZA Then
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Escala Desde", desdeEsc)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Escala Hasta", HastaEsc)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Escala Factor", factorEsc)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Escala Rebaja", rebajaEsc)
'        End If
                
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Ingreso a escala con nueva renta = " & Round(nuevaRtaImp, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Escala Desde " & Round(desdeEsc, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Escala Hasta " & Round(HastaEsc, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Escala Factor " & factorEsc
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Escala Rebaja " & Round(rebajaEsc, 0)
        End If
                
                
        '--------------------------------------------------------------------------------------
        'Retenciones ya realizadas
        '--------------------------------------------------------------------------------------
        retHist = 0
                
        StrSql = "SELECT * FROM ficharet "
        StrSql = StrSql & " WHERE empleado =" & buliq_empleado!ternro
        OpenRecordset StrSql, rs_Consulta

        Do While Not rs_Consulta.EOF
            If (Month(rs_Consulta!Fecha) = arrInpuniCab(I).perMes) And (Year(rs_Consulta!Fecha) = arrInpuniCab(I).perAnio) Then
                retHist = retHist + rs_Consulta!importe
            End If
            rs_Consulta.MoveNext
        Loop

        rs_Consulta.Close
        
'        If HACE_TRAZA Then
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Imp. Recalculado", impRecalculado)
'            Call InsertarTraza(NroCab, 0, 0, arrInpuniCab(I).periodoDesc & " - Retenciones", retHist)
'        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Imp. Recalculado " & Round(ImpRecalculado, 0)
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Retenciones " & Round(retHist, 0)
        End If
        
        iuAjustado = ImpRecalculado - retHist
        iuAjustado = iuAjustado / arrInpuniCab(I).perUTMHist * utmAct
        iuAjustado = IIf(iuAjustado < 0, 0, iuAjustado)
        
        arrInpuniCab(I).monto4 = Round(iuAjustado, 0)
        
        
        '--------------------------------------------------------------------------------------
        'Graba retenciones del recalculo
        '--------------------------------------------------------------------------------------
        If arrInpuniCab(I).monto4 <> 0 Then
            Call InsertarFichaRet(buliq_empleado!ternro, arrInpuniCab(I).fechaRet, arrInpuniCab(I).monto4, nroPro)
        End If
        
        
        '--------------------------------------------------------------------------------------
        'Creo inpunicab
        '--------------------------------------------------------------------------------------
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & arrInpuniCab(I).periodoDesc & " - Impuesto Ajustado " & Round(iuAjustado, 0)
            Flog.writeline Espacios(Tabulador * 1) & "Se crea impunicab"
        End If
        StrSql = "INSERT INTO impuni_cab (pliqnro,cliqnro,gratprop,difimponibleact,rentaimpoact,impunicoaju,difimponible)"
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & " " & arrInpuniCab(I).periodo
        StrSql = StrSql & "," & NroCab
        StrSql = StrSql & "," & arrInpuniCab(I).monto1
        StrSql = StrSql & "," & arrInpuniCab(I).monto2
        StrSql = StrSql & "," & arrInpuniCab(I).monto3
        StrSql = StrSql & "," & arrInpuniCab(I).monto4
        StrSql = StrSql & "," & arrInpuniCab(I).monto5
        StrSql = StrSql & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    
    Next I


    Call BorrarTempTable(TTempWF_EscalaUTM)
    
If rs_Consulta.State = adStateOpen Then rs_Consulta.Close
Set rs_Consulta = Nothing
If rs_datos.State = adStateOpen Then rs_datos.Close
Set rs_datos = Nothing

End Sub



Public Function trabajoEnElMes(ByVal mesPer As Integer, ByVal anioPer As Long) As Boolean
' ---------------------------------------------------------------------------------------------
' Descripcion: Verifica si el empleado trabajo en el mes.
'              Si el empleado entro luego del primero del mes o se fue antes
'              del ultimo dia, entonces no cuenta
' Autor      : Martin Ferraro
' Fecha      :
' Ultima Mod.: 16/11/2007
' ---------------------------------------------------------------------------------------------

'Variables
Dim Desde As Date
Dim Hasta As Date
Dim trabajo As Boolean

'Recorsets
Dim rs_Fases As New ADODB.Recordset

trabajo = True

'Busco fechas del periodo
Desde = C_Date("01/" & CStr(mesPer) & "/" & CStr(anioPer))
If mesPer = 12 Then
    Hasta = C_Date("01/01/" & CStr(anioPer + 1)) - 1
Else
    Hasta = C_Date("01/" & CStr(mesPer + 1) & "/" & CStr(anioPer)) - 1
End If

'Busco fases
StrSql = "SELECT altfec, bajfec"
StrSql = StrSql & " FROM fases"
StrSql = StrSql & " WHERE fases.empleado = " & buliq_empleado!ternro
OpenRecordset StrSql, rs_Fases
Do While Not rs_Fases.EOF
    
    'Control Fecha desde
    If Not EsNulo(rs_Fases!altfec) Then
        If ((Month(rs_Fases!altfec) = Month(Desde)) And (Year(rs_Fases!altfec) = Year(Desde)) And (rs_Fases!altfec > Desde)) Then
            trabajo = False
        End If
    End If
    
    'Control Fecha hasta
    If Not EsNulo(rs_Fases!bajfec) Then
        If ((Month(rs_Fases!bajfec) = Month(Hasta)) And (Year(rs_Fases!bajfec) = Year(Hasta)) And (rs_Fases!bajfec < Hasta)) Then
            trabajo = False
        End If
    End If
    
    rs_Fases.MoveNext
Loop
rs_Fases.Close

trabajoEnElMes = trabajo

If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing
End Function



Public Function calculoLeyRet(ByVal mes As Long, ByVal Anio As Long, ByVal acuNro As Long) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Suma un acumulador para los procesos del periodo que no son de recalculo
' Autor      : Martin Ferraro
' Fecha      :
' Ultima Mod.: 16/11/2007
' ---------------------------------------------------------------------------------------------
Dim rs_Acum As New ADODB.Recordset
Dim salidaAux As Double

salidaAux = 0

StrSql = "SELECT sum(abs(acu_liq.almonto)) monto "
StrSql = StrSql & " FROM periodo"
StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro"
StrSql = StrSql & " INNER JOIN tipoproc ON tipoproc.tprocnro = proceso.tprocnro"
StrSql = StrSql & " AND tprocrecalculo <> -1"
StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro"
StrSql = StrSql & " AND cabliq.empleado = " & buliq_empleado!ternro
StrSql = StrSql & " INNER JOIN acu_liq ON acu_liq.cliqnro = cabliq.cliqnro"
StrSql = StrSql & " AND acu_liq.acunro = " & acuNro
StrSql = StrSql & " Where periodo.pliqanio = " & Anio
StrSql = StrSql & " AND periodo.pliqmes = " & mes
OpenRecordset StrSql, rs_Acum

If Not rs_Acum.EOF Then
    salidaAux = IIf(EsNulo(rs_Acum!Monto), 0, rs_Acum!Monto)
End If
rs_Acum.Close

calculoLeyRet = salidaAux

If rs_Acum.State = adStateOpen Then rs_Acum.Close
Set rs_Acum = Nothing
End Function



Public Sub liqgra01(ByVal tipoNro As Long, ByVal parNro As Long, ByRef Monto As Double)
' ---------------------------------------------------------------------------------------------
' Descripcion: Evalua para todos los conceptos de un tipo, la busqueda asociada
'              segun el alcance de resolucion (parametro) y devuelve los
'              resultados.
' Autor      : Martin Ferraro
' Fecha      :
' Ultima Mod.: 16/11/2007
' ---------------------------------------------------------------------------------------------

Dim rs_Conceptos As New ADODB.Recordset
Dim rs_Consulta As New ADODB.Recordset
Dim rs_Alcance As New ADODB.Recordset
Dim rs_ForTpa As New ADODB.Recordset
Dim rs_ConForTpa As New ADODB.Recordset
Dim rs_Programa As New ADODB.Recordset
Dim rs_CftSegun As New ADODB.Recordset

Dim Alcance As Boolean
Dim Origen As Long
Dim ValParam As Double
Dim val As Double
Dim fec As Date
Dim ListaEstr As String
Dim TipoEstr As Long
Dim OK As Boolean
Dim Grupo As Long
Dim tpa_parametro As Long
Dim Resultado As Double
Dim valParamBuscado As Double

    
    'Primero se controla que se que liquide el concepto y luego se busca el parametro %

    Monto = 0
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 2) & "----------------------"
        Flog.writeline Espacios(Tabulador * 2) & "liqgra01"
        Flog.writeline Espacios(Tabulador * 2) & "----------------------"
    End If
    
        
    '----------------------------------------------------------------
    'Busco todos los conceptos del tipo
    '----------------------------------------------------------------
    StrSql = "SELECT concepto.*, formula.*"
    StrSql = StrSql & " FROM concepto"
    StrSql = StrSql & " INNER JOIN con_tp ON con_tp.concnro = concepto.concnro"
    StrSql = StrSql & " AND con_tp.tprocnro = " & buliq_proceso!tprocnro
    StrSql = StrSql & " INNER JOIN formula ON formula.fornro = concepto.fornro"
    StrSql = StrSql & " WHERE concepto.tconnro = " & tipoNro
    StrSql = StrSql & " AND (concepto.concvalid = 0 or ( concdesde <= " & ConvFecha(Fecha_Inicio)
    StrSql = StrSql & " AND conchasta >= " & ConvFecha(Fecha_Fin) & "))"
    StrSql = StrSql & " ORDER BY concepto.concorden"
    OpenRecordset StrSql, rs_Conceptos
    
    Do While Not rs_Conceptos.EOF
        
        If CBool(USA_DEBUG) Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 2) & "Concepto para recalculo: " & rs_Conceptos!Conccod
            Flog.writeline Espacios(Tabulador * 2) & "-------------- Alcance del concepto -----------------------"
        End If
        
        '----------------------------------------------------------------
        '----------------------------------------------------------------
        'Control del alcance del Concepto
        '----------------------------------------------------------------
        '----------------------------------------------------------------
        'Nivel
        'Si es 0 busca el origen con el empleado
        'Si es 1 hay que buscar la estructura
        'Si es 2 global, no debo controlar nada
        Alcance = False
        
        StrSql = "SELECT *"
        StrSql = StrSql & " FROM cge_segun"
        StrSql = StrSql & " WHERE concnro = " & rs_Conceptos!concnro
        StrSql = StrSql & " AND (("
        StrSql = StrSql & " nivel = 0 AND origen = " & buliq_empleado!ternro & ") OR "
        StrSql = StrSql & " (nivel = 1) OR "
        StrSql = StrSql & " (nivel = 2)) "
        StrSql = StrSql & " ORDER BY nivel"
        OpenRecordset StrSql, rs_Alcance
        
        If Not rs_Alcance.EOF Then
            
            Select Case rs_Alcance!Nivel
                Case 0:
                    'Alcance por empleado
                    Alcance = True
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 2) & "Alcance por empleado"
                    End If
                Case 1:
                    'Alcance por estructura
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 2) & "Alcance por estructura"
                    End If

                    TipoEstr = rs_Alcance!Entidad
                    
                    ListaEstr = "0"
                    Do While Not rs_Alcance.EOF
                        ListaEstr = ListaEstr & ", " & rs_Alcance!Origen
                        rs_Alcance.MoveNext
                    Loop
                    
                    'Busco si el empleado posee la estructura
                    StrSql = " SELECT tenro, estrnro FROM his_estructura"
                    StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro & " AND"
                    StrSql = StrSql & " tenro = " & TipoEstr & " AND"
                    StrSql = StrSql & " his_estructura.estrnro IN (" & ListaEstr & ") AND"
                    StrSql = StrSql & " (htetdesde <= " & ConvFecha(Empleado_Fecha_Fin) & ") AND"
                    StrSql = StrSql & " ((" & ConvFecha(Empleado_Fecha_Fin) & " <= htethasta) or (htethasta is null))"
                    OpenRecordset StrSql, rs_Consulta
                        
                    Alcance = Not rs_Consulta.EOF
                    
                    If CBool(USA_DEBUG) Then
                        If Not Alcance Then
                            Flog.writeline Espacios(Tabulador * 2) & "No hay estructuras " & ListaEstr & " de tipo " & TipoEstr
                        Else
                            Flog.writeline Espacios(Tabulador * 2) & "Alcance Estructura Tipo: " & TipoEstr & " Estructura: " & rs_Consulta!estrnro
                        End If
                    End If
                    
                    rs_Consulta.Close
                    
                Case 2:
                    'Alcance global
                    Alcance = True
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 2) & "Alcance Global"
                    End If
                
                Case Else:
                    'No se enconcontro el alcance del concepto
                    Alcance = False
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 2) & "No se enconcontro el alcance del concepto"
                    End If
            End Select
            
        Else
            'No se enconcontro el alcance del concepto
            Alcance = False
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 2) & "No se enconcontro el alcance del concepto"
            End If
        End If
        rs_Alcance.Close
        
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 2) & "Alcance = " & CStr(CBool(Alcance))
            Flog.writeline Espacios(Tabulador * 2) & "-------------- Alcance del concepto -----------------------"
        End If
        
        
        If Not Alcance Then
            GoTo SGT_CONCEPTO
        End If
        
        
        '----------------------------------------------------------------
        '----------------------------------------------------------------
        'Parametros de la formula concepto
        '----------------------------------------------------------------
        '----------------------------------------------------------------
        ValParam = 0
        
        Parametro = 0
        tpa_parametro = 0
        Concepto_Retroactivo = False
        concepto_pliqdesde = 0
        concepto_pliqhasta = 0
        
        
        'Limpiar los wf de pasaje de parametros
        Call LimpiarTempTable(TTempWF_tpa)
        
        StrSql = "SELECT *"
        StrSql = StrSql & " FROM for_tpa"
        StrSql = StrSql & " WHERE for_tpa.fornro = " & rs_Conceptos!fornro
        StrSql = StrSql & " ORDER BY for_tpa.ftorden"
        OpenRecordset StrSql, rs_ForTpa
        
        If rs_ForTpa.EOF Then
            'La formula no tiene parametros o no se encontro la formula
            Flog.writeline Espacios(Tabulador * 3) & "No se encontraron parametros"
            GoTo SGT_CONCEPTO
        End If
        
        'Inicializo el valor del parametro buscado en cero
        valParamBuscado = 0
        
        Do While Not rs_ForTpa.EOF
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Parametro: " & rs_ForTpa!tpanro
            End If
            
            '----------------------------------------------------------------
            'Alcance de los Parametros
            '----------------------------------------------------------------
            StrSql = "SELECT * FROM cft_segun "
            StrSql = StrSql & " WHERE concnro = " & rs_Conceptos!concnro
            StrSql = StrSql & " AND tpanro = " & rs_ForTpa!tpanro
            StrSql = StrSql & " AND fornro = " & rs_ForTpa!fornro & " AND (("
            StrSql = StrSql & " nivel = 0 AND origen = " & buliq_empleado!ternro & ") OR "
            StrSql = StrSql & " (nivel = 1) OR "
            StrSql = StrSql & " (nivel = 2)) "
            StrSql = StrSql & " ORDER BY nivel"
            OpenRecordset StrSql, rs_CftSegun
            
            Alcance = False
            Grupo = 0
            
            If rs_CftSegun.EOF Then
                'No se encontro el alcances
                GoTo SGT_CONCEPTO
            End If
            
            'Excepcion por empleado
            If rs_CftSegun!Nivel = 0 Then
                Alcance = True
                Grupo = 0
            End If
            
            'Analisis de excepcion estructura y alcance global
            Do While Not rs_CftSegun.EOF And Not Alcance

                If rs_CftSegun!Nivel = 1 Then
                    'Por estructura, busco la pertenencia
                    StrSql = "SELECT tenro, estrnro FROM his_estructura"
                    StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro & " AND"
                    StrSql = StrSql & " tenro =" & rs_CftSegun!Entidad & " AND"
                    StrSql = StrSql & " estrnro =" & rs_CftSegun!Origen & " AND"
                    StrSql = StrSql & " (htetdesde <= " & ConvFecha(Empleado_Fecha_Fin) & ") AND"
                    StrSql = StrSql & " ((" & ConvFecha(Empleado_Fecha_Fin) & " <= htethasta) or (htethasta is null))"
                    OpenRecordset StrSql, rs_Consulta

                    If Not rs_Consulta.EOF Then
                        Alcance = True
                        Grupo = rs_CftSegun!Origen
                    End If
                    rs_Consulta.Close
                Else
                    'el alcance es global (nivel 2)
                    Alcance = True
                    Grupo = 0
                End If
                
                'Para dejar posicionado en CftSegun cuando encuentro el alcance
                If Not Alcance Then
                    rs_CftSegun.MoveNext
                End If
            Loop
            
            If Not Alcance Then
                'Parametro sin configurar el alcance
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Parametro no supera el alcance"
                End If
                
'                If HACE_TRAZA Then
'                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Parametro sin configurar el alcance", 0)
'                End If
                
                GoTo SGT_CONCEPTO
            End If
            
            
            '----------------------------------------------------------------
            'Busquedas asociadas a los parametros
            '----------------------------------------------------------------
            If EsNulo(rs_CftSegun!Selecc) Then
                StrSql = "SELECT * FROM con_for_tpa WHERE concnro = " & rs_Conceptos!concnro
                StrSql = StrSql & " AND fornro =" & rs_ForTpa!fornro
                StrSql = StrSql & " AND tpanro =" & rs_ForTpa!tpanro
                StrSql = StrSql & " AND nivel =" & rs_CftSegun!Nivel
            Else
                StrSql = "SELECT * FROM con_for_tpa WHERE concnro = " & rs_Conceptos!concnro
                StrSql = StrSql & " AND fornro =" & rs_ForTpa!fornro
                StrSql = StrSql & " AND tpanro =" & rs_ForTpa!tpanro
                StrSql = StrSql & " AND nivel =" & rs_CftSegun!Nivel
                StrSql = StrSql & " AND selecc ='" & rs_CftSegun!Selecc & "'"
            End If
            OpenRecordset StrSql, rs_ConForTpa
            
            If rs_ConForTpa.EOF Then
                'Parametro sin configurar la busqueda
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Parametro sin configurar la busqueda"
                End If
                
'                If HACE_TRAZA Then
'                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Parametro sin configurar la busqueda", 0)
'                End If
                
                GoTo SGT_CONCEPTO
            End If
            
            
            If EsNulo(rs_ConForTpa!Prognro) Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Nro de busqueda no identificado"
                End If
            
'                If HACE_TRAZA Then
'                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Nro de busqueda no identificado", 0)
'                End If
                ' SIGUIENTE CONCEPTO
                GoTo SGT_CONCEPTO
            End If
            
            
            val = 0
            Valor = 0
            OK = False
            Retroactivo = False
            pliqdesde = 0
            pliqhasta = 0
            
            '----------------------------------------------------------------
            'Busqueda
            '----------------------------------------------------------------
            StrSql = "SELECT * FROM programa WHERE prognro = " & rs_ConForTpa!Prognro
            OpenRecordset StrSql, rs_Programa
            
            If rs_Programa.EOF Then
                'No se encontro la busqueda del parametro
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Nro de busqueda no identificado"
                End If
            
'                If HACE_TRAZA Then
'                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Nro de busqueda no identificado", 0)
'                End If

                GoTo SGT_CONCEPTO
            End If
            
            ' si es automatico y la busqueda esta marcada como que puede usar cache, verificar el cache del empleado
            If CBool(rs_ConForTpa!cftauto) Then
                '----------------------------------------------------------------
                'Busqueda Automatica
                '----------------------------------------------------------------
                
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Busqueda Automatica: " & rs_Programa!Prognro
                End If
                
                'Busqueda generada
                If CBool(rs_Programa!Progarchest) Then
                    If CBool(rs_Programa!Progcache) Then
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "busca en cache"
                        End If
                       
                       If objCache.EsSimboloDefinido(CStr(rs_Programa!Prognro)) Then
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & " está en cache "
                            End If
                            
                            val = objCache.Valor(CStr(rs_Programa!Prognro))
                            OK = True
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & " NO está en cache. Ejecuto la busqueda"
                            End If
                            
                            Call EjecutarBusqueda(rs_Programa!Tprognro, rs_Conceptos!concnro, rs_ConForTpa!Prognro, val, fec, OK)
                            
                            If Not OK Then
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Error en búsqueda de parametro"
                                End If
                                
'                                If HACE_TRAZA Then
'                                    Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Error en búsqueda de parametro", 0)
'                                End If
                                
                                ' SIGUIENTE CONCEPTO
                                GoTo SGT_CONCEPTO
                            Else
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Resultado de la busqueda: " & val
                                End If
                            End If
                            'VERRRRRRRRRRRRRRRr
                            ' insertar en el cache del empleado
                            'Call objCache.Insertar_Simbolo(CStr(Arr_Programa(Arr_con_for_tpa(Indice_Actual).Prognro).Prognro), val)
                        End If
                    Else
                        ' busqueda automatica, primera vez
                        Call EjecutarBusqueda(rs_Programa!Tprognro, rs_Conceptos!concnro, rs_ConForTpa!Prognro, val, fec, OK)
                        
                        If Not OK Then
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Error en búsqueda de parametro"
                            End If
                            
'                            If HACE_TRAZA Then
'                                Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Error en búsqueda de parametro", 0)
'                            End If
                        
                            ' SIGUIENTE CONCEPTO
                            GoTo SGT_CONCEPTO
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Resultado de la busqueda: " & val
                            End If
                        End If
                    End If
                    
                'Busqueda no generada
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & " busqueda no generada "
                    End If
                    
'                    If HACE_TRAZA Then
'                        Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "Búsqueda no está generada", 0)
'                    End If
                    
                End If
                
            Else
                '----------------------------------------------------------------
                'Busqueda No Automatica
                '----------------------------------------------------------------
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Busqueda NO Automatica: " & rs_ConForTpa!Prognro
                End If
                
                ' Busqueda no Automatica, Novedades buscar
                If Not NovedadesHist Then
                    Call Bus_NovGegi(rs_ConForTpa!Prognro, rs_Conceptos!concnro, rs_ForTpa!tpanro, Empleado_Fecha_Inicio, Empleado_Fecha_Fin, Grupo, OK, val)
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "en historico "
                    End If

                    Call Bus_NovGegiHis(rs_ConForTpa!Prognro, rs_Conceptos!concnro, rs_ForTpa!tpanro, Empleado_Fecha_Inicio, Empleado_Fecha_Fin, Grupo, OK, val)
                End If
                
                If Not OK Then
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "No se encontró la Novedad"
                    End If
                    
'                    If HACE_TRAZA Then
'                        Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, Arr_For_Tpa(Indice_Actual_For_Tpa).tpanro, "No se encontró la Novedad", 0)
'                    End If
                    
                    ' SIGUIENTE CONCEPTO
                    GoTo SGT_CONCEPTO
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "Resultado de la busqueda: " & val
                    End If
                End If
            
            End If 'Buqueda
            
            
            rs_Programa.Close


            '----------------------------------------------------------------
            'Si se obtuvo el parametro satisfactoriamente
            '----------------------------------------------------------------
            If OK Then
                'parametro Correcto. Inserto en el wf_tpa
                Call insertar_wf_tpa(rs_ForTpa!tpanro, rs_ForTpa!ftorden, rs_Conceptos!concabr, val, fec)
                
                If rs_ForTpa!ftimprime Then
                    ' Guarda el parametro imprimible
                    Parametro = val
                    tpa_parametro = rs_ForTpa!tpanro
                End If
                
                'Si es el parametro buscado (es uno por concepto)
                If parNro = rs_ForTpa!tpanro Then
                    valParamBuscado = val
                End If

            Else
                'No se pudo obtener el parametro.
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se obtuvo el parametro"
                End If
                
                GoTo SGT_CONCEPTO
            End If

            rs_ForTpa.MoveNext
        Loop
        
        rs_ForTpa.Close
        
        '----------------------------------------------------------------
        'EJECUCION DE LA FORMULA DEL CONCEPTO
        '----------------------------------------------------------------
        exito = False
        Resultado = 0
        
        ' setear todos los parametros en wf_tpa en la tabla de simbolos
        Call CargarTablaParametros
        
        ' reviso si la formula es externa o interna
        If rs_Conceptos!Fortipo = 3 Then 'Configurable
            ' Evalua la expresion de la formula
            Resultado = CDbl(eval.Evaluate(Trim(rs_Conceptos!Forexpresion), exito, True))
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Formula: " & rs_Conceptos!fornro
            End If
            
        Else 'es una formula codificada en vb (un procedimiento que la resuleve)
            'el tema es como resulevo a que formula llamar
            If rs_Conceptos!Fortipo = 2 Then 'No configurable
                Resultado = EjecutarFormulaNoConfigurable(Trim(rs_Conceptos!Forprog))
            Else ' de sistema
                Resultado = EjecutarFormulaDeSistema(Trim(rs_Conceptos!Forprog))
            End If
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Formula: " & rs_Conceptos!Forprog
            End If
        End If
        
        '----------------------------------------------------------------
        'Control de formula incorrecta
        '----------------------------------------------------------------
        If (Not exito) Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "La formula no devolvio un OK"
            End If
            
'            If HACE_TRAZA Then
'                Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, 9999, "La formula no devolvio un OK", 0)
'            End If
            
            ' SIGUIENTE CONCEPTO
            GoTo SGT_CONCEPTO
        End If
        
        
        '----------------------------------------------------------------
        'Control de resultado = 0
        '----------------------------------------------------------------
        ' Redondeo por presicion para trabajar con dos decimales
        ' Hay que sacarlo del concepto
        Resultado = Round(Resultado, rs_Conceptos!Conccantdec)
        
        If (Resultado = 0) Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "La formula devolvio CERO"
            End If
        
'            If HACE_TRAZA Then
'                Call InsertarTraza(NroCab, Arr_conceptos(Concepto_Actual).concnro, 9999, "La formula devolvio CERO", 0)
'            End If
            ' SIGUIENTE CONCEPTO
            GoTo SGT_CONCEPTO
        End If
        
        '----------------------------------------------------------------
        'Si liquido correctamente acumulo el valor del parametro buscado para todos los conceptos
        '----------------------------------------------------------------
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 2) & "El Concepto contribuye con " & valParamBuscado
        End If

        ' inserto en el cache de detliq
        ' El monto
        Call objCache_detliq_Monto.Insertar_Simbolo(CStr(rs_Conceptos!concnro), Resultado)
        ' La cantidad
        'Call objCache_detliq_Cantidad.Insertar_Simbolo(CStr(rs_Conceptos!concnro), Parametro)


        Monto = Monto + valParamBuscado
        
                
SGT_CONCEPTO: rs_Conceptos.MoveNext
        
        
    Loop 'Fin cada concepto
    rs_Conceptos.Close
    
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 2) & "Valor resultado: " & Monto
        Flog.writeline
    End If
    
If rs_Conceptos.State = adStateOpen Then rs_Conceptos.Close
Set rs_Conceptos = Nothing
If rs_Consulta.State = adStateOpen Then rs_Consulta.Close
Set rs_Consulta = Nothing
If rs_ForTpa.State = adStateOpen Then rs_ForTpa.Close
Set rs_ForTpa = Nothing
If rs_ConForTpa.State = adStateOpen Then rs_ConForTpa.Close
Set rs_ConForTpa = Nothing
If rs_Programa.State = adStateOpen Then rs_Programa.Close
Set rs_Programa = Nothing
If rs_Alcance.State = adStateOpen Then rs_Alcance.Close
Set rs_Alcance = Nothing
If rs_CftSegun.State = adStateOpen Then rs_CftSegun.Close
Set rs_CftSegun = Nothing

End Sub


Public Sub calcularImponible(ByVal periHist As Long, ByVal nroAcum As Long, ByRef impoSinTope As Double)
'Busca como aportaron los conceptos al acumulador en el periodo de recalculo (imponible sin topear).
'OJO si se cambia la config del acum!!!!
Dim rs_consult As New ADODB.Recordset

    
    impoSinTope = 0
    
    StrSql = "SELECT sum(detliq.dlimonto) monto"
    StrSql = StrSql & " FROM Proceso"
    StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro"
    StrSql = StrSql & " AND cabliq.empleado = " & buliq_empleado!ternro
    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro"
    StrSql = StrSql & " INNER JOIN con_acum ON con_acum.concnro = detliq.concnro"
    StrSql = StrSql & " AND con_acum.acunro = " & nroAcum
    StrSql = StrSql & " WHERE proceso.pliqnro = " & periHist
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        If Not EsNulo(rs_consult!Monto) Then
            impoSinTope = rs_consult!Monto
        End If
    End If
    rs_consult.Close
    
    '------------------------------------------------------------------------
    'Suma lo que haya acumulado de otros c lculos de Gratificaci¢n anteriores
    '------------------------------------------------------------------------
    StrSql = "SELECT sum(impmesarg.imamonto) monto"
    StrSql = StrSql & " From impmesarg"
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqmes = impmesarg.imames"
    StrSql = StrSql & " AND periodo.pliqanio = impmesarg.imaanio"
    StrSql = StrSql & " AND impmesarg.ternro = " & buliq_empleado!ternro
    StrSql = StrSql & " AND impmesarg.acunro = " & nroAcum
    StrSql = StrSql & " WHERE periodo.pliqnro = " & periHist
    StrSql = StrSql & " AND impmesarg.tconnro = 2"
    OpenRecordset StrSql, rs_consult
    If Not rs_consult.EOF Then
        If Not EsNulo(rs_consult!Monto) Then
            impoSinTope = impoSinTope + rs_consult!Monto
        End If
    End If
    rs_consult.Close
    
    impoSinTope = Round(impoSinTope, 0)
    
If rs_consult.State = adStateOpen Then rs_consult.Close
Set rs_consult = Nothing

End Sub


Public Function for_RecalcConcepto(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Recalculo de conceptos para el impuesto unico
' Autor      : Martin
' Fecha      :
' Ultima Mod.: 20/01/2009
' ---------------------------------------------------------------------------------------------
Const c_ConcCodRec = 109
Const c_Bono = 104
Const c_Tope = 1002
Const c_AcuImp = 98
Const c_LiqImp = 1005

'Parametros
Dim ConcCodRec As Long
Dim Bono As Double
Dim Tope As Double
Dim AcuImp As Long
Dim LiqImp As Integer

Dim EncConcCodRec As Boolean
Dim EncBono As Boolean
Dim EncTope As Boolean
Dim EncAcuImp As Boolean
Dim EncLiqImp As Boolean

Dim CantPerRec As Long
Dim BonoPeriodo As Double
Dim ImpoMesHist As Double
Dim UFHist As Double
Dim MontoConcRecalc As Double
Dim DifUFHist As Double
Dim Porc As Double
Dim DifUFHistPorc As Double
Dim TotalDifUFHistPorc As Double

Dim rs_Periodos As New ADODB.Recordset
Dim rs_consult As New ADODB.Recordset

    LiqImp = 1
        
    EncConcCodRec = False
    EncBono = False
    EncTope = False
    EncAcuImp = False
    Bien = False

    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_consult
    
    Do While Not rs_consult.EOF
    
        Select Case rs_consult!tipoparam
            Case c_ConcCodRec:
                ConcCodRec = rs_consult!Valor
                EncConcCodRec = True
            Case c_Bono:
                Bono = rs_consult!Valor
                EncBono = True
            Case c_Tope:
                Tope = rs_consult!Valor
                EncTope = True
            Case c_AcuImp:
                AcuImp = rs_consult!Valor
                EncAcuImp = True
            Case c_LiqImp:
                LiqImp = rs_consult!Valor
                EncLiqImp = True
            Case Else
        End Select
        
        rs_consult.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not EncConcCodRec Or Not EncBono Or Not EncTope Or Not EncAcuImp Then
        Exit Function
    End If

    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "---------Parametros-----------------------------"
        Flog.writeline Espacios(Tabulador * 3) & "Concepto a Recalcular: " & ConcCodRec
        Flog.writeline Espacios(Tabulador * 3) & "Valor Bono: " & Bono
        Flog.writeline Espacios(Tabulador * 3) & "Tope UF: " & Tope
        Flog.writeline Espacios(Tabulador * 3) & "Acumulador Imponible: " & AcuImp
        Flog.writeline Espacios(Tabulador * 3) & "Liq Impuesto: " & LiqImp
        Flog.writeline
    End If
    
    'Busco la cantidad de periodos de recalculo del proceso
    CantPerRec = 0
    StrSql = "SELECT periodo.* FROM impuni_peri "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = impuni_peri.pliqnro"
    StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
    StrSql = StrSql & " ORDER BY periodo.pliqdesde"
    OpenRecordset StrSql, rs_Periodos
    If Not rs_Periodos.EOF Then
        CantPerRec = rs_Periodos.RecordCount
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Cantidad de periodos de Recalculo: " & CantPerRec
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "ERROR. El proceso no tiene periodos de recalculo asociados."
        End If
        for_RecalcConcepto = 0
        Exit Function
    End If
    
    If CantPerRec = 0 Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "ERROR. El proceso no tiene periodos de recalculo asociados."
        End If
        for_RecalcConcepto = 0
        Exit Function
    End If
    
    'Calculo el bono de cada periodo
    BonoPeriodo = Bono / CantPerRec
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 3) & "Bono por Periodo: " & BonoPeriodo
    End If


    Do While Not rs_Periodos.EOF
        
        If CBool(USA_DEBUG) Then
            Flog.writeline
            Flog.writeline Espacios(Tabulador * 3) & "Procesando Periodo de Recalculo " & rs_Periodos!PliqNro & " - " & rs_Periodos!pliqdesc
            Flog.writeline Espacios(Tabulador * 3) & "_______________________________________________________________________________________________"
        End If

        
        'Busco el imponible del mes
        ImpoMesHist = 0
        StrSql = "SELECT ammonto, amcant"
        StrSql = StrSql & " FROM acu_mes"
        StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND acunro = " & AcuImp
        StrSql = StrSql & " AND  amanio = " & rs_Periodos!pliqanio
        StrSql = StrSql & " AND ammes = " & rs_Periodos!pliqmes
        OpenRecordset StrSql, rs_consult
        If Not rs_consult.EOF Then
            ImpoMesHist = rs_consult!ammonto
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Imponible: " & ImpoMesHist
            End If
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontro Imponible del periodo."
            End If
            GoTo SgtPer
        End If
        
        If ImpoMesHist = 0 Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontro Imponible del periodo."
            End If
            GoTo SgtPer
        End If
        
        'Busco el UFHist del mes
        UFHist = 0
        StrSql = "SELECT valor FROM ampo WHERE ampofecha <= " & ConvFecha(rs_Periodos!pliqhasta)
        StrSql = StrSql & " ORDER BY ampofecha DESC"
        OpenRecordset StrSql, rs_consult
        If Not rs_consult.EOF Then
            UFHist = rs_consult!Valor
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "UF: " & UFHist
            End If
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontro UF del periodo."
            End If
            
        End If
        
        If ImpoMesHist < (Tope * UFHist) Then
            
            'Busco el monto del concepto a recalcular en el periodo
            MontoConcRecalc = 0
            StrSql = "SELECT SUM(dlimonto) monto, concepto.concnro"
            StrSql = StrSql & " FROM detliq"
            StrSql = StrSql & " INNER JOIN cabliq ON cabliq.cliqnro = detliq.cliqnro"
            StrSql = StrSql & " AND cabliq.empleado = " & buliq_empleado!ternro
            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
            StrSql = StrSql & " AND proceso.pliqnro = " & rs_Periodos!PliqNro
            StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
            StrSql = StrSql & " AND CAST(concepto.conccod as int) = " & ConcCodRec
            StrSql = StrSql & " GROUP BY concepto.concnro"
            OpenRecordset StrSql, rs_consult
            If Not rs_consult.EOF Then
                If Not EsNulo(rs_consult!Monto) Then
                    MontoConcRecalc = rs_consult!Monto
                End If
            End If
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Monto del concepto en el periodo: " & MontoConcRecalc
            End If
            
            If MontoConcRecalc = 0 Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Paso al siguiente periodo."
                End If
                GoTo SgtPer
            End If
            
            'Calculo porcentaje historico que aplique al concepto
            Porc = MontoConcRecalc / ImpoMesHist * 100
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Porcentaje Historico aplicado al concepto: " & Round(Porc, 2)
            End If
            
            'Calculo lo que me falta para llegar al imponible
            If (ImpoMesHist + BonoPeriodo) < (Tope * UFHist) Then
                DifUFHist = BonoPeriodo
            Else
                DifUFHist = (Tope * UFHist) - ImpoMesHist
            End If
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Diferencia con imponible: " & DifUFHist
            End If
            
            'Aplico el porcentaje a la diferencia
            DifUFHistPorc = DifUFHist * Round(Porc, 2) / 100
            
            If LiqImp = 1 Then
                StrSql = "INSERT INTO impuni_cab (pliqnro,cliqnro,gratprop,difimponibleact,rentaimpoact,impunicoaju,difimponible,concnro)"
                StrSql = StrSql & " VALUES ("
                StrSql = StrSql & " " & rs_Periodos!PliqNro
                StrSql = StrSql & "," & buliq_cabliq!cliqnro
                StrSql = StrSql & "," & Abs(BonoPeriodo)
                StrSql = StrSql & "," & Abs(DifUFHist)
                StrSql = StrSql & "," & Abs(MontoConcRecalc)
                StrSql = StrSql & "," & Abs(MontoConcRecalc + DifUFHistPorc)
                StrSql = StrSql & "," & Abs(DifUFHistPorc)
                StrSql = StrSql & "," & Buliq_Concepto(Concepto_Actual).concnro
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            'Acumulo
            TotalDifUFHistPorc = TotalDifUFHistPorc + DifUFHistPorc
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Se acumula: " & DifUFHistPorc
            End If
            
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "El imponible del mes es mayor o igual al tope por UF del mes. No acumula."
            End If
        End If
        
SgtPer: rs_Periodos.MoveNext

    Loop
    
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 3) & "Total acumulado: " & TotalDifUFHistPorc
    End If
    
    Bien = True
    exito = Bien
    for_RecalcConcepto = Abs(TotalDifUFHistPorc) * -1
    

'Libero Recordset
If rs_Periodos.State = adStateOpen Then rs_Periodos.Close
Set rs_Periodos = Nothing
If rs_consult.State = adStateOpen Then rs_consult.Close
Set rs_consult = Nothing

End Function


Public Function for_RecalcImpuestoUnico(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Recalculo del impuesto unico
' Autor      : Martin
' Fecha      :
' Ultima Mod.: 20/01/2009
' ---------------------------------------------------------------------------------------------

Dim CantPerRec As Long
Dim BonoPeriodo As Double
Dim ImpoMesHist As Double
Dim ImpPagado As Double
Dim ImpoHistBono As Double
Dim BonoAjustado As Double
Dim UTMHist As Double
Dim DeduccHist As Double
Dim NuevoImpoMesHist As Double
Dim desdeEsc As Double
Dim HastaEsc As Double
Dim factorEsc As Double
Dim rebajaEsc As Double
Dim ImpRecalculado As Double
Dim ImpRecalculadoAcum As Double

Dim rs_Periodos As New ADODB.Recordset
Dim rs_consult As New ADODB.Recordset

Const c_Bono = 104
Const c_AcuImp = 98

'Parametros
Dim Bono As Double
Dim AcuImp As Long

Dim EncBono As Boolean
Dim EncAcuImp As Boolean


    Bien = False
    ImpRecalculadoAcum = 0
    
    EncBono = False
    EncAcuImp = False
   

    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_consult
    
    Do While Not rs_consult.EOF
    
        Select Case rs_consult!tipoparam
            Case c_Bono:
                Bono = rs_consult!Valor
                EncBono = True
            Case c_AcuImp:
                AcuImp = rs_consult!Valor
                EncAcuImp = True
            Case Else
        End Select
        
        rs_consult.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not EncBono Or Not EncAcuImp Then
        Exit Function
    End If

    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "---------Parametros-----------------------------"
        Flog.writeline Espacios(Tabulador * 3) & "Valor Bono: " & Bono
        Flog.writeline Espacios(Tabulador * 3) & "Acumulador Imponible: " & AcuImp
        Flog.writeline
    End If

    'Creo la tabla temporal para la Escala de UTM
    Call CreateTempTable(TTempWF_EscalaUTM)
    
    
    'Busco la cantidad de periodos de recalculo del proceso
    CantPerRec = 0
    StrSql = "SELECT periodo.* FROM impuni_peri "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = impuni_peri.pliqnro"
    StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
    OpenRecordset StrSql, rs_Periodos
    If Not rs_Periodos.EOF Then
        CantPerRec = rs_Periodos.RecordCount
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Cantidad de periodos de Recalculo: " & CantPerRec
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "ERROR. El proceso no tiene periodos de recalculo asociados."
        End If
        for_RecalcImpuestoUnico = 0
        Exit Function
    End If
    
    If CantPerRec = 0 Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "ERROR. El proceso no tiene periodos de recalculo asociados."
        End If
        for_RecalcImpuestoUnico = 0
        Exit Function
    End If
    
    'Calculo el bono de cada periodo
    BonoPeriodo = Bono / CantPerRec
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 3) & "Bono por Periodo: " & FormatNumber(BonoPeriodo, 2)
    End If
    
    
    Do While Not rs_Periodos.EOF
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Procesando Periodo de Recalculo " & rs_Periodos!PliqNro & " - " & rs_Periodos!pliqdesc
            Flog.writeline Espacios(Tabulador * 3) & "_______________________________________________________________________________________________"
        End If
        
        'Busco el UTM del mes
        UTMHist = rs_Periodos!pliqutm
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "UTM Mes: " & FormatNumber(UTMHist, 2)
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "UTM Actual: " & FormatNumber(buliq_periodo!pliqutm, 2)
        End If
        
        BonoAjustado = (BonoPeriodo / buliq_periodo!pliqutm) * UTMHist
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Bono Ajustado UTM: " & FormatNumber(BonoAjustado, 2)
        End If
        
        'Busco el imponible del mes
        ImpoMesHist = 0
        StrSql = "SELECT ammonto, amcant"
        StrSql = StrSql & " FROM acu_mes"
        StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND acunro = " & AcuImp
        StrSql = StrSql & " AND  amanio = " & rs_Periodos!pliqanio
        StrSql = StrSql & " AND ammes = " & rs_Periodos!pliqmes
        OpenRecordset StrSql, rs_consult
        
        If Not rs_consult.EOF Then
            ImpoMesHist = rs_consult!ammonto
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Imponible: " & FormatNumber(rs_consult!ammonto, 2)
            End If
        Else
            ImpoMesHist = 0
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontro Imponible del periodo."
            End If
        End If
        
        ImpoHistBono = ImpoMesHist + BonoAjustado
        
        
        'Buscar en impunicab con el ajuste
        DeduccHist = 0
        'StrSql = "SELECT SUM(impunicoaju) monto FROM impuni_cab"
        StrSql = "SELECT SUM(difimponible) monto FROM impuni_cab"
        StrSql = StrSql & " WHERE pliqnro = " & rs_Periodos!PliqNro
        StrSql = StrSql & " AND cliqnro = " & buliq_cabliq!cliqnro
        OpenRecordset StrSql, rs_consult
        If Not rs_consult.EOF Then
            If Not EsNulo(rs_consult!Monto) Then
                DeduccHist = rs_consult!Monto
            End If
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Deduccion Hist Ajus: " & FormatNumber(DeduccHist, 2)
        End If
        
        
        NuevoImpoMesHist = ImpoHistBono - DeduccHist
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Nuevo Imponible: " & FormatNumber(NuevoImpoMesHist, 2)
        End If
        
        
        'Carga la escala de UTM y Multiplica la escala del Imp. Unico por el valor de UTM del Periodo historico
        Call insertar_wf_escalautm(rs_Periodos!pliqmes, rs_Periodos!pliqanio, UTMHist)
    
        
        'Entra en la escala actualizada por el valor de UTM
        desdeEsc = 0
        HastaEsc = 0
        factorEsc = 0
        rebajaEsc = 0
        ImpRecalculado = 0
        
        StrSql = "SELECT * FROM " & TTempWF_EscalaUTM
        StrSql = StrSql & " WHERE desde < " & Abs(NuevoImpoMesHist)
        StrSql = StrSql & " AND hasta >= " & Abs(NuevoImpoMesHist)
        OpenRecordset StrSql, rs_consult
        
        If Not rs_consult.EOF Then
            desdeEsc = rs_consult!Desde
            HastaEsc = rs_consult!Hasta
            factorEsc = rs_consult!factor
            rebajaEsc = rs_consult!rebaja
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Entra en escala con: " & Abs(NuevoImpoMesHist)
                Flog.writeline Espacios(Tabulador * 4) & "Desde: " & desdeEsc
                Flog.writeline Espacios(Tabulador * 4) & "Hasta: " & HastaEsc
                Flog.writeline Espacios(Tabulador * 4) & "Factor: " & FormatNumber(factorEsc, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Rebaja: " & FormatNumber(rebajaEsc, 2)
            End If
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encuentra escala con " & FormatNumber(NuevoImpoMesHist, 2)
            End If
        End If
        
        
        ImpRecalculado = (NuevoImpoMesHist * factorEsc) - rebajaEsc
        
        'Restar al impuesto lo ya liquidado por impuesto unico
        'Entrar nuevamente a escala con el Imponible historico
        
        StrSql = "SELECT * FROM " & TTempWF_EscalaUTM
        StrSql = StrSql & " WHERE desde < " & Abs(ImpoMesHist)
        StrSql = StrSql & " AND hasta >= " & Abs(ImpoMesHist)
        OpenRecordset StrSql, rs_consult
        
        If Not rs_consult.EOF Then
            desdeEsc = rs_consult!Desde
            HastaEsc = rs_consult!Hasta
            factorEsc = rs_consult!factor
            rebajaEsc = rs_consult!rebaja
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Entra en escala Nuevamente con: " & FormatNumber(Abs(ImpoMesHist), 2)
                Flog.writeline Espacios(Tabulador * 4) & "Desde: " & desdeEsc
                Flog.writeline Espacios(Tabulador * 4) & "Hasta: " & HastaEsc
                Flog.writeline Espacios(Tabulador * 4) & "Factor: " & FormatNumber(factorEsc, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Rebaja: " & FormatNumber(rebajaEsc, 2)
            End If
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encuentra escala con imp pagado" & FormatNumber(ImpoMesHist, 2)
            End If
        End If
        
        
        ImpPagado = (ImpoMesHist * factorEsc) - rebajaEsc
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Impuesto Pagado: " & FormatNumber(ImpPagado, 2)
        End If
        
        'Resto el impuesto recalculado lo ya pagado
        ImpRecalculado = ImpRecalculado - ImpPagado
                        
        'Acumulo en UTM
        ImpRecalculadoAcum = ImpRecalculadoAcum + (ImpRecalculado / UTMHist)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Acumulo en UTM: " & FormatNumber((ImpRecalculado / UTMHist), 2)
        End If
        
        rs_Periodos.MoveNext
    Loop
    
    exito = True
    
    
    If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "UTM Actual: " & FormatNumber(buliq_periodo!pliqutm, 2)
    End If
        
    for_RecalcImpuestoUnico = Abs(ImpRecalculadoAcum * buliq_periodo!pliqutm) * -1
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Total: " & FormatNumber((ImpRecalculadoAcum * buliq_periodo!pliqutm), 2)
    End If
    
    

'Libero Recordset
If rs_Periodos.State = adStateOpen Then rs_Periodos.Close
Set rs_Periodos = Nothing
If rs_consult.State = adStateOpen Then rs_consult.Close
Set rs_consult = Nothing

End Function

