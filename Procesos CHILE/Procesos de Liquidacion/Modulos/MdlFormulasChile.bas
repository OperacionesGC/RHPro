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
 'Call CreateTempTable(TTempWF_EscalaUTM)

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
    StrSql = "SELECT * FROM ficharet " & NOLOCK & _
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
    
    'Call BorrarTempTable(TTempWF_EscalaUTM)
    
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
Dim OtrasReliq As Double
Dim ImpoMesHistReliq As Double

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
        ImpoMesHistReliq = 0
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
        
        'Buscar impuni_cab del mismo periodo por si tiene otras reliquidaciones para sumar al imponible
            OtrasReliq = 0
            StrSql = "SELECT SUM(difimponible) Monto FROM impuni_cab "
            StrSql = StrSql & " WHERE impuni_cab.pliqnro = " & rs_Periodos!PliqNro
            StrSql = StrSql & " AND impuni_cab.concnro = " & Buliq_Concepto(Concepto_Actual).concnro
            StrSql = StrSql & " AND impuni_cab.aux1 = " & buliq_empleado!ternro
            StrSql = StrSql & " AND impuni_cab.cliqnro <> " & buliq_cabliq!cliqnro
            OpenRecordset StrSql, rs_consult
            If Not rs_consult.EOF Then
                If Not EsNulo(rs_consult!Monto) Then
                    OtrasReliq = rs_consult!Monto
                End If
            End If
            
            ImpoMesHistReliq = ImpoMesHist + OtrasReliq
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Otras Reliquidaciones en el periodo: " & OtrasReliq
                Flog.writeline Espacios(Tabulador * 4) & "Imponible + Reliq: " & ImpoMesHistReliq
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
        
        'Si ya consumi todo el tope salgo
        If ImpoMesHistReliq < (Tope * UFHist) Then
            
            'Busco el monto del concepto a recalcular en el periodo
            MontoConcRecalc = 0
'            StrSql = "SELECT SUM(dlimonto) monto, concepto.concnro"
'            StrSql = StrSql & " FROM detliq"
'            StrSql = StrSql & " INNER JOIN cabliq ON cabliq.cliqnro = detliq.cliqnro"
'            StrSql = StrSql & " AND cabliq.empleado = " & buliq_empleado!ternro
'            StrSql = StrSql & " INNER JOIN proceso ON proceso.pronro = cabliq.pronro"
'            StrSql = StrSql & " AND proceso.pliqnro = " & rs_Periodos!PliqNro
'            StrSql = StrSql & " INNER JOIN concepto ON concepto.concnro = detliq.concnro"
'            StrSql = StrSql & " AND CAST(concepto.conccod as int) = " & ConcCodRec
'            StrSql = StrSql & " GROUP BY concepto.concnro"

             'Busco el monto del concepto a recalcular en el periodo en un acumulador para poder migrarlo
            StrSql = "SELECT SUM(ammonto) Monto"
            StrSql = StrSql & " FROM acu_mes"
            StrSql = StrSql & " WHERE acu_mes.ternro = " & buliq_empleado!ternro
            StrSql = StrSql & " AND acu_mes.acunro = " & ConcCodRec
            StrSql = StrSql & " AND acu_mes.amanio = " & rs_Periodos!pliqanio
            StrSql = StrSql & " AND acu_mes.ammes = " & rs_Periodos!pliqmes
            OpenRecordset StrSql, rs_consult
            If Not rs_consult.EOF Then
                If Not EsNulo(rs_consult!Monto) Then
                    MontoConcRecalc = rs_consult!Monto
                End If
            End If
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Monto del acum/cpto en el periodo: " & MontoConcRecalc
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
            If (ImpoMesHistReliq + BonoPeriodo) < (Tope * UFHist) Then
                DifUFHist = BonoPeriodo
            Else
                DifUFHist = (Tope * UFHist) - ImpoMesHistReliq
            End If
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Diferencia con imponible: " & DifUFHist
            End If
            
            'Aplico el porcentaje a la diferencia
            DifUFHistPorc = DifUFHist * Round(Porc, 2) / 100
            
            If LiqImp = 1 Then
                StrSql = "INSERT INTO impuni_cab (pliqnro,cliqnro,gratprop,difimponibleact,rentaimpoact,impunicoaju,difimponible,concnro,aux1)"
                StrSql = StrSql & " VALUES ("
                StrSql = StrSql & " " & rs_Periodos!PliqNro
                StrSql = StrSql & "," & buliq_cabliq!cliqnro
                StrSql = StrSql & "," & Abs(BonoPeriodo)
                StrSql = StrSql & "," & Abs(DifUFHist)
                StrSql = StrSql & "," & Abs(ImpoMesHistReliq)
                StrSql = StrSql & "," & Abs(MontoConcRecalc + DifUFHistPorc)
                StrSql = StrSql & "," & Abs(DifUFHistPorc)
                StrSql = StrSql & "," & Buliq_Concepto(Concepto_Actual).concnro
                StrSql = StrSql & "," & buliq_empleado!ternro
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
Dim PorcZonaExtHist As Double
Dim ImpPagado As Double
Dim ImpoHistBono As Double
Dim BonoAjustado As Double
Dim UTMHist As Double
Dim EUSHist As Double
Dim DeduccHist As Double
Dim NuevoImpoMesHist As Double
Dim desdeEsc As Double
Dim HastaEsc As Double
Dim factorEsc As Double
Dim rebajaEsc As Double
Dim ImpRecalculado As Double
Dim ImpRecalculadoAcum As Double
Dim ImpoAntRecalc As Double

Dim rs_Periodos As New ADODB.Recordset
Dim rs_consult As New ADODB.Recordset

Const c_Bono = 104
Const c_AcuImp = 98
Const c_AcuPorc = 35

'Parametros
Dim Bono As Double
Dim AcuImp As Long
Dim AcuPorc As Long

Dim EncBono As Boolean
Dim EncAcuImp As Boolean
Dim EncAcuPorc As Boolean


    Bien = False
    ImpRecalculadoAcum = 0
    PorcZonaExtHist = 0
     
    EncBono = False
    EncAcuImp = False
    EncAcuPorc = False

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
            Case c_AcuPorc:
                AcuPorc = rs_consult!Valor
                EncAcuPorc = True
            Case Else
        End Select
        
        rs_consult.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not EncBono Or Not EncAcuImp Then
        Exit Function
    End If
    
    If Not EncAcuPorc Then
        AcuPorc = 0
    End If
    
    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "---------Parametros-----------------------------"
        Flog.writeline Espacios(Tabulador * 3) & "Valor Bono: " & Bono
        Flog.writeline Espacios(Tabulador * 3) & "Acumulador Imponible: " & AcuImp
        Flog.writeline Espacios(Tabulador * 3) & "Acumulador % Zona Extrema: " & AcuPorc
        Flog.writeline
    End If

    'Creo la tabla temporal para la Escala de UTM
    'Call CreateTempTable(TTempWF_EscalaUTM)
    
    
    'Busco la cantidad de periodos de recalculo del proceso
    CantPerRec = 0
    StrSql = "SELECT periodo.* FROM impuni_peri "
    StrSql = StrSql & " INNER JOIN periodo ON periodo.pliqnro = impuni_peri.pliqnro"
    StrSql = StrSql & " WHERE pronro = " & buliq_proceso!pronro
    StrSql = StrSql & " ORDER BY periodo.pliqdesde "
    
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
        Flog.writeline Espacios(Tabulador * 3) & "Bono por Periodo: " & FormatNumber(BonoPeriodo, 2)
        Flog.writeline Espacios(Tabulador * 3) & "UTM Actual: " & FormatNumber(buliq_periodo!pliqutm, 2)
        Flog.writeline
    End If
    
    Do While Not rs_Periodos.EOF
        
        If CBool(USA_DEBUG) Then
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
                Flog.writeline Espacios(Tabulador * 4) & "Imponible Periodo: " & FormatNumber(rs_consult!ammonto, 2)
            End If
        Else
            ImpoMesHist = 0
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontro Imponible en el periodo con Acum: " & AcuImp
                Flog.writeline Espacios(Tabulador * 4) & "Siguiente periodo. "
                GoTo sgt_periodo
            End If
        End If
        If rs_consult.State = adStateOpen Then rs_consult.Close
        
        'Busco el % de Zona Extrema del mes
        PorcZonaExtHist = 0
        StrSql = "SELECT ammonto, amcant"
        StrSql = StrSql & " FROM acu_mes"
        StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND acunro = " & AcuPorc
        StrSql = StrSql & " AND  amanio = " & rs_Periodos!pliqanio
        StrSql = StrSql & " AND ammes = " & rs_Periodos!pliqmes
        OpenRecordset StrSql, rs_consult
        
        If Not rs_consult.EOF Then
            PorcZonaExtHist = rs_consult!ammonto
        Else
            PorcZonaExtHist = 0
        End If
        If rs_consult.State = adStateOpen Then rs_consult.Close
        
        'Busco el UTM del mes
        UTMHist = rs_Periodos!pliqutm
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "UTM Periodo: " & FormatNumber(UTMHist, 2)
        End If
                
        BonoAjustado = (BonoPeriodo / buliq_periodo!pliqutm) * UTMHist
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Bono Ajustado UTM: " & FormatNumber(BonoAjustado, 2)
        End If
        
        
        'Buscar en impunicab con el ajuste de cada cpto reliquidado de deducciones
        DeduccHist = 0
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
            Flog.writeline Espacios(Tabulador * 4) & "Deducciones (Cptos Reliq): " & FormatNumber(DeduccHist, 2)
        End If
        
        'El ajuste de cada cpto reliquidado de deducciones lo paso a historico
        DeduccHist = (DeduccHist / buliq_periodo!pliqutm) * UTMHist
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Deducciones Hist Ajus (Cptos Reliq): " & FormatNumber(DeduccHist, 2)
        End If
        
        
        'Buscar en impunicab el imponible anterior usado en la escala por recalculo de IU anterior
        ImpoAntRecalc = 0
        StrSql = "SELECT SUM(rentaimpoact) ImpoRec FROM impuni_cab"
        StrSql = StrSql & " WHERE impuni_cab.pliqnro = " & rs_Periodos!PliqNro
        StrSql = StrSql & " AND impuni_cab.concnro = " & Buliq_Concepto(Concepto_Actual).concnro
        StrSql = StrSql & " AND impuni_cab.aux1 = " & buliq_empleado!ternro
        StrSql = StrSql & " AND impuni_cab.cliqnro <> " & buliq_cabliq!cliqnro
        OpenRecordset StrSql, rs_consult
        If Not rs_consult.EOF Then
            If Not EsNulo(rs_consult!ImpoRec) Then
                ImpoAntRecalc = rs_consult!ImpoRec
            End If
            
        End If
        If CBool(USA_DEBUG) Then
              Flog.writeline Espacios(Tabulador * 4) & "Imponible Periodo IU otras Reliq: " & FormatNumber(ImpoAntRecalc, 2)
        End If
        
        'si no hay imponible anterior guardado en impuni_cab tomo el imponible del acu_mes
        If Not (ImpoAntRecalc = 0) Then
             ImpoMesHist = ImpoAntRecalc
        End If
        
          
        ImpoHistBono = Round((ImpoMesHist + BonoAjustado), 0)
        
        NuevoImpoMesHist = Round((ImpoHistBono - DeduccHist), 0)
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Imponible Periodo + Bono Ajustado: " & FormatNumber(ImpoHistBono, 0)
            Flog.writeline Espacios(Tabulador * 4) & "Nuevo Tributable (ImpPer + Bono - Deducc Hist): " & FormatNumber(NuevoImpoMesHist, 0)
        End If
        
        'Calculo de Zona Extrema
        EUSHist = 0
        If PorcZonaExtHist > 0 Then
            'Busco la Escala Unica Salarial del mes para Zona Extrema
            If Not EsNulo(rs_Periodos!pliqescunisal) Then
              EUSHist = rs_Periodos!pliqescunisal
            End If
        
            MontoZonaExt1 = (EUSHist * PorcZonaExtHist) / 100
            MontoZonaExt2 = (NuevoImpoMesHist * PorcZonaExtHist) / 100
            If MontoZonaExt1 < MontoZonaExt2 Then
                NuevoImpoMesHist = Round((NuevoImpoMesHist - MontoZonaExt1), 0)
            Else
                NuevoImpoMesHist = Round((NuevoImpoMesHist - MontoZonaExt2), 0)
            End If
        
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Calculo de Zona Extrema: "
                Flog.writeline Espacios(Tabulador * 4) & "% Zona Extrema Periodo: " & FormatNumber(PorcZonaExtHist, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Escala Unica Salarial Periodo: " & FormatNumber(EUSHist, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Monto1: " & FormatNumber(MontoZonaExt1, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Monto2: " & FormatNumber(MontoZonaExt2, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Tributable menos Zona Extrema( menor de Monto1 o Monto2): " & FormatNumber(NuevoImpoMesHist, 0)
             End If
        Else
           If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No tiene calculo de Zona Extrema "
           End If
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
                Flog.writeline Espacios(Tabulador * 4) & "Entra en escala con Tributable: " & FormatNumber(Abs(NuevoImpoMesHist), 0)
                Flog.writeline Espacios(Tabulador * 5) & "Desde: " & FormatNumber(desdeEsc, 2)
                Flog.writeline Espacios(Tabulador * 5) & "Hasta: " & FormatNumber(HastaEsc, 2)
                Flog.writeline Espacios(Tabulador * 5) & "Factor: " & FormatNumber(factorEsc, 2)
                Flog.writeline Espacios(Tabulador * 5) & "Rebaja: " & FormatNumber(rebajaEsc, 2)
            End If
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encuentra escala con " & FormatNumber(NuevoImpoMesHist, 0)
            End If
        End If
        
        ' Nuevo impuesto
        ImpRecalculado = (NuevoImpoMesHist * factorEsc) - rebajaEsc
        
        If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Impuesto Unico Recalc " & FormatNumber(ImpRecalculado, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Buscar Imp Unico ya pagado en el Periodo"
        End If
        'Restar al impuesto lo ya liquidado por impuesto unico
        'Entrar nuevamente a escala con el Imponible historico
        
        'Calculo de Zona Extrema para restar del Imponible historico
        If PorcZonaExtHist > 0 Then
            MontoZonaExt1 = (EUSHist * PorcZonaExtHist) / 100
            MontoZonaExt2 = (ImpoMesHist * PorcZonaExtHist) / 100
            If MontoZonaExt1 < MontoZonaExt2 Then
                ImpoMesHist = ImpoMesHist - MontoZonaExt1
            Else
                ImpoMesHist = ImpoMesHist - MontoZonaExt2
            End If
        
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Nuevo Calculo de Zona Extrema: "
                Flog.writeline Espacios(Tabulador * 4) & "% Zona Extrema Periodo: " & FormatNumber(PorcZonaExtHist, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Escala Unica Salarial Periodo: " & FormatNumber(EUSHist, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Monto1: " & FormatNumber(MontoZonaExt1, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Monto2: " & FormatNumber(MontoZonaExt2, 2)
                Flog.writeline Espacios(Tabulador * 4) & "Tributable Hist menos Zona Extrema( menor de Monto1 o Monto2): " & FormatNumber(ImpoMesHist, 2)
             End If
        Else
           If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No tiene calculo de Zona Extrema "
           End If
        End If
        
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
                Flog.writeline Espacios(Tabulador * 5) & "Desde: " & FormatNumber(desdeEsc, 2)
                Flog.writeline Espacios(Tabulador * 5) & "Hasta: " & FormatNumber(HastaEsc, 2)
                Flog.writeline Espacios(Tabulador * 5) & "Factor: " & FormatNumber(factorEsc, 2)
                Flog.writeline Espacios(Tabulador * 5) & "Rebaja: " & FormatNumber(rebajaEsc, 2)
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
            Flog.writeline Espacios(Tabulador * 4) & "Resultado del Impuesto a Pagar (Imp Recal - Imp Pagado): " & FormatNumber(ImpRecalculado, 2)
            Flog.writeline Espacios(Tabulador * 4) & "Acumulo en UTM: " & FormatNumber((ImpRecalculado / UTMHist), 2)
            Flog.writeline
        End If
        
        'Guardo en impuni_cab
        StrSql = "INSERT INTO impuni_cab (pliqnro,cliqnro,gratprop,difimponibleact,rentaimpoact,impunicoaju,difimponible,concnro,aux1)"
                StrSql = StrSql & " VALUES ("
                StrSql = StrSql & " " & rs_Periodos!PliqNro
                StrSql = StrSql & "," & buliq_cabliq!cliqnro
                StrSql = StrSql & "," & Abs(BonoPeriodo)
                StrSql = StrSql & "," & Abs(0)
                StrSql = StrSql & "," & Abs(NuevoImpoMesHist)
                StrSql = StrSql & "," & Abs(ImpRecalculado)
                StrSql = StrSql & "," & Abs(ImpRecalculado)
                StrSql = StrSql & "," & Buliq_Concepto(Concepto_Actual).concnro
                StrSql = StrSql & "," & buliq_empleado!ternro
                StrSql = StrSql & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
        
sgt_periodo:
        rs_Periodos.MoveNext
    Loop
    
    exito = True
        
    for_RecalcImpuestoUnico = Abs(ImpRecalculadoAcum * buliq_periodo!pliqutm) * -1
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "_______________________________________________________________________________________________"
        Flog.writeline Espacios(Tabulador * 3) & "Total Recalc. IU a Pagar: " & FormatNumber((ImpRecalculadoAcum * buliq_periodo!pliqutm), 2)
    End If
    
    

'Libero Recordset
If rs_Periodos.State = adStateOpen Then rs_Periodos.Close
Set rs_Periodos = Nothing
If rs_consult.State = adStateOpen Then rs_consult.Close
Set rs_consult = Nothing

End Function

