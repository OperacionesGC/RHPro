Attribute VB_Name = "Mdlformulas"
Option Explicit

Public Function EjecutarFormulaNoConfigurable(ByVal nombre As String) As Single
Dim Nro As String

    Select Case UCase(nombre)
'-------------------------------------------------
' Inicio Formulas de Moño Azul
    Case "FOR_203":
        EjecutarFormulaNoConfigurable = for_203
    Case "FOR_204":
        EjecutarFormulaNoConfigurable = for_204
    Case "FOR_207":
        EjecutarFormulaNoConfigurable = for_207
    Case "FOR_PROM_PRODUCCION":
        EjecutarFormulaNoConfigurable = For_Prom_Produccion
' Fin Formulas de Moño Azul
'-------------------------------------------------

    End Select
End Function

Public Function EjecutarFormulaDeSistema(ByVal nombre As String) As Single

    Select Case UCase(nombre)
    Case "FOR_NEG":
        EjecutarFormulaDeSistema = for_neg
    Case "FOR_POS":
        EjecutarFormulaDeSistema = for_pos
    Case "FOR_PORC_POS":
        EjecutarFormulaDeSistema = for_porc_pos
    Case "FOR_PORC_NEG":
        EjecutarFormulaDeSistema = for_porc_neg
    Case "FOR_GANANCIAS":
        EjecutarFormulaDeSistema = for_Ganancias(NroCab, fec, Monto, Bien)
    Case "FOR_GANANCIAS_SCHERING":
        EjecutarFormulaDeSistema = for_Ganancias_Schering(NroCab, fec, Monto, Bien)
    Case "FOR_GROSSING3":
        EjecutarFormulaDeSistema = for_Grossing3(NroCab, fec, Monto, Bien)
    Case "FOR_GROSSING5":
        EjecutarFormulaDeSistema = for_Grossing5
    Case "FOR_BASEACCI":
        EjecutarFormulaDeSistema = For_Baseacci
    Case "FOR_NIVELAR":
        EjecutarFormulaDeSistema = for_Nivelar(fec)
    Case "FOR_IRP":
        EjecutarFormulaDeSistema = for_irp(NroCab, fec, Monto, Bien)
    Case "FOR_IRP_FRANJA":
        EjecutarFormulaDeSistema = for_irp_franja(NroCab, fec, Monto, Bien)
    Case Else
    End Select
   
End Function

' ---------------------------------------------------------
' Modulo de fórmulas conocidas
' ---------------------------------------------------------

Public Function for_neg() As Single
' Monto Negativo
Dim v_monto As Single
Dim rs_wf_tpa As New ADODB.Recordset
Dim Encontro As Boolean
    
    exito = False
    Encontro = False
    
    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    If Not rs_wf_tpa.EOF Then
        v_monto = rs_wf_tpa!Valor
        Encontro = True
    End If

    ' si no se obtuvieron los parametros, ==> Error.
    If Encontro Then
        v_monto = -v_monto
    Else
        Exit Function
    End If
    
    exito = True
    for_neg = v_monto
    
End Function


Public Function for_pos() As Single
' Monto Positivo
Dim v_monto As Single
Dim rs_wf_tpa As New ADODB.Recordset
Dim Encontro As Boolean
    
    exito = False
    Encontro = False
    
    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    If Not rs_wf_tpa.EOF Then
        v_monto = rs_wf_tpa!Valor
        Encontro = True
    End If

    ' si no se obtuvieron los parametros, ==> Error.
    If Not Encontro Then
        Exit Function
    End If
    
    exito = True
    for_pos = v_monto
    
End Function

Public Function for_porc_neg() As Single
' Porcentaje Negativo
Const c_porcentaje = 35
Const c_msr = 8

Dim v_porcentaje As Single
Dim v_msr As Single

Dim rs_wf_tpa As New ADODB.Recordset
Dim Encontro_msr As Boolean
Dim Encontro_porcentaje As Boolean

    exito = False
    Encontro_msr = False
    Encontro_porcentaje = False
    
    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case c_msr:
            v_msr = rs_wf_tpa!Valor
            Encontro_msr = True
        Case c_porcentaje:
            v_porcentaje = rs_wf_tpa!Valor
            Encontro_porcentaje = True
        Case Else
        End Select
        
        rs_wf_tpa.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not Encontro_msr Or Not Encontro_porcentaje Then
        Exit Function
    End If
    
    for_porc_neg = -(v_msr * v_porcentaje / 100)
    exito = True
    
End Function


Public Function for_porc_pos() As Single
' Porcentaje Positivo
Const c_porcentaje = 35
Const c_msr = 8

Dim v_porcentaje As Single
Dim v_msr As Single

Dim rs_wf_tpa As New ADODB.Recordset
Dim Encontro_msr As Boolean
Dim Encontro_porcentaje As Boolean

    exito = False
    Encontro_msr = False
    Encontro_porcentaje = False
    
    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case c_msr:
            v_msr = rs_wf_tpa!Valor
            Encontro_msr = True
        Case c_porcentaje:
            v_porcentaje = rs_wf_tpa!Valor
            Encontro_porcentaje = True
        Case Else
        End Select
        
        rs_wf_tpa.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not Encontro_msr Or Not Encontro_porcentaje Then
        Exit Function
    End If
    
      
    for_porc_pos = (v_msr * v_porcentaje / 100)
    exito = True
    
End Function

Public Function for_Ganancias_old(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Single, Bien As Boolean) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias.
' Autor      :
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim p_prorratea As Integer  'Si prorratea o no para liq. finales

'Variables Locales
Dim Devuelve As Single
Dim Tope_Gral As Single
Dim Neto As Single
Dim prorratea As Single
Dim Retencion As Single
Dim Gan_Imponible As Single
Dim Deducciones As Single
Dim Ded_a23 As Single
Dim Por_Deduccion As Single
Dim Impuesto_Escala As Single
Dim Ret_Ant As Single

Dim Ret_mes As Integer
Dim Ret_ano As Integer
Dim fin_mes_ret As Date
Dim ini_anyo_ret As Date
Dim Con_liquid As Integer
Dim i As Integer
Dim Texto As String

'Vectores para manejar el proceso
Dim Items_DDJJ(50) As Single
Dim Items_LIQ(50) As Single
Dim Items_PRORR(50) As Single
Dim Items_OLD_LIQ(50) As Single
Dim Items_TOPE(50) As Single

'Recorsets Auxiliares
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_item As New ADODB.Recordset
Dim rs_valitem As New ADODB.Recordset
Dim rs_desmen As New ADODB.Recordset
Dim rs_desliq As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_escala_ded As New ADODB.Recordset
Dim rs_escala As New ADODB.Recordset
Dim rs_ficharet As New ADODB.Recordset

Dim Hasta As Integer

' FGZ - 12/02/2004
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Single
' FGZ - 12/02/2004

' FGZ - 27/02/2004
Dim Terminar As Boolean
Dim pos1
Dim pos2
' FGZ - 27/02/2004

'Comienzo
p_Devuelve = 1001
p_Tope_Gral = 1002
p_Neto = 1003
p_prorratea = 1005

Bien = False

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

Ret_mes = Month(buliq_proceso!profecpago)
Ret_ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_mes = 12, CDate("01/01/" & Ret_ano + 1) - 1, CDate("01/" & Ret_mes + 1 & "/" & Ret_ano) - 1)
ini_anyo_ret = CDate("01/01/" & Ret_ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).concnro


'Obtencion de los parametros de WorkFile

StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case p_Devuelve:
        Devuelve = rs_wf_tpa!Valor
    Case p_Tope_Gral:
        Tope_Gral = rs_wf_tpa!Valor
    Case p_Neto:
        Neto = rs_wf_tpa!Valor
    Case p_prorratea:
        prorratea = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Máxima Ret. en %" & Tope_Gral
    Flog.writeline Espacios(Tabulador * 1) & "Neto del Mes" & Neto
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Tope_Gral, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Neto, "Neto del Mes", Neto)
End If

'Si es una liq. final no prorratea y tomo la escala de diciembre
If prorratea = 0 Then
    Ret_mes = 12
End If


' Recorro todos los items de Ganancias
StrSql = "SELECT * FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_item

Do While Not rs_item.EOF
    
    Select Case rs_item!itetipotope
    Case 1: ' el valor a tomar es lo que dice la escala
    
        StrSql = "SELECT * FROM valitem WHERE viano =" & Ret_ano & _
                 " AND itenro=" & rs_item!itenro & _
                 " AND vimes =" & Ret_mes
        OpenRecordset StrSql, rs_valitem
        
        Do While Not rs_valitem.EOF
            Items_DDJJ(rs_valitem!itenro) = rs_valitem!vimonto
            Items_TOPE(rs_valitem!itenro) = rs_valitem!vimonto
            
            rs_valitem.MoveNext
        Loop
    ' End case 1
    ' ------------------------------------------------------------------------
    
    Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
        ' Busco la declaracion jurada
        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!ternro & _
                 " AND desano=" & Ret_ano & _
                 " AND itenro = " & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
        
        Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_item!itenro = 3 Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                Else
                    If rs_desmen!desmenprorra = 0 Then 'no es parejito
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec
                    Else
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                    End If
                End If
            End If
            
            rs_desmen.MoveNext
        Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto
            'Si el desliq prorratea debo proporcionarlo
            Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 Or CBool(rs_desliq!dlprorratea)) And (prorratea = 1), rs_desliq!dlmonto / (13 - Month(rs_desliq!dlfecha)) * (Ret_mes - Month(rs_desliq!dlfecha) + 1), rs_desliq!dlmonto)

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itaprorratea) Then
                    If CBool(rs_itemacum!itasigno) Then
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    End If
                Else
                    If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    End If
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
' FGZ - 12/02/2004 Desde acá -------------------------
'        StrSql = "SELECT * FROM itemacum " & _
'                 " INNER JOIN acu_liq ON itemacum.acunro = acu_liq.acunro " & _
'                 " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND itenro =" & rs_item!itenro & _
'                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemacum
'
'        Do While Not rs_itemacum.EOF
'            If CBool(rs_itemacum!itaprorratea) Then
'                If CBool(rs_itemacum!itasigno) Then
'                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + rs_itemacum!almonto
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, rs_itemacum!almonto / (13 - Ret_mes), rs_itemacum!almonto)
'                Else
'                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - rs_itemacum!almonto
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, rs_itemacum!almonto / (13 - Ret_mes), rs_itemacum!almonto)
'                End If
'            Else
'                If CBool(rs_itemacum!itasigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemacum!almonto / (13 - Ret_mes), rs_itemacum!almonto)
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemacum!almonto / (13 - Ret_mes), rs_itemacum!almonto)
'                End If
'            End If
'
'            rs_itemacum.MoveNext
'        Loop
' FGZ - 12/02/2004 Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) Then
                If CBool(rs_itemconc!itcsigno) Then
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                End If
            Else
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                End If
            End If
        
            rs_itemconc.MoveNext
        Loop
    ' End case 2
    ' ------------------------------------------------------------------------
        
' ****************************************************************************
' *  OJO QUEDA PENDIENTE EL PRORRATEO PARA LOS ITEMS DE TIPO 3 Y 5           *
' ****************************************************************************


     Case 3: 'TOMO LOS VALORES DE LA DDJJ Y LIQUIDACION Y EL TOPE PARA APLICARLO
     
        StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                 " AND vimes = " & Ret_mes & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_valitem
         Do While Not rs_valitem.EOF
            Items_TOPE(rs_item!itenro) = rs_valitem!vimonto
         
            rs_valitem.MoveNext
         Loop
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
         
            rs_desmen.MoveNext
         Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
' FGZ - 12/02/2004 Desde acá -------------------------
'        StrSql = "SELECT * FROM itemacum " & _
'                 " INNER JOIN acu_liq ON itemacum.acunro = acu_liq.acunro " & _
'                 " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND itenro =" & rs_item!itenro & _
'                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemacum
'
'        Do While Not rs_itemacum.EOF
'                If CBool(rs_itemacum!itasigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
'                End If
'
'            rs_itemacum.MoveNext
'        Loop
' FGZ - 12/02/2004 Hasta acá -------------------------

        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
        
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_item!itesigno) Then
            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
        End If
        
    ' End case 3
    ' ------------------------------------------------------------------------
   
    Case 4:
        ' Tomo los valores de la DDJJ y el valor de la escala (cargas de familia)
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfechas) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Month(rs_desmen!desfechas) - Month(rs_desmen!desfecdes) + 1)
            Else
                If Month(rs_desmen!desfecdes) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
                End If
            End If
        
            rs_desmen.MoveNext
         Loop
        
        If Items_DDJJ(rs_item!itenro) > 0 Then
            StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                     " AND vimes = " & Ret_mes & _
                     " AND itenro =" & rs_item!itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_item!itenro) = rs_valitem!vimonto / Ret_mes * Items_DDJJ(rs_item!itenro)
             
                rs_valitem.MoveNext
             Loop
        End If
    ' End case 4
    ' ------------------------------------------------------------------------
        
    Case 5:
        i = 1
        Hasta = IIf(50 > Len(rs_item!iteitemstope), 50, rs_item!iteitemstope)
        Terminar = False
        Do While i <= Hasta And Not Terminar
            pos1 = i
            pos2 = InStr(i, rs_item!iteitemstope, ",") - 1
            If pos2 > 0 Then
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
            Else
                pos2 = Len(rs_item!iteitemstope)
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
                Terminar = True
            End If
            
            If Texto <> "" Then
                If Mid(Texto, 1, 1) = "-" Then
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                Else
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                End If
            End If
            i = pos2 + 2
        Loop
        
        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) * rs_item!iteporctope / 100
    
    
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
         
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
            ' Tocado por Maxi 26/05/2004 faltaba el parejito
            'If Month(rs_desmen!desfecdes) <= Ret_mes Then
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
            'Else
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
            'End If
         
            rs_desmen.MoveNext
         Loop
    
    
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro=" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
' FGZ - 12/02/2004 Desde acá -------------------------
'        StrSql = "SELECT * FROM itemacum " & _
'                 " INNER JOIN acu_liq ON itemacum.acunro = acu_liq.acunro " & _
'                 " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND itenro =" & rs_item!itenro & _
'                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemacum
'
'        Do While Not rs_itemacum.EOF
'                If CBool(rs_itemacum!itasigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
'                End If
'
'            rs_itemacum.MoveNext
'        Loop
' FGZ - 12/02/2004 Hasta acá -------------------------

        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
    
        'TOPEO LOS VALORES
        If Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro)
        End If
    ' End case 5
    ' ------------------------------------------------------------------------
    Case Else:
    End Select
    
    
    ' SI SE TOMA PARA LA GANANCIA NETA, DA VUELTA EL SIGNO DEL TOPE SOLO PARA ITEMS
    ' QUE SE TOPEAN DE ALGUNA FORMA Y NO SALEN DEL RECIBO DE SUELDO. "Como saber que no sale del Recibo"
    
    If rs_item!itenro > 7 Then
        Items_TOPE(rs_item!itenro) = IIf(CBool(rs_item!itesigno), Items_TOPE(rs_item!itenro), Abs(Items_TOPE(rs_item!itenro)))
    End If
    
    'Armo la traza del item
    If CBool(USA_DEBUG) Then
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-DDJJ"
        Flog.writeline Espacios(Tabulador * 1) & Items_DDJJ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Liq"
        Flog.writeline Espacios(Tabulador * 1) & Items_LIQ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-LiqAnt"
        Flog.writeline Espacios(Tabulador * 1) & Items_OLD_LIQ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Prorr"
        Flog.writeline Espacios(Tabulador * 1) & Items_PRORR(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Tope"
        Flog.writeline Espacios(Tabulador * 1) & Items_TOPE(rs_item!itenro)
    End If
    If HACE_TRAZA Then
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_DDJJ(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_LIQ(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_OLD_LIQ(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_PRORR(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_TOPE(rs_item!itenro))
    End If
        
    'Calcula la Ganancia Imponible
    If CBool(rs_item!itesigno) Then
        Gan_Imponible = Gan_Imponible + Items_TOPE(rs_item!itenro)
    Else
        If (rs_item!itetipotope = 1) Or (rs_item!itetipotope = 4) Then
            Ded_a23 = Ded_a23 - Items_TOPE(rs_item!itenro)
        Else
            Deducciones = Deducciones - Items_TOPE(rs_item!itenro)
        End If
    End If
            
    rs_item.MoveNext
Loop
            
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "9- Ganancia Neta " & Gan_Imponible
        Flog.writeline Espacios(Tabulador * 1) & "9- Total Deducciones" & Deducciones
        Flog.writeline Espacios(Tabulador * 1) & "9- Total art. 23" & Ded_a23
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Neta ", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Total art. 23", Ded_a23)
    End If
            
    
    ' Calculo el porcentaje de deduccion segun la ganancia neta
    If Ret_ano >= 2000 And Gan_Imponible > 0 Then
        StrSql = "SELECT * FROM escala_ded " & _
                 " WHERE esd_topeinf <= " & (Gan_Imponible / Ret_mes * 12) & _
                 " AND esd_topesup >=" & (Gan_Imponible / Ret_mes * 12)
        OpenRecordset StrSql, rs_escala_ded
    
        If Not rs_escala_ded.EOF Then
            Por_Deduccion = rs_escala_ded!esd_porcentaje
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "No hay esc. dedu para" & Gan_Imponible
            End If
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "No hay esc. dedu para", Gan_Imponible)
            End If
            ' No se ha encontrado la escala de deduccion para el valor gan_imponible
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "9- % a tomar deduc." & Por_Deduccion
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- % a tomar deduc.", Por_Deduccion)
        End If
        
        'Aplico el porcentaje a las deducciones
        Ded_a23 = Ded_a23 * Por_Deduccion / 100
    End If
            
    
    ' Calculo la Ganancia imponible
    Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "9- Ganancia Imponible" & Gan_Imponible
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Imponible", Gan_Imponible)
    End If
            
                
    If Gan_Imponible > 0 Then
        'Entrar en la escala con las ganancias acumuladas
        StrSql = "SELECT * FROM escala " & _
                 " WHERE escmes =" & Ret_mes & _
                 " AND escano =" & Ret_ano & _
                 " AND escinf <= " & Gan_Imponible & _
                 " AND escsup >= " & Gan_Imponible
        OpenRecordset StrSql, rs_escala
        
        If Not rs_escala.EOF Then
            Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
        Else
            Impuesto_Escala = 0
        End If
    Else
        Impuesto_Escala = 0
    End If
            
            
    ' Calculo las retenciones ya realizadas
    Ret_Ant = 0
        
    'For each ficharet where ficharet.empleado = buliq-empleado.ternro
    '                    And Month(ficharet.fecha) <= ret-mes
    '                    And Year(ficharet.fecha) = ret-ano NO-LOCK:
    '    Assign Ret-ant = Ret-Ant + ficharet.importe.
    'End.
    
    'como no puede utilizar la funcion month() en sql
    'levanto todas las ficharet del tercero y hago la pregunta dentro del loop
    StrSql = "SELECT * FROM ficharet " & _
             " WHERE empleado =" & buliq_empleado!ternro
    OpenRecordset StrSql, rs_ficharet
    
    Do While Not rs_ficharet.EOF
        If (Month(rs_ficharet!Fecha) <= Ret_mes) And (Year(rs_ficharet!Fecha) = Ret_ano) Then
            Ret_Ant = Ret_Ant + rs_ficharet!Importe
        End If
        rs_ficharet.MoveNext
    Loop
    
    
    'Calcular la retencion
    Retencion = Impuesto_Escala - Ret_Ant
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "Retenciones anteriores" & Ret_Ant
        If Gan_Imponible > 0 Then
            If Not rs_escala.EOF Then
                rs_escala.MoveFirst
                If Not rs_escala.EOF Then
                    Flog.writeline Espacios(Tabulador * 1) & "Escala Impuesto" & rs_escala!escporexe
                    Flog.writeline Espacios(Tabulador * 1) & "Impuesto por escala" & Impuesto_Escala
                    Flog.writeline Espacios(Tabulador * 1) & "A Retener/Devolver" & Retencion
                Else
                    Flog.writeline Espacios(Tabulador * 1) & "Escala Impuesto" & "0"
                    Flog.writeline Espacios(Tabulador * 1) & "Impuesto por escala" & Impuesto_Escala
                    Flog.writeline Espacios(Tabulador * 1) & "A Retener/Devolver" & Retencion
                End If
            End If
        End If
        
    End If
    
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Retenciones anteriores", Ret_Ant)
        If Gan_Imponible > 0 Then
            If Not rs_escala.EOF Then
                rs_escala.MoveFirst
                If Not rs_escala.EOF Then
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Escala Impuesto", rs_escala!escporexe)
                Else
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Escala Impuesto", 0)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver", Retencion)
                End If
            End If
        End If
    End If
    
    ' Verifico si es una devolucion y si devuelve el concepto
    If Devuelve = 0 And Retencion < 0 Then
        Retencion = 0
    End If
    
    ' FGZ - 14/04/2004
    If Retencion <> 0 Then
        ' Verificar que la rtencion no supere el 30% del Neto del Mes
        If Retencion > (Neto * (Tope_Gral / 100)) Then
            Retencion = Neto * (Tope_Gral / 100)
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "A Retener/Devolver, x Tope General" & Retencion
            End If
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver, x Tope General", Retencion)
            End If
        End If
        Monto = -Retencion
    Else
        Monto = 0
    End If
    'Monto = -Retencion
    Bien = True
    
        
    'Retenciones / Devoluciones
    If Retencion <> 0 Then
        Call InsertarFichaRet(buliq_empleado!ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
    End If
    
    ' Grabo todos los items de la liquidacion actual
    i = 1
    Hasta = 50
    Do While i <= Hasta
        If Items_LIQ(i) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!ternro & "," & _
                     ConvFecha(buliq_proceso!profecpago) & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_LIQ(i) & "," & _
                     "0," & _
                     i & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        If Items_PRORR(i) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!ternro & "," & _
                     ConvFecha(buliq_proceso!profecpago) & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_PRORR(i) & "," & _
                     "0," & _
                     i & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        i = i + 1
    Loop

exito = Bien
for_Ganancias_old = Monto
End Function


Public Function for_Ganancias(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Single, Bien As Boolean) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias.
' Autor      :
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim p_prorratea As Integer  'Si prorratea o no para liq. finales

'Variables Locales
Dim Devuelve As Single
Dim Tope_Gral As Single
Dim Neto As Single
Dim prorratea As Single
Dim Retencion As Single
Dim Gan_Imponible As Single
Dim Deducciones As Single
Dim Ded_a23 As Single
Dim Por_Deduccion As Single
Dim Impuesto_Escala As Single
Dim Ret_Ant As Single

Dim Ret_mes As Integer
Dim Ret_ano As Integer
Dim fin_mes_ret As Date
Dim ini_anyo_ret As Date
Dim Con_liquid As Integer
Dim i As Integer
Dim j As Integer
Dim Texto As String

'Vectores para manejar el proceso
Dim Items_DDJJ(50) As Single
Dim Items_LIQ(50) As Single
Dim Items_PRORR(50) As Single
Dim Items_OLD_LIQ(50) As Single
Dim Items_TOPE(50) As Single
Dim Items_ART_23(50) As Boolean

'Recorsets Auxiliares
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_item As New ADODB.Recordset
Dim rs_valitem As New ADODB.Recordset
Dim rs_desmen As New ADODB.Recordset
Dim rs_desliq As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_escala_ded As New ADODB.Recordset
Dim rs_escala As New ADODB.Recordset
Dim rs_ficharet As New ADODB.Recordset
Dim rs_traza_gan_items_tope As New ADODB.Recordset

Dim Hasta As Integer

' FGZ - 12/02/2004
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Single
' FGZ - 12/02/2004

' FGZ - 27/02/2004
Dim Terminar As Boolean
Dim pos1
Dim pos2
' FGZ - 27/02/2004

'Comienzo
p_Devuelve = 1001
p_Tope_Gral = 1002
p_Neto = 1003
p_prorratea = 1005

Bien = False


'FGZ - 19/04/2004
Dim Total_Empresa As Single
Dim Tope As Integer
'Dim rs_Rep19 As New ADODB.Recordset
Dim rs_traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Single
Total_Empresa = 0
Tope = 10

' Primero limpio la traza
StrSql = "DELETE FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
objConn.Execute StrSql, , adExecuteNoRecords

' Guardo la traza de Ganancia en traza_gan para utilizarla luego cuando se genere el reporte
StrSql = "INSERT INTO traza_gan (pliqnro,pronro,concnro,fecha_pago,empresa,ternro,empleg) VALUES (" & _
         buliq_periodo!PliqNro & "," & _
         buliq_proceso!pronro & "," & _
         Buliq_Concepto(Concepto_Actual).concnro & "," & _
         ConvFecha(buliq_proceso!profecpago) & "," & _
         NroEmp & "," & _
         buliq_empleado!ternro & "," & _
         buliq_empleado!empleg & _
         ")"
objConn.Execute StrSql, , adExecuteNoRecords

'FGZ - 22/12/2004
'Leo la tabla
StrSql = "SELECT * FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
OpenRecordset StrSql, rs_traza_gan
    

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

Ret_mes = Month(buliq_proceso!profecpago)
Ret_ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_mes = 12, CDate("01/01/" & Ret_ano + 1) - 1, CDate("01/" & Ret_mes + 1 & "/" & Ret_ano) - 1)
ini_anyo_ret = CDate("01/01/" & Ret_ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).concnro

'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case p_Devuelve:
        Devuelve = rs_wf_tpa!Valor
    Case p_Tope_Gral:
        Tope_Gral = rs_wf_tpa!Valor
    Case p_Neto:
        Neto = rs_wf_tpa!Valor
    Case p_prorratea:
        prorratea = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

'Si es una liq. final no prorratea y tomo la escala de diciembre
If prorratea = 0 Then
    Ret_mes = 12
    'FGZ - 27/09/2004
    fin_mes_ret = IIf(Ret_mes = 12, CDate("01/01/" & Ret_ano + 1) - 1, CDate("01/" & Ret_mes + 1 & "/" & Ret_ano) - 1)
End If

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 3) & "---------Formula-----------------------------"
    Flog.writeline Espacios(Tabulador * 3) & "Fecha del Proceso " & buliq_proceso!profecpago
    Flog.writeline Espacios(Tabulador * 3) & "Mes de Retencion " & Ret_mes
    Flog.writeline Espacios(Tabulador * 3) & "Año de Retencion " & Ret_ano
    Flog.writeline Espacios(Tabulador * 3) & "Fin mes de Retencion " & fin_mes_ret
    
    Flog.writeline Espacios(Tabulador * 3) & "Máxima Ret. en % " & Tope_Gral
    Flog.writeline Espacios(Tabulador * 3) & "Neto del Mes " & Neto
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Tope_Gral, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Neto, "Neto del Mes", Neto)
End If


'Limpiar items que suman al articulo 23
For i = 1 To 50
    Items_ART_23(i) = False
Next i



' Recorro todos los items de Ganancias
StrSql = "SELECT * FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_item

Do While Not rs_item.EOF
    
    Select Case rs_item!itetipotope
    Case 1: ' el valor a tomar es lo que dice la escala
    
        StrSql = "SELECT * FROM valitem WHERE viano =" & Ret_ano & _
                 " AND itenro=" & rs_item!itenro & _
                 " AND vimes =" & Ret_mes
        OpenRecordset StrSql, rs_valitem
        
        Do While Not rs_valitem.EOF
            Items_DDJJ(rs_valitem!itenro) = rs_valitem!vimonto
            Items_TOPE(rs_valitem!itenro) = rs_valitem!vimonto
            
            rs_valitem.MoveNext
        Loop
    ' End case 1
    ' ------------------------------------------------------------------------
    
    Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
        ' Busco la declaracion jurada
        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!ternro & _
                 " AND desano=" & Ret_ano & _
                 " AND itenro = " & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
        
        Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_item!itenro = 3 Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                    
                    'FGZ - 19/04/2004
                    If rs_item!itenro <= 4 Then
                        If Not EsNulo(rs_desmen!descuit) Then
                            i = 11
                            Distinto = rs_traza_gan!Cuit_Entidad11 <> rs_desmen!descuit
                            Do While (i <= Tope) And Distinto
                                i = i + 1
                                Select Case i
                                Case 11:
                                    Distinto = rs_traza_gan!Cuit_Entidad11 <> rs_desmen!descuit
                                Case 12:
                                    Distinto = rs_traza_gan!CUIT_Entidad12 <> rs_desmen!descuit
                                Case 13:
                                    Distinto = rs_traza_gan!CUIT_Entidad13 <> rs_desmen!descuit
                                Case 14:
                                    Distinto = rs_traza_gan!CUIT_Entidad14 <> rs_desmen!descuit
                                End Select
                            Loop
                          
                            If i > Tope And i <= 14 Then
                                StrSql = "UPDATE traza_gan SET "
                                StrSql = StrSql & " cuit_entidad" & i & "='" & rs_desmen!descuit & "',"
                                StrSql = StrSql & " entidad" & i & "='" & rs_desmen!DesRazsoc & "',"
                                StrSql = StrSql & " monto_entidad" & i & "=" & rs_desmen!desmondec
                                StrSql = StrSql & " WHERE "
                                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'FGZ - 22/12/2004
                                'Leo la tabla
                                StrSql = "SELECT * FROM traza_gan WHERE "
                                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
                                OpenRecordset StrSql, rs_traza_gan
                                
                                
                                Tope = Tope + 1
                            Else
                                If i = 15 Then
                                    Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!empleg ' empleado.empleg
                                Else
                                    StrSql = "UPDATE traza_gan SET "
                                    StrSql = StrSql & " monto_entidad" & i & "= monto_entidad" & i & " + " & rs_desmen!desmondec
                                    StrSql = StrSql & " WHERE "
                                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    'FGZ - 22/12/2004
                                    'Leo la tabla
                                    StrSql = "SELECT * FROM traza_gan WHERE "
                                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                    If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
                                    OpenRecordset StrSql, rs_traza_gan
                                End If
                            End If
                        Else
                            Total_Empresa = Total_Empresa + rs_desmen!desmondec
                        End If
                    End If
                    'FGZ - 19/04/2004
                    
                Else
                    If rs_desmen!desmenprorra = 0 Then 'no es parejito
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec
                    Else
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                    End If
                End If
            End If
            
            
            rs_desmen.MoveNext
        Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto
            'Si el desliq prorratea debo proporcionarlo
            Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 Or CBool(rs_desliq!dlprorratea)) And (prorratea = 1), rs_desliq!dlmonto / (13 - Month(rs_desliq!dlfecha)) * (Ret_mes - Month(rs_desliq!dlfecha) + 1), rs_desliq!dlmonto)

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itaprorratea) Then
                    If CBool(rs_itemacum!itasigno) Then
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    End If
                Else
                    If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    End If
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) Then
                If CBool(rs_itemconc!itcsigno) Then
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                End If
            Else
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                End If
            End If
        
            rs_itemconc.MoveNext
        Loop
    ' End case 2
    ' ------------------------------------------------------------------------
        
' ****************************************************************************
' *  OJO QUEDA PENDIENTE EL PRORRATEO PARA LOS ITEMS DE TIPO 3 Y 5           *
' ****************************************************************************


     Case 3: 'TOMO LOS VALORES DE LA DDJJ Y LIQUIDACION Y EL TOPE PARA APLICARLO
     
        StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                 " AND vimes = " & Ret_mes & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_valitem
         Do While Not rs_valitem.EOF
            Items_TOPE(rs_item!itenro) = rs_valitem!vimonto
         
            rs_valitem.MoveNext
         Loop
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
         
            rs_desmen.MoveNext
         Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
        
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_item!itesigno) Then
            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
        End If
        
    ' End case 3
    ' ------------------------------------------------------------------------
    Case 4:
        ' Tomo los valores de la DDJJ y el valor de la escala (cargas de familia)
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfechas) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Month(rs_desmen!desfechas) - Month(rs_desmen!desfecdes) + 1)
            Else
                If Month(rs_desmen!desfecdes) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
                End If
            End If
        
            rs_desmen.MoveNext
         Loop
        
        If Items_DDJJ(rs_item!itenro) > 0 Then
            StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                     " AND vimes = " & Ret_mes & _
                     " AND itenro =" & rs_item!itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_item!itenro) = rs_valitem!vimonto / Ret_mes * Items_DDJJ(rs_item!itenro)
             
                rs_valitem.MoveNext
             Loop
        End If
    ' End case 4
    ' ------------------------------------------------------------------------
        
    Case 5:
        i = 1
        j = 1
        'Hasta = IIf(50 > Len(rs_item!iteitemstope), 50, rs_item!iteitemstope)
        Hasta = 50
        Terminar = False
        Do While j <= Hasta And Not Terminar
            pos1 = i
            pos2 = InStr(i, rs_item!iteitemstope, ",") - 1
            If pos2 > 0 Then
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
            Else
                pos2 = Len(rs_item!iteitemstope)
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
                Terminar = True
            End If
            
            If Texto <> "" Then
                If Mid(Texto, 1, 1) = "-" Then
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                Else
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                End If
            End If
            i = pos2 + 2
            j = j + 1
        Loop
        
        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) * rs_item!iteporctope / 100
    
    
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
            ' Tocado por Maxi 26/05/2004 faltaba el parejito
            'If Month(rs_desmen!desfecdes) <= Ret_mes Then
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
            'Else
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
            'End If
         
            ' FGZ - 19/04/2004
            If rs_item!itenro = 20 Then 'Honorarios medicos
                If Not EsNulo(rs_desmen!descuit) Then
                    StrSql = "UPDATE traza_gan SET "
                    StrSql = StrSql & " cuit_entidad9 ='" & rs_desmen!descuit & "',"
                    StrSql = StrSql & " entidad9='" & rs_desmen!DesRazsoc & "',"
                    StrSql = StrSql & " monto_entidad9=" & rs_desmen!desmondec
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'FGZ - 22/12/2004
                    'Leo la tabla
                    StrSql = "SELECT * FROM traza_gan WHERE "
                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                    If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
                    OpenRecordset StrSql, rs_traza_gan
                    
                    Tope = Tope + 1
                End If
            End If
            If rs_item!itenro = 22 Then 'Impuesto al debito bancario
                StrSql = "UPDATE traza_gan SET "
                StrSql = StrSql & " promo =" & rs_desmen!desmondec
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            
                'FGZ - 22/12/2004
                'Leo la tabla
                StrSql = "SELECT * FROM traza_gan WHERE "
                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
                OpenRecordset StrSql, rs_traza_gan
            End If
            ' FGZ - 19/04/2004
            
            rs_desmen.MoveNext
         Loop
    
    
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro=" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
' FGZ - 12/02/2004 Desde acá -------------------------
'        StrSql = "SELECT * FROM itemacum " & _
'                 " INNER JOIN acu_liq ON itemacum.acunro = acu_liq.acunro " & _
'                 " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND itenro =" & rs_item!itenro & _
'                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemacum
'
'        Do While Not rs_itemacum.EOF
'                If CBool(rs_itemacum!itasigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
'                End If
'
'            rs_itemacum.MoveNext
'        Loop
' FGZ - 12/02/2004 Hasta acá -------------------------

        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
' FGZ - 22/06/2004
'        'TOPEO LOS VALORES
'        If Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro) < Items_TOPE(rs_item!itenro) Then
'            Items_TOPE(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro)
'        End If

' FGZ - 22/06/2004
' puse lo mismo que para el itemtope 3
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_item!itesigno) Then
            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
        End If

    ' End case 5
    ' ------------------------------------------------------------------------
    Case Else:
    End Select
    
    
    ' SI SE TOMA PARA LA GANANCIA NETA, DA VUELTA EL SIGNO DEL TOPE SOLO PARA ITEMS
    ' QUE SE TOPEAN DE ALGUNA FORMA Y NO SALEN DEL RECIBO DE SUELDO.
    ' "Como saber que no sale del Recibo" ?
    
    If rs_item!itenro > 7 Then
        Items_TOPE(rs_item!itenro) = IIf(CBool(rs_item!itesigno), Items_TOPE(rs_item!itenro), Abs(Items_TOPE(rs_item!itenro)))
    End If
    
    
    'Armo la traza del item
    If CBool(USA_DEBUG) Then
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-DDJJ" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_DDJJ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Liq" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_LIQ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-LiqAnt" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_OLD_LIQ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Prorr" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_PRORR(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Tope" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_TOPE(rs_item!itenro)
    End If
    If HACE_TRAZA Then
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_DDJJ(rs_item!itenro))
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_LIQ(rs_item!itenro))
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_OLD_LIQ(rs_item!itenro))
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_PRORR(rs_item!itenro))
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_TOPE(rs_item!itenro))
    End If
        
    
    'Calcula la Ganancia Imponible
    If CBool(rs_item!itesigno) Then
        Gan_Imponible = Gan_Imponible + Items_TOPE(rs_item!itenro)
    Else
        If (rs_item!itetipotope = 1) Or (rs_item!itetipotope = 4) Then
            Ded_a23 = Ded_a23 - Items_TOPE(rs_item!itenro)
            Items_ART_23(rs_item!itenro) = True
        Else
            Deducciones = Deducciones - Items_TOPE(rs_item!itenro)
        End If
    End If
            
    rs_item.MoveNext
Loop
            
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Neta: " & Gan_Imponible
        Flog.writeline Espacios(Tabulador * 3) & "9- Total Deducciones: " & Deducciones
        Flog.writeline Espacios(Tabulador * 3) & "9- Total art. 23: " & Ded_a23
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Neta ", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Total art. 23", Ded_a23)
    End If
            
    
    ' Calculo el porcentaje de deduccion segun la ganancia neta
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia neta para el 648
    StrSql = "UPDATE traza_gan SET "
    StrSql = StrSql & " ganneta =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM traza_gan WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
    OpenRecordset StrSql, rs_traza_gan
    
    If Ret_ano >= 2000 And Gan_Imponible > 0 Then
        StrSql = "SELECT * FROM escala_ded " & _
                 " WHERE esd_topeinf <= " & ((Gan_Imponible + Deducciones) / Ret_mes * 12) & _
                 " AND esd_topesup >=" & ((Gan_Imponible + Deducciones) / Ret_mes * 12)
        OpenRecordset StrSql, rs_escala_ded
    
        If Not rs_escala_ded.EOF Then
            Por_Deduccion = rs_escala_ded!esd_porcentaje
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "No hay esc. dedu para" & Gan_Imponible
            End If
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "No hay esc. dedu para", Gan_Imponible)
            End If
            ' No se ha encontrado la escala de deduccion para el valor gan_imponible
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- % a tomar deduc." & Por_Deduccion
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- % a tomar deduc.", Por_Deduccion)
        End If
        
        'Aplico el porcentaje a las deducciones
        Ded_a23 = Ded_a23 * Por_Deduccion / 100
    End If
            
    
    ' Calculo la Ganancia imponible
    Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia imponible para el 648
    StrSql = "UPDATE traza_gan SET "
    StrSql = StrSql & " ganimpo =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
    OpenRecordset StrSql, rs_traza_gan
    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible" & Gan_Imponible
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Imponible", Gan_Imponible)
    End If
            
                
    If Gan_Imponible > 0 Then
        'Entrar en la escala con las ganancias acumuladas
        StrSql = "SELECT * FROM escala " & _
                 " WHERE escmes =" & Ret_mes & _
                 " AND escano =" & Ret_ano & _
                 " AND escinf <= " & Gan_Imponible & _
                 " AND escsup >= " & Gan_Imponible
        OpenRecordset StrSql, rs_escala
        
        If Not rs_escala.EOF Then
            Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
        Else
            Impuesto_Escala = 0
        End If
    Else
        Impuesto_Escala = 0
    End If
            
            
    ' FGZ - 19/04/2004
    Otros = 0
    i = 18
    
    Do While i <= 50
        Otros = Otros + Abs(Items_TOPE(i))
        i = i + 1
    Loop
    
'    'FGZ - 18/04/2005
'    'antes de esto ya tienen que tener aplicado el % todos los items del art 23
'    For i = 1 To 50
'        If Items_ART_23(i) Then
'            If Por_Deduccion <> 0 Then
'                Items_TOPE(i) = Items_TOPE(i) * Por_Deduccion / 100
'            End If
'        End If
'    Next i
'
    
    StrSql = "UPDATE traza_gan SET "
    StrSql = StrSql & "  msr =" & Items_TOPE(1) + Items_TOPE(3) + Items_TOPE(4)
    StrSql = StrSql & ", nomsr =" & Items_TOPE(2)
    'StrSql = StrSql & ", nogan = 0"
    StrSql = StrSql & ", conyuge =" & Items_TOPE(10)
    StrSql = StrSql & ", hijo =" & Items_TOPE(11)
    StrSql = StrSql & ", otras_cargas =" & Items_TOPE(12)
    StrSql = StrSql & ", car_flia =" & Items_TOPE(10) + Items_TOPE(11) + Items_TOPE(12)
    StrSql = StrSql & ", prima_seguro =" & Abs(Items_TOPE(8))
    StrSql = StrSql & ", sepelio =" & Abs(Items_TOPE(9))
    StrSql = StrSql & ", osocial =" & -Items_TOPE(6)
    StrSql = StrSql & ", cuota_medico =" & Abs(Items_TOPE(13))
    StrSql = StrSql & ", jubilacion =" & -(Items_TOPE(5))
    StrSql = StrSql & ", sindicato =" & -(Items_TOPE(7))
    StrSql = StrSql & ", donacion =" & Abs(Items_TOPE(15))
    StrSql = StrSql & ", otras =" & Otros
    StrSql = StrSql & ", dedesp =" & (Items_TOPE(16))
    StrSql = StrSql & ", noimpo =" & (Items_TOPE(17))
    StrSql = StrSql & ", seguro_retiro =" & Abs(Items_TOPE(14))
    StrSql = StrSql & ", amortizacion =" & Total_Empresa
    StrSql = StrSql & ", viaticos = 0"
    StrSql = StrSql & ", imp_deter =" & Impuesto_Escala
    StrSql = StrSql & ", saldo =" & Abs(Items_TOPE(14))
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
'    'Determinar el saldo
'    StrSql = "SELECT * FROM traza_gan WHERE "
'    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
'    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
'    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
'    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'    StrSql = StrSql & " AND empresa =" & NroEmp
'    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
'    OpenRecordset StrSql, rs_traza_gan
'
'    If Not rs_traza_gan.EOF Then
'        StrSql = "UPDATE traza_gan SET "
'        StrSql = StrSql & "  saldo =" & rs_traza_gan!imp_deter + IIf(EsNulo(rs_traza_gan!retenciones), 0, rs_traza_gan!retenciones) - IIf(EsNulo(rs_traza_gan!promo), 0, rs_traza_gan!promo)
'        StrSql = StrSql & " WHERE "
'        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
'        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
'        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
'        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'        StrSql = StrSql & " AND empresa =" & NroEmp
'        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        'FGZ - 22/12/2004
'        'Leo la tabla
'        StrSql = "SELECT * FROM traza_gan WHERE "
'        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
'        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
'        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
'        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'        StrSql = StrSql & " AND empresa =" & NroEmp
'        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
'        If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
'        OpenRecordset StrSql, rs_traza_gan
'    End If
'    ' FGZ - 19/04/2004
            
                
    ' Calculo las retenciones ya realizadas
    Ret_Ant = 0
        
    'For each ficharet where ficharet.empleado = buliq-empleado.ternro
    '                    And Month(ficharet.fecha) <= ret-mes
    '                    And Year(ficharet.fecha) = ret-ano NO-LOCK:
    '    Assign Ret-ant = Ret-Ant + ficharet.importe.
    'End.
    
    'como no puede utilizar la funcion month() en sql
    'levanto todas las ficharet del tercero y hago la pregunta dentro del loop
    StrSql = "SELECT * FROM ficharet " & _
             " WHERE empleado =" & buliq_empleado!ternro
    OpenRecordset StrSql, rs_ficharet
    
    Do While Not rs_ficharet.EOF
        If (Month(rs_ficharet!Fecha) <= Ret_mes) And (Year(rs_ficharet!Fecha) = Ret_ano) Then
            Ret_Ant = Ret_Ant + rs_ficharet!Importe
        End If
        rs_ficharet.MoveNext
    Loop
    
    
    'Calcular la retencion
    Retencion = Impuesto_Escala - Ret_Ant
    
    'FGZ - 30/12/2004
    'Determinar el saldo
    StrSql = "SELECT * FROM traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    OpenRecordset StrSql, rs_traza_gan
    
    If Not rs_traza_gan.EOF Then
        StrSql = "UPDATE traza_gan SET "
        StrSql = StrSql & "  saldo =" & Retencion
        StrSql = StrSql & "  ,retenciones =" & Ret_Ant
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 22/12/2004
        'Leo la tabla
        StrSql = "SELECT * FROM traza_gan WHERE "
        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
        If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
        OpenRecordset StrSql, rs_traza_gan
    End If
    ' FGZ - 19/04/2004
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Retenciones anteriores " & Ret_Ant
        If Gan_Imponible > 0 Then
            If Not rs_escala.EOF Then
                rs_escala.MoveFirst
                If Not rs_escala.EOF Then
                    Flog.writeline Espacios(Tabulador * 3) & "Escala Impuesto" & rs_escala!escporexe
                    Flog.writeline Espacios(Tabulador * 3) & "Impuesto por escala" & Impuesto_Escala
                    Flog.writeline Espacios(Tabulador * 3) & "A Retener/Devolver" & Retencion
                Else
                    Flog.writeline Espacios(Tabulador * 3) & "Escala Impuesto" & "0"
                    Flog.writeline Espacios(Tabulador * 3) & "Impuesto por escala" & Impuesto_Escala
                    Flog.writeline Espacios(Tabulador * 3) & "A Retener/Devolver" & Retencion
                End If
            End If
        End If
    End If
    
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Retenciones anteriores", Ret_Ant)
        If Gan_Imponible > 0 Then
            If Not rs_escala.EOF Then
                rs_escala.MoveFirst
                If Not rs_escala.EOF Then
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Escala Impuesto", rs_escala!escporexe)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver", Retencion)
                Else
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Escala Impuesto", 0)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver", Retencion)
                End If
            End If
        End If
    End If
    
    ' Verifico si es una devolucion y si devuelve el concepto
    If Devuelve = 0 And Retencion < 0 Then
        Retencion = 0
    End If
    
    ' FGZ - 14/04/2004
    If Retencion <> 0 Then
        ' Verificar que la rtencion no supere el 30% del Neto del Mes
        If Retencion > (Neto * (Tope_Gral / 100)) Then
            Retencion = Neto * (Tope_Gral / 100)
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "A Retener/Devolver, x Tope General " & Retencion
            End If
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver, x Tope General", Retencion)
            End If
        End If
        Monto = -Retencion
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "La Retencion es " & Monto
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "La Retencion es Cero"
        End If
        Monto = 0
    End If
    'Monto = -Retencion
    Bien = True
    
        
    'Retenciones / Devoluciones
    If Retencion <> 0 Then
        Call InsertarFichaRet(buliq_empleado!ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
    End If
    
    ' Grabo todos los items de la liquidacion actual
    i = 1
    Hasta = 50
    Do While i <= Hasta
        'FGZ 23/04/2004
        ' guardo los item_tope para poder usarlo en el reporte de Ganancias
        If Items_TOPE(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope
            
            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET monto =" & Items_TOPE(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
        If Items_LIQ(i) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!ternro & "," & _
                     ConvFecha(buliq_proceso!profecpago) & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_LIQ(i) & "," & _
                     "0," & _
                     i & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        If Items_PRORR(i) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!ternro & "," & _
                     ConvFecha(buliq_proceso!profecpago) & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_PRORR(i) & "," & _
                     "0," & _
                     i & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        
        'FGZ 31/03/2005
        ' guardo los item_ddjj para poder usarlo en el reporte de Ganancias
        If Items_DDJJ(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope

            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,ddjj,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_DDJJ(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET ddjj =" & Items_DDJJ(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope

            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET prorr =" & Items_PRORR(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_old_liq para poder usarlo en el reporte de Ganancias
        If Items_OLD_LIQ(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope

            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET old_liq =" & Items_OLD_LIQ(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_liq para poder usarlo en el reporte de Ganancias
        If Items_LIQ(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope

            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_LIQ(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET liq =" & Items_LIQ(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 31/03/2005
        
        i = i + 1
    Loop

    exito = Bien
    for_Ganancias = Monto
End Function


Public Function for_Ganancias_Schering(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Single, Bien As Boolean) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias. Customizacion para Schering.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.: 15/06/20005
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim p_prorratea As Integer  'Si prorratea o no para liq. finales

'Variables Locales
Dim Devuelve As Single
Dim Tope_Gral As Single
Dim Neto As Single
Dim prorratea As Single
Dim Retencion As Single
Dim Gan_Imponible As Single
Dim Deducciones As Single
Dim Ded_a23 As Single
Dim Por_Deduccion As Single
Dim Impuesto_Escala As Single
Dim Ret_Ant As Single

Dim Ret_mes As Integer
Dim Ret_ano As Integer
Dim fin_mes_ret As Date
Dim ini_anyo_ret As Date
Dim Con_liquid As Integer
Dim i As Integer
Dim j As Integer
Dim Texto As String

'Vectores para manejar el proceso
Dim Items_DDJJ(50) As Single
Dim Items_LIQ(50) As Single
Dim Items_PRORR(50) As Single
Dim Items_OLD_LIQ(50) As Single
Dim Items_TOPE(50) As Single
Dim Items_ART_23(50) As Boolean

'Recorsets Auxiliares
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_item As New ADODB.Recordset
Dim rs_valitem As New ADODB.Recordset
Dim rs_desmen As New ADODB.Recordset
Dim rs_desliq As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_escala_ded As New ADODB.Recordset
Dim rs_escala As New ADODB.Recordset
Dim rs_ficharet As New ADODB.Recordset
Dim rs_traza_gan_items_tope As New ADODB.Recordset

Dim Hasta As Integer

' FGZ - 12/02/2004
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Single
' FGZ - 12/02/2004

' FGZ - 27/02/2004
Dim Terminar As Boolean
Dim pos1
Dim pos2
' FGZ - 27/02/2004

'Comienzo
p_Devuelve = 1001
p_Tope_Gral = 1002
p_Neto = 1003
p_prorratea = 1005

Bien = False


'FGZ - 19/04/2004
Dim Total_Empresa As Single
Dim Tope As Integer
'Dim rs_Rep19 As New ADODB.Recordset
Dim rs_traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Single
Total_Empresa = 0
Tope = 10

' Primero limpio la traza
StrSql = "DELETE FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
objConn.Execute StrSql, , adExecuteNoRecords

' Guardo la traza de Ganancia en traza_gan para utilizarla luego cuando se genere el reporte
StrSql = "INSERT INTO traza_gan (pliqnro,pronro,concnro,fecha_pago,empresa,ternro,empleg) VALUES (" & _
         buliq_periodo!PliqNro & "," & _
         buliq_proceso!pronro & "," & _
         Buliq_Concepto(Concepto_Actual).concnro & "," & _
         ConvFecha(buliq_proceso!profecpago) & "," & _
         NroEmp & "," & _
         buliq_empleado!ternro & "," & _
         buliq_empleado!empleg & _
         ")"
objConn.Execute StrSql, , adExecuteNoRecords

'FGZ - 22/12/2004
'Leo la tabla
StrSql = "SELECT * FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
OpenRecordset StrSql, rs_traza_gan
    

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

Ret_mes = Month(buliq_proceso!profecpago)
Ret_ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_mes = 12, CDate("01/01/" & Ret_ano + 1) - 1, CDate("01/" & Ret_mes + 1 & "/" & Ret_ano) - 1)
ini_anyo_ret = CDate("01/01/" & Ret_ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).concnro

'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case p_Devuelve:
        Devuelve = rs_wf_tpa!Valor
    Case p_Tope_Gral:
        Tope_Gral = rs_wf_tpa!Valor
    Case p_Neto:
        Neto = rs_wf_tpa!Valor
    Case p_prorratea:
        prorratea = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

'Si es una liq. final no prorratea y tomo la escala de diciembre
If prorratea = 0 Then
    Ret_mes = 12
    'FGZ - 27/09/2004
    fin_mes_ret = IIf(Ret_mes = 12, CDate("01/01/" & Ret_ano + 1) - 1, CDate("01/" & Ret_mes + 1 & "/" & Ret_ano) - 1)
End If

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 3) & "---------Formula-----------------------------"
    Flog.writeline Espacios(Tabulador * 3) & "Fecha del Proceso " & buliq_proceso!profecpago
    Flog.writeline Espacios(Tabulador * 3) & "Mes de Retencion " & Ret_mes
    Flog.writeline Espacios(Tabulador * 3) & "Año de Retencion " & Ret_ano
    Flog.writeline Espacios(Tabulador * 3) & "Fin mes de Retencion " & fin_mes_ret
    
    Flog.writeline Espacios(Tabulador * 3) & "Máxima Ret. en % " & Tope_Gral
    Flog.writeline Espacios(Tabulador * 3) & "Neto del Mes " & Neto
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Tope_Gral, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Neto, "Neto del Mes", Neto)
End If


'Limpiar items que suman al articulo 23
For i = 1 To 50
    Items_ART_23(i) = False
Next i



' Recorro todos los items de Ganancias
StrSql = "SELECT * FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_item

Do While Not rs_item.EOF
    
    Select Case rs_item!itetipotope
    Case 1: ' el valor a tomar es lo que dice la escala
    
        StrSql = "SELECT * FROM valitem WHERE viano =" & Ret_ano & _
                 " AND itenro=" & rs_item!itenro & _
                 " AND vimes =" & Ret_mes
        OpenRecordset StrSql, rs_valitem
        
        Do While Not rs_valitem.EOF
            Items_DDJJ(rs_valitem!itenro) = rs_valitem!vimonto
            Items_TOPE(rs_valitem!itenro) = rs_valitem!vimonto
            
            rs_valitem.MoveNext
        Loop
    ' End case 1
    ' ------------------------------------------------------------------------
    
    Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
        ' Busco la declaracion jurada
        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!ternro & _
                 " AND desano=" & Ret_ano & _
                 " AND itenro = " & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
        
        Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_item!itenro = 3 Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / 12 * Ret_mes, rs_desmen!desmondec)
                    
                    'FGZ - 19/04/2004
                    If rs_item!itenro <= 4 Then
                        If Not EsNulo(rs_desmen!descuit) Then
                            i = 11
                            Distinto = rs_traza_gan!Cuit_Entidad11 <> rs_desmen!descuit
                            Do While (i <= Tope) And Distinto
                                i = i + 1
                                Select Case i
                                Case 11:
                                    Distinto = rs_traza_gan!Cuit_Entidad11 <> rs_desmen!descuit
                                Case 12:
                                    Distinto = rs_traza_gan!CUIT_Entidad12 <> rs_desmen!descuit
                                Case 13:
                                    Distinto = rs_traza_gan!CUIT_Entidad13 <> rs_desmen!descuit
                                Case 14:
                                    Distinto = rs_traza_gan!CUIT_Entidad14 <> rs_desmen!descuit
                                End Select
                            Loop
                          
                            If i > Tope And i <= 14 Then
                                StrSql = "UPDATE traza_gan SET "
                                StrSql = StrSql & " cuit_entidad" & i & "='" & rs_desmen!descuit & "',"
                                StrSql = StrSql & " entidad" & i & "='" & rs_desmen!DesRazsoc & "',"
                                StrSql = StrSql & " monto_entidad" & i & "=" & rs_desmen!desmondec
                                StrSql = StrSql & " WHERE "
                                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'FGZ - 22/12/2004
                                'Leo la tabla
                                StrSql = "SELECT * FROM traza_gan WHERE "
                                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
                                OpenRecordset StrSql, rs_traza_gan
                                
                                
                                Tope = Tope + 1
                            Else
                                If i = 15 Then
                                    Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!empleg ' empleado.empleg
                                Else
                                    StrSql = "UPDATE traza_gan SET "
                                    StrSql = StrSql & " monto_entidad" & i & "= monto_entidad" & i & " + " & rs_desmen!desmondec
                                    StrSql = StrSql & " WHERE "
                                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    'FGZ - 22/12/2004
                                    'Leo la tabla
                                    StrSql = "SELECT * FROM traza_gan WHERE "
                                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                    If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
                                    OpenRecordset StrSql, rs_traza_gan
                                End If
                            End If
                        Else
                            Total_Empresa = Total_Empresa + rs_desmen!desmondec
                        End If
                    End If
                    'FGZ - 19/04/2004
                    
                Else
                    If rs_desmen!desmenprorra = 0 Then 'no es parejito
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec
                    Else
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / 12 * Ret_mes, rs_desmen!desmondec)
                    End If
                End If
            End If
            
            
            rs_desmen.MoveNext
        Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto
            'Si el desliq prorratea debo proporcionarlo
            'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 Or CBool(rs_desliq!dlprorratea)) And (prorratea = 1), rs_desliq!dlmonto / (13 - Month(rs_desliq!dlfecha)) * (Ret_mes - Month(rs_desliq!dlfecha) + 1), rs_desliq!dlmonto)
            Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 Or CBool(rs_desliq!dlprorratea)) And (prorratea = 1), rs_desliq!dlmonto / 12 * Ret_mes, rs_desliq!dlmonto)
            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itaprorratea) Then
                    If CBool(rs_itemacum!itasigno) Then
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + Aux_Acu_Monto
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (12 * Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - Aux_Acu_Monto
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (12 * Ret_mes), Aux_Acu_Monto)
                    End If
                Else
                    If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (12 * Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (12 * Ret_mes), Aux_Acu_Monto)
                    End If
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) Then
                If CBool(rs_itemconc!itcsigno) Then
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + rs_itemconc!dlimonto
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (12 * Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - rs_itemconc!dlimonto
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (12 * Ret_mes), rs_itemconc!dlimonto)
                End If
            Else
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (12 * Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (12 * Ret_mes), rs_itemconc!dlimonto)
                End If
            End If
        
            rs_itemconc.MoveNext
        Loop
    ' End case 2
    ' ------------------------------------------------------------------------
        
' ****************************************************************************
' *  OJO QUEDA PENDIENTE EL PRORRATEO PARA LOS ITEMS DE TIPO 3 Y 5           *
' ****************************************************************************


     Case 3: 'TOMO LOS VALORES DE LA DDJJ Y LIQUIDACION Y EL TOPE PARA APLICARLO
     
        StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                 " AND vimes = " & Ret_mes & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_valitem
         Do While Not rs_valitem.EOF
            Items_TOPE(rs_item!itenro) = rs_valitem!vimonto
         
            rs_valitem.MoveNext
         Loop
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
         
            rs_desmen.MoveNext
         Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
        
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_item!itesigno) Then
            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
        End If
        
    ' End case 3
    ' ------------------------------------------------------------------------
    Case 4:
        ' Tomo los valores de la DDJJ y el valor de la escala (cargas de familia)
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfechas) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Month(rs_desmen!desfechas) - Month(rs_desmen!desfecdes) + 1)
            Else
                If Month(rs_desmen!desfecdes) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
                End If
            End If
        
            rs_desmen.MoveNext
         Loop
        
        If Items_DDJJ(rs_item!itenro) > 0 Then
            StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                     " AND vimes = " & Ret_mes & _
                     " AND itenro =" & rs_item!itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_item!itenro) = rs_valitem!vimonto / Ret_mes * Items_DDJJ(rs_item!itenro)
             
                rs_valitem.MoveNext
             Loop
        End If
    ' End case 4
    ' ------------------------------------------------------------------------
        
    Case 5:
        i = 1
        j = 1
        'Hasta = IIf(50 > Len(rs_item!iteitemstope), 50, rs_item!iteitemstope)
        Hasta = 50
        Terminar = False
        Do While j <= Hasta And Not Terminar
            pos1 = i
            pos2 = InStr(i, rs_item!iteitemstope, ",") - 1
            If pos2 > 0 Then
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
            Else
                pos2 = Len(rs_item!iteitemstope)
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
                Terminar = True
            End If
            
            If Texto <> "" Then
                If Mid(Texto, 1, 1) = "-" Then
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                Else
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                End If
            End If
            i = pos2 + 2
            j = j + 1
        Loop
        
        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) * rs_item!iteporctope / 100
    
    
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
            ' Tocado por Maxi 26/05/2004 faltaba el parejito
            'If Month(rs_desmen!desfecdes) <= Ret_mes Then
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
            'Else
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
            'End If
         
            ' FGZ - 19/04/2004
            If rs_item!itenro = 20 Then 'Honorarios medicos
                If Not EsNulo(rs_desmen!descuit) Then
                    StrSql = "UPDATE traza_gan SET "
                    StrSql = StrSql & " cuit_entidad9 ='" & rs_desmen!descuit & "',"
                    StrSql = StrSql & " entidad9='" & rs_desmen!DesRazsoc & "',"
                    StrSql = StrSql & " monto_entidad9=" & rs_desmen!desmondec
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'FGZ - 22/12/2004
                    'Leo la tabla
                    StrSql = "SELECT * FROM traza_gan WHERE "
                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                    If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
                    OpenRecordset StrSql, rs_traza_gan
                    
                    Tope = Tope + 1
                End If
            End If
            If rs_item!itenro = 22 Then 'Impuesto al debito bancario
                StrSql = "UPDATE traza_gan SET "
                StrSql = StrSql & " promo =" & rs_desmen!desmondec
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            
                'FGZ - 22/12/2004
                'Leo la tabla
                StrSql = "SELECT * FROM traza_gan WHERE "
                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
                OpenRecordset StrSql, rs_traza_gan
            End If
            ' FGZ - 19/04/2004
            
            rs_desmen.MoveNext
         Loop
    
    
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro=" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
' FGZ - 12/02/2004 Desde acá -------------------------
'        StrSql = "SELECT * FROM itemacum " & _
'                 " INNER JOIN acu_liq ON itemacum.acunro = acu_liq.acunro " & _
'                 " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND itenro =" & rs_item!itenro & _
'                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemacum
'
'        Do While Not rs_itemacum.EOF
'                If CBool(rs_itemacum!itasigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
'                End If
'
'            rs_itemacum.MoveNext
'        Loop
' FGZ - 12/02/2004 Hasta acá -------------------------

        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
' FGZ - 22/06/2004
'        'TOPEO LOS VALORES
'        If Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro) < Items_TOPE(rs_item!itenro) Then
'            Items_TOPE(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro)
'        End If

' FGZ - 22/06/2004
' puse lo mismo que para el itemtope 3
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_item!itesigno) Then
            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
        End If

    ' End case 5
    ' ------------------------------------------------------------------------
    Case Else:
    End Select
    
    
    ' SI SE TOMA PARA LA GANANCIA NETA, DA VUELTA EL SIGNO DEL TOPE SOLO PARA ITEMS
    ' QUE SE TOPEAN DE ALGUNA FORMA Y NO SALEN DEL RECIBO DE SUELDO.
    ' "Como saber que no sale del Recibo" ?
    
    If rs_item!itenro > 7 Then
        Items_TOPE(rs_item!itenro) = IIf(CBool(rs_item!itesigno), Items_TOPE(rs_item!itenro), Abs(Items_TOPE(rs_item!itenro)))
    End If
    
    
    'Armo la traza del item
    If CBool(USA_DEBUG) Then
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-DDJJ" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_DDJJ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Liq" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_LIQ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-LiqAnt" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_OLD_LIQ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Prorr" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_PRORR(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Tope" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_TOPE(rs_item!itenro)
    End If
    If HACE_TRAZA Then
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_DDJJ(rs_item!itenro))
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_LIQ(rs_item!itenro))
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_OLD_LIQ(rs_item!itenro))
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_PRORR(rs_item!itenro))
        Texto = Format(CStr(rs_item!itenro), "00") & "-" & rs_item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_TOPE(rs_item!itenro))
    End If
        
    
    'Calcula la Ganancia Imponible
    If CBool(rs_item!itesigno) Then
        Gan_Imponible = Gan_Imponible + Items_TOPE(rs_item!itenro)
    Else
        If (rs_item!itetipotope = 1) Or (rs_item!itetipotope = 4) Then
            Ded_a23 = Ded_a23 - Items_TOPE(rs_item!itenro)
            Items_ART_23(rs_item!itenro) = True
        Else
            Deducciones = Deducciones - Items_TOPE(rs_item!itenro)
        End If
    End If
            
    rs_item.MoveNext
Loop
            
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Neta: " & Gan_Imponible
        Flog.writeline Espacios(Tabulador * 3) & "9- Total Deducciones: " & Deducciones
        Flog.writeline Espacios(Tabulador * 3) & "9- Total art. 23: " & Ded_a23
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Neta ", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Total art. 23", Ded_a23)
    End If
            
    
    ' Calculo el porcentaje de deduccion segun la ganancia neta
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia neta para el 648
    StrSql = "UPDATE traza_gan SET "
    StrSql = StrSql & " ganneta =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM traza_gan WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
    OpenRecordset StrSql, rs_traza_gan
    
    If Ret_ano >= 2000 And Gan_Imponible > 0 Then
        StrSql = "SELECT * FROM escala_ded " & _
                 " WHERE esd_topeinf <= " & ((Gan_Imponible + Deducciones) / Ret_mes * 12) & _
                 " AND esd_topesup >=" & ((Gan_Imponible + Deducciones) / Ret_mes * 12)
        OpenRecordset StrSql, rs_escala_ded
    
        If Not rs_escala_ded.EOF Then
            Por_Deduccion = rs_escala_ded!esd_porcentaje
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "No hay esc. dedu para" & Gan_Imponible
            End If
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "No hay esc. dedu para", Gan_Imponible)
            End If
            ' No se ha encontrado la escala de deduccion para el valor gan_imponible
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- % a tomar deduc." & Por_Deduccion
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- % a tomar deduc.", Por_Deduccion)
        End If
        
        'Aplico el porcentaje a las deducciones
        Ded_a23 = Ded_a23 * Por_Deduccion / 100
    End If
            
    
    ' Calculo la Ganancia imponible
    Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia imponible para el 648
    StrSql = "UPDATE traza_gan SET "
    StrSql = StrSql & " ganimpo =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
    OpenRecordset StrSql, rs_traza_gan
    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible" & Gan_Imponible
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Imponible", Gan_Imponible)
    End If
            
                
    If Gan_Imponible > 0 Then
        'Entrar en la escala con las ganancias acumuladas
        StrSql = "SELECT * FROM escala " & _
                 " WHERE escmes =" & Ret_mes & _
                 " AND escano =" & Ret_ano & _
                 " AND escinf <= " & Gan_Imponible & _
                 " AND escsup >= " & Gan_Imponible
        OpenRecordset StrSql, rs_escala
        
        If Not rs_escala.EOF Then
            Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
        Else
            Impuesto_Escala = 0
        End If
    Else
        Impuesto_Escala = 0
    End If
            
            
    ' FGZ - 19/04/2004
    Otros = 0
    i = 18
    
    Do While i <= 50
        Otros = Otros + Abs(Items_TOPE(i))
        i = i + 1
    Loop
    
'    'FGZ - 18/04/2005
'    'antes de esto ya tienen que tener aplicado el % todos los items del art 23
'    For i = 1 To 50
'        If Items_ART_23(i) Then
'            If Por_Deduccion <> 0 Then
'                Items_TOPE(i) = Items_TOPE(i) * Por_Deduccion / 100
'            End If
'        End If
'    Next i
'
    
    StrSql = "UPDATE traza_gan SET "
    StrSql = StrSql & "  msr =" & Items_TOPE(1) + Items_TOPE(3) + Items_TOPE(4)
    StrSql = StrSql & ", nomsr =" & Items_TOPE(2)
    'StrSql = StrSql & ", nogan = 0"
    StrSql = StrSql & ", conyuge =" & Items_TOPE(10)
    StrSql = StrSql & ", hijo =" & Items_TOPE(11)
    StrSql = StrSql & ", otras_cargas =" & Items_TOPE(12)
    StrSql = StrSql & ", car_flia =" & Items_TOPE(10) + Items_TOPE(11) + Items_TOPE(12)
    StrSql = StrSql & ", prima_seguro =" & Abs(Items_TOPE(8))
    StrSql = StrSql & ", sepelio =" & Abs(Items_TOPE(9))
    StrSql = StrSql & ", osocial =" & -Items_TOPE(6)
    StrSql = StrSql & ", cuota_medico =" & Abs(Items_TOPE(13))
    StrSql = StrSql & ", jubilacion =" & -(Items_TOPE(5))
    StrSql = StrSql & ", sindicato =" & -(Items_TOPE(7))
    StrSql = StrSql & ", donacion =" & Abs(Items_TOPE(15))
    StrSql = StrSql & ", otras =" & Otros
    StrSql = StrSql & ", dedesp =" & (Items_TOPE(16))
    StrSql = StrSql & ", noimpo =" & (Items_TOPE(17))
    StrSql = StrSql & ", seguro_retiro =" & Abs(Items_TOPE(14))
    StrSql = StrSql & ", amortizacion =" & Total_Empresa
    StrSql = StrSql & ", viaticos = 0"
    StrSql = StrSql & ", imp_deter =" & Impuesto_Escala
    StrSql = StrSql & ", saldo =" & Abs(Items_TOPE(14))
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
'    'Determinar el saldo
'    StrSql = "SELECT * FROM traza_gan WHERE "
'    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
'    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
'    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
'    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'    StrSql = StrSql & " AND empresa =" & NroEmp
'    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
'    OpenRecordset StrSql, rs_traza_gan
'
'    If Not rs_traza_gan.EOF Then
'        StrSql = "UPDATE traza_gan SET "
'        StrSql = StrSql & "  saldo =" & rs_traza_gan!imp_deter + IIf(EsNulo(rs_traza_gan!retenciones), 0, rs_traza_gan!retenciones) - IIf(EsNulo(rs_traza_gan!promo), 0, rs_traza_gan!promo)
'        StrSql = StrSql & " WHERE "
'        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
'        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
'        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
'        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'        StrSql = StrSql & " AND empresa =" & NroEmp
'        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        'FGZ - 22/12/2004
'        'Leo la tabla
'        StrSql = "SELECT * FROM traza_gan WHERE "
'        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
'        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
'        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
'        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'        StrSql = StrSql & " AND empresa =" & NroEmp
'        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
'        If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
'        OpenRecordset StrSql, rs_traza_gan
'    End If
'    ' FGZ - 19/04/2004
            
                
    ' Calculo las retenciones ya realizadas
    Ret_Ant = 0
        
    'For each ficharet where ficharet.empleado = buliq-empleado.ternro
    '                    And Month(ficharet.fecha) <= ret-mes
    '                    And Year(ficharet.fecha) = ret-ano NO-LOCK:
    '    Assign Ret-ant = Ret-Ant + ficharet.importe.
    'End.
    
    'como no puede utilizar la funcion month() en sql
    'levanto todas las ficharet del tercero y hago la pregunta dentro del loop
    StrSql = "SELECT * FROM ficharet " & _
             " WHERE empleado =" & buliq_empleado!ternro
    OpenRecordset StrSql, rs_ficharet
    
    Do While Not rs_ficharet.EOF
        If (Month(rs_ficharet!Fecha) <= Ret_mes) And (Year(rs_ficharet!Fecha) = Ret_ano) Then
            Ret_Ant = Ret_Ant + rs_ficharet!Importe
        End If
        rs_ficharet.MoveNext
    Loop
    
    
    'Calcular la retencion
    Retencion = Impuesto_Escala - Ret_Ant
    
    'FGZ - 30/12/2004
    'Determinar el saldo
    StrSql = "SELECT * FROM traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    OpenRecordset StrSql, rs_traza_gan
    
    If Not rs_traza_gan.EOF Then
        StrSql = "UPDATE traza_gan SET "
        StrSql = StrSql & "  saldo =" & Retencion
        StrSql = StrSql & "  ,retenciones =" & Ret_Ant
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 22/12/2004
        'Leo la tabla
        StrSql = "SELECT * FROM traza_gan WHERE "
        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
        If rs_traza_gan.State = adStateOpen Then rs_traza_gan.Close
        OpenRecordset StrSql, rs_traza_gan
    End If
    ' FGZ - 19/04/2004
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Retenciones anteriores " & Ret_Ant
        If Gan_Imponible > 0 Then
            If Not rs_escala.EOF Then
                rs_escala.MoveFirst
                If Not rs_escala.EOF Then
                    Flog.writeline Espacios(Tabulador * 3) & "Escala Impuesto" & rs_escala!escporexe
                    Flog.writeline Espacios(Tabulador * 3) & "Impuesto por escala" & Impuesto_Escala
                    Flog.writeline Espacios(Tabulador * 3) & "A Retener/Devolver" & Retencion
                Else
                    Flog.writeline Espacios(Tabulador * 3) & "Escala Impuesto" & "0"
                    Flog.writeline Espacios(Tabulador * 3) & "Impuesto por escala" & Impuesto_Escala
                    Flog.writeline Espacios(Tabulador * 3) & "A Retener/Devolver" & Retencion
                End If
            End If
        End If
    End If
    
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Retenciones anteriores", Ret_Ant)
        If Gan_Imponible > 0 Then
            If Not rs_escala.EOF Then
                rs_escala.MoveFirst
                If Not rs_escala.EOF Then
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Escala Impuesto", rs_escala!escporexe)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver", Retencion)
                Else
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Escala Impuesto", 0)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver", Retencion)
                End If
            End If
        End If
    End If
    
    ' Verifico si es una devolucion y si devuelve el concepto
    If Devuelve = 0 And Retencion < 0 Then
        Retencion = 0
    End If
    
    ' FGZ - 14/04/2004
    If Retencion <> 0 Then
        ' Verificar que la rtencion no supere el 30% del Neto del Mes
        If Retencion > (Neto * (Tope_Gral / 100)) Then
            Retencion = Neto * (Tope_Gral / 100)
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "A Retener/Devolver, x Tope General " & Retencion
            End If
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver, x Tope General", Retencion)
            End If
        End If
        Monto = -Retencion
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "La Retencion es " & Monto
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "La Retencion es Cero"
        End If
        Monto = 0
    End If
    'Monto = -Retencion
    Bien = True
    
        
    'Retenciones / Devoluciones
    If Retencion <> 0 Then
        Call InsertarFichaRet(buliq_empleado!ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
    End If
    
    ' Grabo todos los items de la liquidacion actual
    i = 1
    Hasta = 50
    Do While i <= Hasta
        'FGZ 23/04/2004
        ' guardo los item_tope para poder usarlo en el reporte de Ganancias
        If Items_TOPE(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope
            
            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET monto =" & Items_TOPE(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
        If Items_LIQ(i) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!ternro & "," & _
                     ConvFecha(buliq_proceso!profecpago) & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_LIQ(i) & "," & _
                     "0," & _
                     i & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        If Items_PRORR(i) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!ternro & "," & _
                     ConvFecha(buliq_proceso!profecpago) & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_PRORR(i) & "," & _
                     "0," & _
                     i & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        
        'FGZ 31/03/2005
        ' guardo los item_ddjj para poder usarlo en el reporte de Ganancias
        If Items_DDJJ(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope

            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,ddjj,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_DDJJ(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET ddjj =" & Items_DDJJ(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope

            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET prorr =" & Items_PRORR(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_old_liq para poder usarlo en el reporte de Ganancias
        If Items_OLD_LIQ(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope

            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET old_liq =" & Items_OLD_LIQ(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_liq para poder usarlo en el reporte de Ganancias
        If Items_LIQ(i) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_items_tope "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & i
            OpenRecordset StrSql, rs_traza_gan_items_tope

            If rs_traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_items_tope (ternro,pronro,liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_LIQ(i) & "," & _
                         NroEmp & "," & _
                         i & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_items_tope SET liq =" & Items_LIQ(i)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & i
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 31/03/2005
        
        i = i + 1
    Loop

    exito = Bien
    for_Ganancias_Schering = Monto
End Function


Public Sub InsertarFichaRet(ByVal ternro As Long, ByVal Fecha As Date, ByVal Importe As Single, ByVal pronro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion           : Graba un registro de ficharet (cabliq/fecha/proceso).
' Autor                 : Lic. Mauricio Heidt -
' Fecha                 : 5/2/98
' Traducido por         : FGZ
' Parametros de entrada : Número de tercero
'                       : Fecha de pago
'                       : Importe Importe
'                       : Numero de proceso de liquidacion
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim FechaAux As Date

    If EsNulo(Fecha) Then
        FechaAux = Date
    Else
        FechaAux = Fecha
    End If
    
    StrSql = "INSERT INTO ficharet (empleado,fecha,pronro,importe,liqsistema) VALUES (" & _
             ternro & "," & _
             ConvFecha(FechaAux) & "," & _
             pronro & "," & _
             Importe & "," & _
             "-1" & _
             ")"
    objConn.Execute StrSql, , adExecuteNoRecords


End Sub

Public Function for_Grossing3(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Single, Bien As Boolean) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion           : Programa para el calculo de retencion de ganancia.
' Autor                 :
' Fecha                 :
' Traducido por         : FGZ
' Parametros de entrada :
' Ultima Mod.           :
' Descripcion           :
' ---------------------------------------------------------------------------------------------

' Tipos de parametros usados
Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim prorratea As Single
Dim p_prorratea As Integer  'Si prorratea o no para liq. finales
Dim p_Diciembre As Integer  'Toma la escala de Diciembre
Dim p_Porcgross As Integer  'Porcentaje del grossing a efectuar

Dim p_grossant1 As Integer  'Importe de grossing anterior dentro de la liquidacion
Dim p_grossant2 As Integer  'Importe de grossing anterior dentro de la liquidacion
Dim p_acuimp As Integer  'Acumulador base sobre el que se calculan las cargas
Dim p_porcjub As Integer  'Porcentaje de Jubiliacion
Dim p_porcley As Integer  'Porcentaje de Ley
Dim p_porcosoc As Integer  'Porcentaje de Obra social
Dim p_porcsin As Integer  'Porcentaje de sindicato

'Variables Locales
Dim AMPO As Single
Dim Porcapo As Single
Dim Devuelve As Single
Dim Tope_Gral As Single
Dim Neto As Single
Dim Retencion As Single
Dim Gan_Imponible As Single
Dim Gan_Neta As Single
Dim Deducciones As Single
Dim Ded_a23 As Single
Dim Por_Deduccion As Single
Dim Impuesto_Escala As Single
Dim Ret_Ant As Single
Dim Porcgross As Single
Dim Ajustecargas As Single
Dim Ajustededucc As Single
Dim Topeescala As Single
Dim M2asgross As Single
Dim M3asgross As Single
Dim M4asgross As Single
Dim M5asgross As Single
Dim Masgross As Single
Dim Porcant As Single
Dim Diciembre As Single

Dim grossant1 As Integer  'Importe de grossing anterior dentro de la liquidacion
Dim grossant2 As Integer  'Importe de grossing anterior dentro de la liquidacion
Dim AcuImp As Integer  'Acumulador base sobre el que se calculan las cargas
Dim PorcJub As Integer  'Porcentaje de Jubiliacion
Dim PorcLey As Integer  'Porcentaje de Ley
Dim PorcOSoc As Integer  'Porcentaje de Obra social
Dim PorcSin As Integer  'Porcentaje de sindicato

Dim Ret_mes As Integer
Dim Ret_ano As Integer
Dim Aux_Fin_Ret As Date
Dim Aux_Inicio_Ret As Date
Dim Con_liquid As Integer
Dim i As Integer
Dim Texto As String

Dim baja_ded As Single
Dim baja1 As Single
Dim baja2 As Single

Dim AuxOSPriv As Single
Dim AjuOSPriv As Single

'Vectores para manejar el proceso
Dim Items_DDJJ(50) As Single
Dim Items_LIQ(50) As Single
Dim Items_PRORR(50) As Single
Dim Items_OLD_LIQ(50) As Single
Dim Items_TOPE(50) As Single

'----------------------------------------------------
'Fin calculo tope y aportes para tomar en el grossing
'----------------------------------------------------
Dim Acumulado_Mes As Single
Dim Ampo_A_Aplicar As Single
Dim Contador As Integer
Dim MesSemestreDesde As Integer
Dim MesSemestreHasta As Integer
Dim Valor_Ampo As Single
Dim Valor_Ampo_Cont As Single
Dim Tope_Monto_Proporcional As Single
Dim Monto_Tope As Single
Dim Cant_Ampo_Proporcionar(5) As Single
Dim Cant_Diaria_Ampos(5) As Single
Dim Monto_wf(5) As Single

'Recorsets Auxiliares
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Ampo As New ADODB.Recordset
Dim rs_WF_impproarg As New ADODB.Recordset
Dim rs_ImpMesArg As New ADODB.Recordset
Dim rs_AmpoConTpa As New ADODB.Recordset
Dim rs_item As New ADODB.Recordset
Dim rs_AcuLiq As New ADODB.Recordset
Dim rs_valitem As New ADODB.Recordset
Dim rs_desmen As New ADODB.Recordset
Dim rs_desliq As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_escala_ded As New ADODB.Recordset
Dim rs_escala As New ADODB.Recordset
Dim rs_ficharet As New ADODB.Recordset

Dim Hasta As Integer

' FGZ - 27/02/2004
Dim Terminar As Boolean
Dim pos1
Dim pos2
' FGZ - 27/02/2004

' FGZ - 12/02/2004
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Single
' FGZ - 12/02/2004


'Inicializacion de variables
p_Devuelve = 1001
p_Tope_Gral = 1002
p_Neto = 1003
p_Diciembre = 1005
p_Porcgross = 1010

p_grossant1 = 1011
p_grossant2 = 1012
p_acuimp = 1013
p_porcjub = 1014
p_porcley = 1015
p_porcosoc = 1016
p_porcsin = 1017
p_prorratea = 1005

AMPO = 4800
Porcapo = 0
Devuelve = 1
Tope_Gral = 100
Neto = 999999
Retencion = 0
Gan_Imponible = 0
Gan_Neta = 0
Deducciones = 0
Ded_a23 = 0
Por_Deduccion = 0
Impuesto_Escala = 0
Ret_Ant = 0
Porcgross = 100
Ajustecargas = 0
Ajustededucc = 0
Topeescala = 0
M2asgross = 0
M3asgross = 0
M4asgross = 0
M5asgross = 0
Masgross = 0
Porcant = 0
Diciembre = 0

grossant1 = 0
grossant2 = 0
AcuImp = 34
PorcJub = 0
PorcLey = 0
PorcOSoc = 0
PorcSin = 0

AuxOSPriv = 0
AjuOSPriv = 0

Acumulado_Mes = 0

Bien = False

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

Ret_mes = Month(buliq_proceso!profecpago)
Ret_ano = Year(buliq_proceso!profecpago)
Con_liquid = Buliq_Concepto(Concepto_Actual).concnro


'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case p_Devuelve:
        Devuelve = rs_wf_tpa!Valor
    Case p_Tope_Gral:
        Tope_Gral = rs_wf_tpa!Valor
    Case p_Neto:
        Neto = rs_wf_tpa!Valor
    Case p_Porcgross:
        Porcgross = rs_wf_tpa!Valor
    Case p_Diciembre:
        Diciembre = rs_wf_tpa!Valor
    Case p_grossant1:
        grossant1 = rs_wf_tpa!Valor
    Case p_grossant2:
        grossant2 = rs_wf_tpa!Valor
    Case p_acuimp:
        AcuImp = rs_wf_tpa!Valor
    Case p_porcjub:
        PorcJub = rs_wf_tpa!Valor
    Case p_porcley:
        PorcLey = rs_wf_tpa!Valor
    Case p_porcosoc:
        PorcOSoc = rs_wf_tpa!Valor
    Case p_porcsin:
        PorcSin = rs_wf_tpa!Valor
    Case p_prorratea:
        prorratea = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop


If Diciembre <> 1 Then
    Ret_mes = 12
End If

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Mes de Retencion " & Ret_mes
    Flog.writeline Espacios(Tabulador * 1) & "Año de Retencion " & Ret_ano
    
    Flog.writeline Espacios(Tabulador * 1) & "Máxima Ret. en %" & Tope_Gral
    Flog.writeline Espacios(Tabulador * 1) & "Neto del Mes" & Neto
End If


If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Tope_Gral, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Neto, "Neto del Mes", Neto)
End If

If Ret_mes > 6 Then
    MesSemestreDesde = 7
    MesSemestreHasta = 12
Else
    MesSemestreDesde = 7
    MesSemestreHasta = 12
End If


Porcapo = PorcJub + PorcLey + PorcOSoc + PorcSin


StrSql = "SELECT * FROM ampo WHERE ampofecha <=" & ConvFecha(buliq_proceso!profecpago)
OpenRecordset StrSql, rs_Ampo


If Not rs_Ampo.EOF Then
    rs_Ampo.MoveLast
    
    Valor_Ampo = rs_Ampo!Valor
    AMPO = 0
    Valor_Ampo_Cont = rs_Ampo!contvalor
End If


StrSql = "SELECT * FROM " & TTempWF_impproarg & " WHERE acunro =" & AcuImp
OpenRecordset StrSql, rs_WF_impproarg

Do While Not rs_WF_impproarg.EOF
    'Buscar el acumulado mensual Imponible que me interesa
    StrSql = "SELECT * FROM impmesarg " & _
             " WHERE acunro =" & rs_WF_impproarg!acunro & _
             " AND imaanio = " & buliq_periodo!pliqanio & _
             " AND imames = " & buliq_periodo!pliqmes & _
             " AND ternro = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_ImpMesArg
     
    
    For Contador = 1 To 3 'puede ir hasta 5
            
        Cant_Diaria_Ampos(Contador) = rs_Ampo!ampodiario
        Cant_Ampo_Proporcionar(Contador) = 0
        
        StrSql = "SELECT * FROM ampocontpa " & _
                 " INNER JOIN concepto ON concepto.concnro = ampocontpa " & _
                 " INNER JOIN detliq ON detliq.concnro = concepto.concnro " & _
                 " WHERE concepto.tconnro = " & Contador & _
                 " AND detliq.cliqnro = " & buliq_cabliq!cliqnro
        OpenRecordset StrSql, rs_AmpoConTpa
        
        Do While rs_AmpoConTpa.EOF
            Cant_Ampo_Proporcionar(Contador) = Cant_Ampo_Proporcionar(Contador) + IIf(CBool(rs_AmpoConTpa!signo), rs_AmpoConTpa!dlicant, rs_AmpoConTpa!dlicant * -1)
         
            rs_AmpoConTpa.MoveNext
        Loop
         
        'Buscar los imponible ya utilizados para ser restados:
        '   Sueldo(1) y Vacaciones(2): solo lo del mes
        '   SAC(3): lo del semestre que estoy liquidando
        If Contador = 1 Or Contador = 2 Then
            ' solo tomar lo del mes como ya acumulado
            If Not rs_ImpMesArg.EOF Then
                If Contador = 1 Then
                    Acumulado_Mes = rs_ImpMesArg!imamonto_1
                Else
                    Acumulado_Mes = rs_ImpMesArg!imamonto_2
                End If
            Else
                Acumulado_Mes = 0
            End If
        End If

        If Contador = 3 Then
            ' Tomar el semestre como ya acumulado
            Acumulado_Mes = 0

            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE acunro = " & rs_WF_impproarg!acunro & _
                     " AND imaanio = " & buliq_periodo!pliqanio & _
                     " AND imames >= " & MesSemestreDesde & _
                     " AND imames <= " & MesSemestreHasta & _
                     " AND ternro = " & buliq_empleado!ternro
            OpenRecordset StrSql, rs_ImpMesArg

            Do While Not rs_ImpMesArg.EOF
                Acumulado_Mes = Acumulado_Mes + rs_ImpMesArg!imamonto_3

                rs_ImpMesArg.MoveNext
            Loop
        End If
        ' calculo de ya acumulado de SAC


        ' Verificar los topes grales. contra mes mas el proceso actual
        ' si lo supera se guarda al resto hasta el tope, nunca puede ser negativo.

        ' ------------ CONTROL PARA LOS TOPES DEL AMPO --------------------
        ' Calcular el tope de AMPO proporcionado de acurdo a la cantidad de dias que estoy liquidando

        If CBool(rs_WF_impproarg!tope_aporte) Then
            Ampo_A_Aplicar = Valor_Ampo
        Else
            Ampo_A_Aplicar = Valor_Ampo_Cont
        End If

        If Cant_Ampo_Proporcionar(1) = 0 Or Not CBool(rs_Ampo!ampopropo) Then
            If rs_WF_impproarg!tope_aporte Then
                  Select Case Contador
                  Case 1:
                      Tope_Monto_Proporcional = buliq_impgralarg!ipgtopemonto_1
                  Case 2:
                      Tope_Monto_Proporcional = buliq_impgralarg!ipgtopemonto_2
                  Case 3:
                      Tope_Monto_Proporcional = buliq_impgralarg!ipgtopemonto_3
                  End Select
            Else
              Tope_Monto_Proporcional = Valor_Ampo_Cont * rs_Ampo!ampomax
            End If
        Else
            Select Case Contador
            Case 1:
                Tope_Monto_Proporcional = (Cant_Ampo_Proporcionar(1) * Cant_Diaria_Ampos(1)) * Ampo_A_Aplicar
            Case 2:
                Tope_Monto_Proporcional = (Cant_Ampo_Proporcionar(2) * Cant_Diaria_Ampos(2)) * Ampo_A_Aplicar
            Case 3:
                Tope_Monto_Proporcional = (Cant_Ampo_Proporcionar(3) * Cant_Diaria_Ampos(3)) * Ampo_A_Aplicar
            End Select
        End If

        Select Case Contador
        Case 1:
            Monto_wf(Contador) = rs_WF_impproarg!ipamonto_1
        Case 2:
            Monto_wf(Contador) = rs_WF_impproarg!ipamonto_2
        Case 3:
            Monto_wf(Contador) = rs_WF_impproarg!ipamonto_3
        End Select
        
        If (Acumulado_Mes + Monto_wf(Contador)) > Tope_Monto_Proporcional Then
            ' Actualizo
            Monto_wf(Contador) = Tope_Monto_Proporcional - Acumulado_Mes
        End If
        
        If Monto_wf(Contador) > 0 Then
            AMPO = AMPO + (Tope_Monto_Proporcional - Acumulado_Mes)
        End If

    Next Contador
    
    StrSql = "SELECT * FROM acu_liq WHERE acunro =" & rs_WF_impproarg!acunro & _
             " AND cliqnro = " & buliq_cabliq!cliqnro
    OpenRecordset StrSql, rs_AcuLiq ' Debe existir siempre

    If Monto_wf(1) > grossant1 Then
        Monto_wf(1) = Monto_wf(1) - grossant1
    End If
    If Monto_wf(2) > grossant2 Then
        Monto_wf(2) = Monto_wf(2) - grossant2
    End If

    If Not rs_AcuLiq.EOF Then
        ' Pisar en el Acu_liq
        Monto_Tope = Monto_wf(1) + Monto_wf(2) + Monto_wf(3) + Monto_wf(4) + Monto_wf(5)
    End If
    
    PorcJub = IIf(PorcJub <> 0, (Monto_Tope * PorcJub / 100), 0)
    PorcLey = IIf(PorcLey <> 0, (Monto_Tope * PorcLey / 100), 0)
    PorcOSoc = IIf(PorcOSoc <> 0, (Monto_Tope * PorcOSoc / 100), 0)
    PorcSin = IIf(PorcSin <> 0, (Monto_Tope * PorcSin / 100), 0)

    ' Armo la traza del item
    If HACE_TRAZA Then
        Texto = "0 - Grossing 1"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, grossant1)
        Texto = "0 - Grossing 2"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, grossant2)
        Texto = "0 - Acu Base Cargas"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, AcuImp)
        Texto = "0 - AMPO Gral"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, AMPO)
        Texto = "0 - AMPO tipo 1"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Monto_wf(1))
        Texto = "0 - AMPO tipo 2"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Monto_wf(2))
        Texto = "0 - AMPO tipo 3"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Monto_wf(3))
        Texto = "0 - Jubilacion"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, PorcJub)
        Texto = "0 - Ley"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, PorcLey)
        Texto = "0 - O. Social"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, PorcOSoc)
        Texto = "0 - Sindicato"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, PorcSin)
    End If

    rs_WF_impproarg.MoveNext
Loop


' -------------------------------------------------------------------------
' Fin calculo tope y aportes del empleado para tomar en el grosssing
' -------------------------------------------------------------------------

' Recorro todos los items de Ganancias
StrSql = "SELECT * FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_item

Do While Not rs_item.EOF
    Select Case rs_item!itetipotope
    Case 1: ' el valor a tomar es lo que dice la escala
    
        StrSql = "SELECT * FROM valitem WHERE viano =" & Ret_ano & _
                 " AND itenro=" & rs_item!itenro & _
                 " AND vimes =" & Ret_mes
        OpenRecordset StrSql, rs_valitem
        
        Do While Not rs_valitem.EOF
            Items_DDJJ(rs_valitem!itenro) = rs_valitem!vimonto
            Items_TOPE(rs_valitem!itenro) = rs_valitem!vimonto
            
            rs_valitem.MoveNext
        Loop
    ' End case 1
    ' ------------------------------------------------------------------------

    Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
        ' Busco la declaracion jurada
        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!ternro & _
                 " AND desano=" & Ret_ano & _
                 " AND itenro = " & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
        
        Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_item!itenro = 3 Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                Else
                    If rs_desmen!desmenprorra = 0 Then 'no es parejito
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec
                    Else
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                    End If
                End If
                    
'                Else
'                    If rs_desmen!desmonprorra = 0 Then 'no es parejito
'                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
'                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec
'                    Else
'                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
'                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
'                    End If
'                End If
            End If
            
            rs_desmen.MoveNext
        Loop
        
        Aux_Inicio_Ret = CDate("01/01/ " & Ret_ano)
        Aux_Fin_Ret = DateAdd("d", -1, CDate("01/" & (Ret_mes + 1) & "/" & Ret_ano))
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha <= " & ConvFecha(Aux_Fin_Ret) & _
                 " AND dlfecha >= " & ConvFecha(Aux_Inicio_Ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto
            'Si el desliq prorratea debo proporcionarlo
            'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 Or CBool(rs_desliq!dlprorratea)), rs_desliq!dlmonto / (13 - Month(rs_desliq!dlfecha)) * (Ret_mes - Month(rs_desliq!dlfecha) + 1), rs_desliq!dlmonto)
            Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 Or CBool(rs_desliq!dlprorratea)) And (prorratea = 1), rs_desliq!dlmonto / (13 - Month(rs_desliq!dlfecha)) * (Ret_mes - Month(rs_desliq!dlfecha) + 1), rs_desliq!dlmonto)
            
            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itaprorratea) Then
                    If CBool(rs_itemacum!itasigno) Then
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    End If
                Else
                    If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    End If
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) Then
                If CBool(rs_itemconc!itcsigno) Then
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                End If
            Else
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                End If
            End If
        
            rs_itemconc.MoveNext
        Loop
    ' End case 2
    ' ------------------------------------------------------------------------
        
     Case 3: 'TOMO LOS VALORES DE LA DDJJ Y LIQUIDACION Y EL TOPE PARA APLICARLO

        StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                 " AND vimes = " & Ret_mes & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_valitem
         Do While Not rs_valitem.EOF
            Items_TOPE(rs_item!itenro) = rs_valitem!vimonto

            rs_valitem.MoveNext
         Loop

        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
         
            rs_desmen.MoveNext
         Loop

        Aux_Inicio_Ret = CDate("01/01/ " & Ret_ano)
        Aux_Fin_Ret = DateAdd("d", -1, CDate("01/" & (Ret_mes + 1) & "/" & Ret_ano))

        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(Aux_Inicio_Ret) & _
                 " AND dlfecha <= " & ConvFecha(Aux_Fin_Ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop

        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------

        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
        
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
        End If

        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_item!itesigno) Then
            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
        End If

    ' End case 3
    ' ------------------------------------------------------------------------

    Case 4:
        ' Tomo los valores de la DDJJ y el valor de la escala (cargas de familia)

        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfechas) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Month(rs_desmen!desfechas) - Month(rs_desmen!desfecdes) + 1)
            Else
                If Month(rs_desmen!desfecdes) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
                End If
            End If

            rs_desmen.MoveNext
         Loop

        If Items_DDJJ(rs_item!itenro) > 0 Then
            StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                     " AND vimes = " & Ret_mes & _
                     " AND itenro =" & rs_item!itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_item!itenro) = rs_valitem!vimonto / Ret_mes * Items_DDJJ(rs_item!itenro)

                rs_valitem.MoveNext
             Loop
        End If
    ' End case 4
    ' ------------------------------------------------------------------------

    Case 5:
        i = 1
        Hasta = IIf(50 > Len(rs_item!iteitemstope), 50, rs_item!iteitemstope)
        Terminar = False
        Do While i <= Hasta And Not Terminar
            pos1 = i
            pos2 = InStr(i, rs_item!iteitemstope, ",") - 1
            If pos2 > 0 Then
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
            Else
                pos2 = Len(rs_item!iteitemstope)
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
                Terminar = True
            End If
            
            If Texto <> "" Then
                If Mid(Texto, 1, 1) = "-" Then
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                Else
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                End If
            End If
            i = pos2 + 2
        Loop
        
        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) * rs_item!iteporctope / 100


        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
            ' Tocado por Maxi 26/05/2004 faltaba el parejito
            'If Month(rs_desmen!desfecdes) <= Ret_mes Then
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
            'Else
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
            'End If
         
            rs_desmen.MoveNext
         Loop


        Aux_Inicio_Ret = CDate("01/01/ " & Ret_ano)
        Aux_Fin_Ret = DateAdd("d", -1, CDate("01/" & (Ret_mes + 1) & "/" & Ret_ano))

        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(Aux_Inicio_Ret) & _
                 " AND dlfecha <= " & ConvFecha(Aux_Fin_Ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop


        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------

'        'Busco los acumuladores de la liquidacion
'        StrSql = "SELECT * FROM itemacum " & _
'                 " INNER JOIN acu_liq ON itemacum.acunro = acu_liq.acunro " & _
'                 " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND itenro =" & rs_item!itenro & _
'                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemacum
'
'        Do While Not rs_itemacum.EOF
'                If CBool(rs_itemacum!itasigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
'                End If
'
'            rs_itemacum.MoveNext
'        Loop

        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
' FGZ - 22/06/2004
'        'TOPEO LOS VALORES
'        If Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro) < Items_TOPE(rs_item!itenro) Then
'            Items_TOPE(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro)
'        End If

' FGZ - 22/06/2004
' puse lo mismo que para el itemtope 3
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_item!itesigno) Then
            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
        End If

    ' End case 5
    ' ------------------------------------------------------------------------
    Case Else:
    End Select

    If rs_item!itenro > 7 Then
        Items_TOPE(rs_item!itenro) = IIf(CBool(rs_item!itesigno), Items_TOPE(rs_item!itenro), Abs(Items_TOPE(rs_item!itenro)))
    End If
    
    'Armo la traza del item
    If CBool(USA_DEBUG) Then
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-DDJJ"
        Flog.writeline Espacios(Tabulador * 1) & Items_DDJJ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Liq"
        Flog.writeline Espacios(Tabulador * 1) & Items_LIQ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-LiqAnt"
        Flog.writeline Espacios(Tabulador * 1) & Items_OLD_LIQ(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Prorr"
        Flog.writeline Espacios(Tabulador * 1) & Items_PRORR(rs_item!itenro)
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Tope"
        Flog.writeline Espacios(Tabulador * 1) & Items_TOPE(rs_item!itenro)
    End If
    If HACE_TRAZA Then
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_DDJJ(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_LIQ(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_OLD_LIQ(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_PRORR(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_TOPE(rs_item!itenro))
    End If
        
    
    'Calcula la Ganancia Imponible
    If CBool(rs_item!itesigno) Then
        Gan_Imponible = Gan_Imponible + Items_TOPE(rs_item!itenro)
    Else
        If (rs_item!itetipotope = 1) Or (rs_item!itetipotope = 4) Then
            Ded_a23 = Ded_a23 - Items_TOPE(rs_item!itenro)
        Else
            Deducciones = Deducciones - Items_TOPE(rs_item!itenro)
        End If
    End If
            
    rs_item.MoveNext
Loop
            
           
'Sumo las simulaciones de jub, ley, O soc. y sindicato
Gan_Imponible = Gan_Imponible - PorcJub - PorcOSoc - PorcLey - PorcSin

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "9- Ganancia Neta " & Gan_Imponible
        Flog.writeline Espacios(Tabulador * 1) & "9- Total Deducciones" & Deducciones
        Flog.writeline Espacios(Tabulador * 1) & "9- Total art. 23" & Ded_a23
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Neta ", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Total art. 23", Ded_a23)
    End If


Gan_Neta = Gan_Imponible

StrSql = "SELECT * FROM escala_ded " & _
         " WHERE esd_topeinf <= " & (Gan_Imponible / Ret_mes * 12) & _
         " AND esd_topesup >=" & (Gan_Imponible / Ret_mes * 12)
OpenRecordset StrSql, rs_escala_ded

If Not rs_escala_ded.EOF Then
    Por_Deduccion = rs_escala_ded!esd_porcentaje
Else
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "No hay esc. dedu para", Gan_Imponible)
    End If
    ' No se ha encontrado la escala de deduccion para el valor gan_imponible
End If

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- % a tomar deduc.", Por_Deduccion)
End If

'Aplico el porcentaje a las deducciones
If Ret_ano >= 2000 Then
    Ded_a23 = Ded_a23 * Por_Deduccion / 100
End If

'Calculo la ganancia imponible
Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Imponible", Gan_Imponible)
End If
        
            
If Gan_Imponible > 0 Then
    'Entrar en la escala con las ganancias acumuladas
    StrSql = "SELECT * FROM escala " & _
             " WHERE escmes =" & Ret_mes & _
             " AND escano =" & Ret_ano & _
             " AND escinf <= " & Gan_Imponible & _
             " AND escsup >= " & Gan_Imponible
    OpenRecordset StrSql, rs_escala
    
    If Not rs_escala.EOF Then
        Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
    Else
        Impuesto_Escala = 0
    End If
Else
    Impuesto_Escala = 0
End If
        
        
' Calculo las retenciones ya realizadas
Ret_ano = 0
'For each ficharet where ficharet.empleado = buliq-empleado.ternro
'                    And Month(ficharet.fecha) <= ret-mes
'                    And Year(ficharet.fecha) = ret-ano NO-LOCK:
'    Assign Ret-ant = Ret-Ant + ficharet.importe.
'End.
'como no puede utilizar la funcion month() en sql
'levanto todas las ficharet del tercero y hago la pregunta dentro del loop
StrSql = "SELECT * FROM ficharet " & _
         " WHERE empleado =" & buliq_empleado!ternro
OpenRecordset StrSql, rs_ficharet

Do While Not rs_ficharet.EOF
    If (Month(rs_ficharet!Fecha) = Ret_mes) And (Year(rs_ficharet!Fecha) = Ret_ano) Then
        Ret_Ant = Ret_Ant + rs_ficharet!Importe
    End If
    rs_ficharet.MoveNext
Loop

'Calcular la retencion
Retencion = Impuesto_Escala - Ret_Ant

' Verifico si es una devolucion y si devuelve el concepto
If Retencion < 0 Then
    Retencion = 0
End If

' Verificar que la rtencion no supere el 30% del Neto del Mes
If Retencion > (Neto * (Tope_Gral / 100)) Then
    Retencion = Neto * (Tope_Gral / 100)
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver, x Tope General", Retencion)
    End If
End If

If Gan_Imponible > 0 Then
    'Calcular el grossing
    If rs_escala.EOF Then
        Topeescala = rs_escala!escsup
        Porcant = rs_escala!escporexe
    Else
        Topeescala = 0
        Porcant = 0
    End If

    If Retencion > 0 Then
        Retencion = Retencion / (1 - (rs_escala!escporexe))
    Else
        Retencion = Retencion * (1 - (rs_escala!escporexe))
    End If
    
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - grossing 1", Retencion)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - porcent 1", rs_escala!escporexe)
    End If
    
    ' Me fijo si se pasa de escala
    If Retencion > 0 And (Gan_Imponible + Retencion) > rs_escala!escsup Then
        StrSql = "SELECT * FROM escala " & _
                 " WHERE escmes =" & Ret_mes & _
                 " AND escano =" & Ret_ano & _
                 " AND escinf <= " & (Gan_Imponible + Retencion) & _
                 " AND escsup >= " & (Gan_Imponible + Retencion)
        OpenRecordset StrSql, rs_escala
        If Not rs_escala.EOF Then
            Masgross = Retencion + Gan_Imponible - rs_escala!escinf
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Excede Escala", Masgross)
        End If
        
        Masgross = Masgross * (1 / (1 - ((rs_escala!escporexe - Porcant) / 100)) - 1)
        M2asgross = Masgross * (1 / (1 - ((Porcant) / 100)) - 1)
        M3asgross = M2asgross * (1 / (1 - ((rs_escala!escporexe - Porcant) / 100)) - 1)
        M4asgross = M3asgross * (1 / (1 - ((Porcant) / 100)) - 1)
        M5asgross = M4asgross * (1 / (1 - ((rs_escala!escporexe - Porcant) / 100)) - 1)
        
        Retencion = Retencion + Masgross + M2asgross + M3asgross + M4asgross + M5asgross
        
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Grossing 2", Masgross + M2asgross + M3asgross + M4asgross + M5asgross)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Porcent 2", rs_escala!escporexe)
        End If
    End If ' If Retencion > 0 And (gan_imponible + Retencion) > rs_escala!escsup Then
        
    'Esto es para compensar los cambios en los porcentajes de deducciones
    If Retencion > 0 Then
        StrSql = "SELECT * FROM escala_ded " & _
                 " WHERE esd_topeinf <= " & ((Gan_Neta + Retencion) / Ret_mes * 12) & _
                 " AND esd_topesup >=" & ((Gan_Neta + Retencion) / Ret_mes * 12)
        OpenRecordset StrSql, rs_escala_ded
                
        If Not rs_escala_ded.EOF Then
            If Por_Deduccion <> rs_escala_ded!esd_porcentaje Then
                Ded_a23 = Ded_a23 / Por_Deduccion * 100
                baja_ded = Ded_a23 * (Por_Deduccion - rs_escala_ded!esd_porcentaje) / 100
                ' lo que queda por encima del tope
                baja2 = IIf((Gan_Imponible + Retencion + baja_ded > Topeescala) And (Gan_Imponible + Retencion < Topeescala), Gan_Imponible + Retencion + baja_ded - Topeescala, 0)
                baja1 = (baja_ded - baja2)
                baja2 = baja2 * (1 / (1 - (Porcant / 100)) - 1)
                baja1 = baja1 * (1 / (1 - (rs_escala!escporexe / 100)) - 1)
                Ajustededucc = baja1 + baja2
            End If
        Else
            ' No se ha encontrado la escala de deduccion para el valor (gan_neta + retencion)
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "No Hay esc. Dedu para ", Gan_Neta)
            End If
        End If
    
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9 - dif. por camb. deduc.", baja_ded)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9 - Ajuste por ded. 1", baja1)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9 - Ajuste por ded. 2", baja2)
        End If
    
        ' Ajuste por la variacion en el tope Obra Social Privada
        AuxOSPriv = Items_LIQ(13) + Items_OLD_LIQ(13) + Items_DDJJ(13) + Items_TOPE(13)
        If AuxOSPriv > (Ajustededucc + Retencion) * 0.05 Then
            AjuOSPriv = (Ajustededucc + Retencion) * 0.05
            AjuOSPriv = AjuOSPriv * (rs_escala!escporexe)
            AjuOSPriv = AjuOSPriv + (AjuOSPriv * rs_escala!escporexe / 100) + (AjuOSPriv * rs_escala!escporexe / 100) * (rs_escala!escporexe / 100)
        End If
        
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "91 - AjusteDeducc.", Ajustededucc)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "91 - escporexe", rs_escala!escporexe)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "91 - Dif. por OSP", AuxOSPriv)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "91 - Ajuste por OSP", AjuOSPriv)
        End If
    End If 'If Retencion > 0 Then
End If 'If gan_imponible > 0 Then

' Ajustes poe cambios en cargas sociales
If Items_LIQ(1) + Items_LIQ(3) + Items_LIQ(4) > AMPO Then
    Ajustecargas = 0
Else
    If Items_LIQ(1) + Items_LIQ(3) + Items_LIQ(4) + Ajustededucc + Retencion > AMPO Then
        Ajustecargas = (Ajustededucc + Retencion) * (1 / (1 - (Porcapo / 100)))
    Else
        If Retencion > 0 Then
            Ajustecargas = (Ajustededucc + Retencion) * (1 / (1 - (Porcapo / 100)) - 1)
        Else
            Ajustecargas = -(Ajustededucc + Retencion) * (1 / (1 - (Porcapo / 100)) - 1)
        End If
    End If
End If

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Ajuste Cargas", Ajustecargas)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Ajuste Deducc.", Ajustededucc)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Ajuste por OSP", AjuOSPriv)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Retenciones anteriores", Ret_Ant)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Impuesto por Escala", Impuesto_Escala)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver", Retencion)
End If


If Retencion > 0 Then
    Monto = (Retencion + Ajustecargas) * Porcgross / 100 - Ajustededucc - AjuOSPriv - grossant1 - grossant2
Else
    Monto = (Retencion - Ajustecargas)
End If
Bien = True
    

End Function


'Public Function for_Grossing5_old() As Single
'' _____________________________________________________________________________________________
'' Descripcion           : Programa para el clculo de la Retenci¢n de Ganancias (Grossing)
'' Autor                 :
'' Fecha                 : Version Nueva 26_11_2002
'' Traducido por         : FGZ
'' Parametros de entrada :
'' Ultima Mod.           :
'' Descripcion           :
'' _____________________________________________________________________________________________
'
'' Tipos de parametros usados
'
'Dim p_Ampo        As Integer ' 1000   /* Valor del tope ampo */
'Dim p_Devuelve    As Integer ' 1001   /* Si devuelve ganancias o no */
'Dim p_Tope_Gral   As Integer ' 1002   /* Tope Gral de Retenci¢n */
'Dim p_Neto        As Integer ' 1003   /* Base para el tope */
'Dim p_Porcapo     As Integer ' 1004   /* Porcentaje de aporte de cargas sociales */
'Dim p_Diciembre   As Integer ' 1005   /* Toma la escala de diciembre */
'Dim p_Porcgross   As Integer ' 1010   /* Porcentaje de grossing a efectuar */
'Dim p_Imgrossing  As Integer ' 51     /* Sueldo Neto pactado */
''
''/*  Variables Locales */
'Dim AMPO             As Single ' 4800
'Dim Porcapo          As Single ' 0
'Dim Devuelve         As Single ' 1
'Dim Tope_Gral        As Single ' 100
''
'Dim Imp_Neto         As Single ' 0
'Dim Neto             As Single ' 999999
'Dim Retencion        As Single ' 0
'Dim Gan_Imponible    As Single ' 0
'Dim Gan_Neta         As Single ' 0
'Dim Deducciones      As Single ' 0
'Dim Ded_a23          As Single ' 0
'Dim Por_Deduccion    As Single ' 0
'Dim Impuesto_Escala As Single  ' 0
'Dim Ret_Ant          As Single ' 0
'Dim Porc_Gross       As Single ' 100
'Dim Ajustecargas     As Single ' 0
'Dim Difcargas     As Single ' 0
'Dim Ajustededucc     As Single ' 0
'Dim Topeescala       As Single ' 0
'Dim Masgross         As Single ' 0
'Dim M2asgross        As Single ' 0
'Dim M3asgross        As Single ' 0
'Dim M4asgross        As Single ' 0
'Dim M5asgross        As Single ' 0
'Dim Porcant          As Single ' 0
'Dim Diciembre        As Single ' 0
'Dim Imgrossing       As Single ' 0
'Dim Grossingup       As Single ' 0
'
'Dim Ret_mes          As Integer
'Dim Ret_ano          As Integer
'Dim Aux_Fin_Ret As Date
'Dim Aux_Inicio_Ret As Date
'Dim Con_liquid       As Integer
'Dim i                As Integer
'Dim Texto            As String
'
''/* vectores para mejorar el proceso */
'Dim Items_DDJJ(50)       As Single
'Dim Items_LIQ(50)        As Single
'Dim Items_OLD_LIQ(50)    As Single
'Dim Items_TOPE(50)       As Single
'
'Dim Impuesto_Escala2 As Single  ' 0
'Dim Diferencia As Single
'Dim auxi As Single
'Dim Auxi2 As Single
'
'Dim baja_ded As Single
'Dim baja1 As Single
'Dim baja2 As Single
'
''Recorsets Auxiliares
'Dim rs_wf_tpa As New ADODB.Recordset
'Dim rs_Ampo As New ADODB.Recordset
'Dim rs_WF_impproarg As New ADODB.Recordset
'Dim rs_ImpMesArg As New ADODB.Recordset
'Dim rs_AmpoConTpa As New ADODB.Recordset
'Dim rs_item As New ADODB.Recordset
'Dim rs_AcuLiq As New ADODB.Recordset
'Dim rs_valitem As New ADODB.Recordset
'Dim rs_desmen As New ADODB.Recordset
'Dim rs_desliq As New ADODB.Recordset
'Dim rs_itemacum As New ADODB.Recordset
'Dim rs_itemconc As New ADODB.Recordset
'Dim rs_escala_ded As New ADODB.Recordset
'Dim rs_escala As New ADODB.Recordset
'Dim rs_ficharet As New ADODB.Recordset
'Dim Hasta As Integer
'' FGZ - 27/02/2004
'Dim Terminar As Boolean
'Dim pos1
'Dim pos2
'' FGZ - 27/02/2004
'
'' FGZ - 12/02/2004
'Dim rs_acumulador As New ADODB.Recordset
'Dim Acum As Long
'Dim Aux_Acu_Monto As Single
'' FGZ - 12/02/2004
'
''Inicializaciones
'p_Ampo = 1000
'p_Devuelve = 1001
'p_Tope_Gral = 1002
'p_Neto = 1003
'p_Porcapo = 1004
'p_Diciembre = 1005
'p_Porcgross = 1010
'p_Imgrossing = 51
'
'AMPO = 4800
'Porcapo = 0
'Devuelve = 1
'Tope_Gral = 100
'
'Imp_Neto = 0
'Neto = 999999
'Retencion = 0
'Gan_Imponible = 0
'Gan_Neta = 0
'Deducciones = 0
'Ded_a23 = 0
'Por_Deduccion = 0
'Impuesto_Escala = 0
'Ret_Ant = 0
'Porc_Gross = 100
'Ajustecargas = 0
'Difcargas = 0
'Ajustededucc = 0
'Topeescala = 0
'Masgross = 0
'M2asgross = 0
'M3asgross = 0
'M4asgross = 0
'M5asgross = 0
'Porcant = 0
'Diciembre = 0
'Imgrossing = 0
'Grossingup = 0
'
'Impuesto_Escala2 = 0
'Diferencia = 0
'
'' Comienzo
'Bien = False
'
'StrSql = "SELECT * FROM acu_liq WHERE acunro  = 6 "
'OpenRecordset StrSql, rs_AcuLiq
'If Not rs_AcuLiq.EOF Then
'    Imp_Neto = rs_AcuLiq!almonto
'Else
'    Imp_Neto = 0
'End If
'
'If HACE_TRAZA Then
'    Call LimpiarTraza(Buliq_concepto(concepto_actual).concnro)
'End If
'
'Ret_mes = Month(buliq_proceso!profecpago)
'Ret_ano = Year(buliq_proceso!profecpago)
'Con_liquid = Buliq_concepto(concepto_actual).concnro
'
'
''Obtencion de los parametros de WorkFile
'StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha =" & ConvFecha(AFecha)
'OpenRecordset StrSql, rs_wf_tpa
'
'Do While Not rs_wf_tpa.EOF
'    Select Case rs_wf_tpa!tipoparam
'    Case p_Ampo:
'        AMPO = rs_wf_tpa!Valor
'    Case p_Devuelve:
'        Devuelve = rs_wf_tpa!Valor
'    Case p_Tope_Gral:
'        Tope_Gral = rs_wf_tpa!Valor
'    Case p_Neto:
'        Neto = rs_wf_tpa!Valor
'    Case p_Porcgross:
'        Porc_Gross = rs_wf_tpa!Valor
'    Case p_Diciembre:
'        Diciembre = rs_wf_tpa!Valor
'    Case p_Imgrossing:
'        Imgrossing = rs_wf_tpa!Valor
'    Case p_Porcapo:
'        Porcapo = rs_wf_tpa!Valor
'    End Select
'
'    rs_wf_tpa.MoveNext
'Loop
'
'If Diciembre <> 1 Then
'    Ret_mes = 12
'End If
'
'If HACE_TRAZA Then
'    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, p_Tope_Gral, "Máxima Ret. en %", Tope_Gral)
'    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, p_Neto, "Neto del Mes", Neto)
'End If
'
'
'' Recorro todos los items de Ganancias
'StrSql = "SELECT * FROM item ORDER BY itetipotope"
'OpenRecordset StrSql, rs_item
'
'Do While Not rs_item.EOF
'    Select Case rs_item!itetipotope
'    Case 1: ' el valor a tomar es lo que dice la escala
'
'        StrSql = "SELECT * FROM valitem WHERE viano =" & Ret_ano & _
'                 " AND itenro=" & rs_item!itenro & _
'                 " AND vimes =" & Ret_mes
'        OpenRecordset StrSql, rs_valitem
'
'        Do While Not rs_valitem.EOF
'            Items_DDJJ(rs_valitem!itenro) = rs_valitem!vimonto
'            Items_TOPE(rs_valitem!itenro) = rs_valitem!vimonto
'
'            rs_valitem.MoveNext
'        Loop
'    ' End case 1
'    ' ------------------------------------------------------------------------
'
'    Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
'        ' Busco la declaracion jurada
'        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!Ternro & _
'                 " AND desano=" & Ret_ano & _
'                 " AND itenro = " & rs_item!itenro
'        OpenRecordset StrSql, rs_desmen
'
'        Do While Not rs_desmen.EOF
'            If Month(rs_desmen!desfecdes) <= Ret_mes Then
'                If rs_item!itenro = 3 Then
'                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
'                Else
'                    If rs_desmen!desmonreal = 0 Then 'no es parejito
'                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
'                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec
'                    Else
'                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
'                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
'                    End If
'                End If
'            End If
'
'            rs_desmen.MoveNext
'        Loop
'
'        Aux_Inicio_Ret = CDate("01/01/ " & Ret_ano)
'        Aux_Fin_Ret = DateAdd("d", -1, CDate("01/" & (Ret_mes + 1) & "/" & Ret_ano))
'
'        'Busco las liquidaciones anteriores
'        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
'                 " AND empleado = " & buliq_empleado!Ternro & _
'                 " AND dlfecha <= " & ConvFecha(Aux_Fin_Ret) & _
'                 " AND dlfecha >= " & ConvFecha(Aux_Inicio_Ret)
'        OpenRecordset StrSql, rs_desliq
'
'        Do While Not rs_desliq.EOF
'            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto
'            'Si el desliq prorratea debo proporcionarlo
'            Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 Or CBool(rs_desliq!dlprorratea)), rs_desliq!dlmonto / (13 - Month(rs_desliq!dlfecha)) * (Ret_mes - Month(rs_desliq!dlfecha) + 1), rs_desliq!dlmonto)
'
'            rs_desliq.MoveNext
'        Loop
'
'        'Busco los acumuladores de la liquidacion
'        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
'        StrSql = "SELECT * FROM itemacum " & _
'                 " WHERE itenro =" & rs_item!itenro & _
'                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemacum
'        Do While Not rs_itemacum.EOF
'            Acum = CStr(rs_itemacum!acunro)
'            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
'                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
'
'                If CBool(rs_itemacum!itaprorratea) Then
'                    If CBool(rs_itemacum!itasigno) Then
'                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
'                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Aux_Acu_Monto / (13 - Ret_mes)
'                    Else
'                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
'                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Aux_Acu_Monto / (13 - Ret_mes)
'                    End If
'                Else
'                    If CBool(rs_itemacum!itasigno) Then
'                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
'                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3), rs_itemacum!almonto / (13 - Ret_mes), Aux_Acu_Monto)
'                    Else
'                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
'                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3), rs_itemacum!almonto / (13 - Ret_mes), Aux_Acu_Monto)
'                    End If
'                End If
'            End If
'            rs_itemacum.MoveNext
'        Loop
'        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
'
''        'Busco los acumuladores de la liquidacion
''        StrSql = "SELECT * FROM itemacum " & _
''                 " INNER JOIN acu_liq ON itemacum.acunro = acu_liq.acunro " & _
''                 " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
''                 " AND itenro =" & rs_item!itenro & _
''                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
''        OpenRecordset StrSql, rs_itemacum
''
''        Do While Not rs_itemacum.EOF
''            If CBool(rs_itemacum!itaprorratea) Then
''                If CBool(rs_itemacum!itasigno) Then
''                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
''                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_itemacum!almonto / (13 - Ret_mes)
''                Else
''                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
''                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - rs_itemacum!almonto / (13 - Ret_mes)
''                End If
''            Else
''                If CBool(rs_itemacum!itasigno) Then
''                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
''                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3), rs_itemacum!almonto / (13 - Ret_mes), rs_itemacum!almonto)
''                Else
''                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
''                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3), rs_itemacum!almonto / (13 - Ret_mes), rs_itemacum!almonto)
''                End If
''            End If
''
''            rs_itemacum.MoveNext
''        Loop
'
'
'        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
'        ' Busco los conceptos de la liquidacion
'        StrSql = "SELECT * FROM itemconc " & _
'                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
'                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND itemconc.itenro =" & rs_item!itenro & _
'                 " AND (itemconc.itaconcnrodest is null OR itemconc.itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemconc
'
'        Do While Not rs_itemconc
'            If CBool(rs_itemconc!itcprorratea) Then
'                If CBool(rs_itemconc!itcsigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_itemconc!dlimonto / (13 - Ret_mes)
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - rs_itemconc!dlimonto / (13 - Ret_mes)
'                End If
'            Else
'                If CBool(rs_itemacum!itcsigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
'                End If
'            End If
'
'            rs_itemconc.MoveNext
'        Loop
'    ' End case 2
'    ' ------------------------------------------------------------------------
'
'     Case 3: 'TOMO LOS VALORES DE LA DDJJ Y LIQUIDACION Y EL TOPE PARA APLICARLO
'
'        StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
'                 " AND vimes = " & Ret_mes & _
'                 " AND itenro =" & rs_item!itenro
'        OpenRecordset StrSql, rs_valitem
'         Do While Not rs_valitem.EOF
'            Items_TOPE(rs_item!itenro) = rs_valitem!vimonto
'
'            rs_valitem.MoveNext
'         Loop
'
'        'Busco la declaracion Jurada
'        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!Ternro & _
'                 " AND desano = " & Ret_ano & _
'                 " AND itenro =" & rs_item!itenro
'        OpenRecordset StrSql, rs_desmen
'         Do While Not rs_desmen.EOF
'            If Month(rs_desmen!desfecdes) <= Ret_mes Then
'                If rs_desmen!desmonreal = 0 Then ' No es parejito
'                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
'                Else
'                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
'                End If
'            End If
'
'            rs_desmen.MoveNext
'         Loop
'
'        Aux_Inicio_Ret = CDate("01/01/ " & Ret_ano)
'        Aux_Fin_Ret = DateAdd("d", -1, CDate("01/" & (Ret_mes + 1) & "/" & Ret_ano))
'
'        'Busco las liquidaciones anteriores
'        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
'                 " AND empleado = " & buliq_empleado!Ternro & _
'                 " AND dlfecha >= " & ConvFecha(Aux_Inicio_Ret) & _
'                 " AND dlfecha <= " & ConvFecha(Aux_Fin_Ret)
'        OpenRecordset StrSql, rs_desliq
'
'        Do While Not rs_desliq.EOF
'            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto
'
'            rs_desliq.MoveNext
'        Loop
'
'        'Busco los acumuladores de la liquidacion
'        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
'        StrSql = "SELECT * FROM itemacum " & _
'                 " WHERE itenro =" & rs_item!itenro & _
'                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemacum
'        Do While Not rs_itemacum.EOF
'            Acum = CStr(rs_itemacum!acunro)
'            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
'                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
'
'                If CBool(rs_itemacum!itasigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
'                End If
'            End If
'            rs_itemacum.MoveNext
'        Loop
'        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
'
''        'Busco los acumuladores de la liquidacion
''        StrSql = "SELECT * FROM itemacum " & _
''                 " INNER JOIN acu_liq ON itemacum.acunro = acu_liq.acunro " & _
''                 " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
''                 " AND itenro =" & rs_item!itenro & _
''                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
''        OpenRecordset StrSql, rs_itemacum
''
''        Do While Not rs_itemacum.EOF
''                If CBool(rs_itemacum!itasigno) Then
''                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
''                Else
''                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
''                End If
''
''            rs_itemacum.MoveNext
''        Loop
'
'        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
'        ' Busco los conceptos de la liquidacion
'        StrSql = "SELECT * FROM itemconc " & _
'                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
'                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND itemconc.itenro =" & rs_item!itenro & _
'                 " AND (itemconc.itaconcnrodest is null OR itemconc.itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemconc
'
'        Do While Not rs_itemconc
'                If CBool(rs_itemconc!itcsigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
'                End If
'
'            rs_itemconc.MoveNext
'        Loop
'
'
'        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
'        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
'            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
'        End If
'
'        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
'        ' "ACHIQUE" DE GANANCIA IMPONIBLE
'        If CBool(rs_item!itesigno) Then
'            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
'        End If
'
'    ' End case 3
'    ' ------------------------------------------------------------------------
'
'    Case 4:
'        ' Tomo los valores de la DDJJ y el valor de la escala (cargas de familia)
'
'        'Busco la declaracion Jurada
'        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!Ternro & _
'                 " AND desano = " & Ret_ano & _
'                 " AND itenro =" & rs_item!itenro
'        OpenRecordset StrSql, rs_desmen
'         Do While Not rs_desmen.EOF
'            If Month(rs_desmen!desfechas) <= Ret_mes Then
'                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Month(rs_desmen!desfechas) - Month(rs_desmen!desfecdes) + 1)
'            Else
'                If Month(rs_desmen!desfecdes) <= Ret_mes Then
'                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
'                End If
'            End If
'
'            rs_desmen.MoveNext
'         Loop
'
'        If Items_DDJJ(rs_item!itenro) > 0 Then
'            StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
'                     " AND vimes = " & Ret_mes & _
'                     " AND itenro =" & rs_item!itenro
'            OpenRecordset StrSql, rs_valitem
'             Do While Not rs_valitem.EOF
'                Items_TOPE(rs_item!itenro) = rs_valitem!vimonto / Ret_mes * Items_DDJJ(rs_item!itenro)
'
'                rs_valitem.MoveNext
'             Loop
'        End If
'    ' End case 4
'    ' ------------------------------------------------------------------------
'
'    Case 5:
'        i = 1
'        Hasta = IIf(50 > Len(rs_item!iteitemstope), 50, rs_item!iteitemstope)
'        Terminar = False
'        Do While i <= Hasta And Not Terminar
'            pos1 = i
'            pos2 = InStr(i, rs_item!iteitemstope, ",") - 1
'            If pos2 > 0 Then
'                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
'            Else
'                pos2 = Len(rs_item!iteitemstope)
'                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
'                Terminar = True
'            End If
'
'            If Texto <> "" Then
'                If Mid(Texto, 1, 1) = "-" Then
'                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2)
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
'                Else
'                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2)
'                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
'                End If
'            End If
'            i = pos2 + 2
'        Loop
'
'
'        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) * rs_item!iteporctope / 100
'
'
'        'Busco la declaracion Jurada
'        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!Ternro & _
'                 " AND desano = " & Ret_ano & _
'                 " AND itenro =" & rs_item!itenro
'        OpenRecordset StrSql, rs_desmen
'         Do While Not rs_desmen.EOF
'            If Month(rs_desmen!desfecdes) <= Ret_mes Then
'                Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
'            Else
'                Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
'            End If
'
'            rs_desmen.MoveNext
'         Loop
'
'
'        Aux_Inicio_Ret = CDate("01/01/ " & Ret_ano)
'        Aux_Fin_Ret = DateAdd("d", -1, CDate("01/" & (Ret_mes + 1) & "/" & Ret_ano))
'
'        'Busco las liquidaciones anteriores
'        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
'                 " AND empleado = " & buliq_empleado!Ternro & _
'                 " AND dlfecha >= " & ConvFecha(Aux_Inicio_Ret) & _
'                 " AND dlfecha <= " & ConvFecha(Aux_Fin_Ret)
'        OpenRecordset StrSql, rs_desliq
'
'        Do While Not rs_desliq.EOF
'            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto
'
'            rs_desliq.MoveNext
'        Loop
'
'        'Busco los acumuladores de la liquidacion
'        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
'        StrSql = "SELECT * FROM itemacum " & _
'                 " WHERE itenro =" & rs_item!itenro & _
'                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemacum
'        Do While Not rs_itemacum.EOF
'            Acum = CStr(rs_itemacum!acunro)
'            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
'                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
'
'                If CBool(rs_itemacum!itasigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
'                End If
'            End If
'            rs_itemacum.MoveNext
'        Loop
'        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
'
''        'Busco los acumuladores de la liquidacion
''        StrSql = "SELECT * FROM itemacum " & _
''                 " INNER JOIN acu_liq ON itemacum.acunro = acu_liq.acunro " & _
''                 " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
''                 " AND itenro =" & rs_item!itenro & _
''                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
''        OpenRecordset StrSql, rs_itemacum
''
''        Do While Not rs_itemacum.EOF
''                If CBool(rs_itemacum!itasigno) Then
''                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemacum!almonto
''                Else
''                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemacum!almonto
''                End If
''
''            rs_itemacum.MoveNext
''        Loop
'
'        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
'        ' Busco los conceptos de la liquidacion
'        StrSql = "SELECT * FROM itemconc " & _
'                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
'                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND itemconc.itenro =" & rs_item!itenro & _
'                 " AND (itemconc.itaconcnrodest is null OR itemconc.itaconcnrodest = " & Con_liquid & ")"
'        OpenRecordset StrSql, rs_itemconc
'
'        Do While Not rs_itemconc
'                If CBool(rs_itemconc!itcsigno) Then
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
'                Else
'                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
'                End If
'
'            rs_itemconc.MoveNext
'        Loop
'
'        'TOPEO LOS VALORES
'        If Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro) < Items_TOPE(rs_item!itenro) Then
'            Items_TOPE(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro)
'        End If
'    ' End case 5
'    ' ------------------------------------------------------------------------
'    Case Else:
'    End Select
'
'    If rs_item!itenro > 7 Then
'        Items_TOPE(rs_item!itenro) = Abs(Items_TOPE(rs_item!itenro))
'    End If
'
'    'Armo la traza del item
'    If HACE_TRAZA Then
'        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-DDJJ"
'        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, Texto, Items_DDJJ(rs_item!itenro))
'        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Liq"
'        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, Texto, Items_LIQ(rs_item!itenro))
'        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-LiqAnt"
'        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, Texto, Items_OLD_LIQ(rs_item!itenro))
'        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Tope"
'        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, Texto, Items_TOPE(rs_item!itenro))
'    End If
'
'    'Calcula la Ganancia Imponible
'    If CBool(rs_item!itesigno) Then
'        Gan_Imponible = Gan_Imponible + Items_TOPE(rs_item!itenro)
'    Else
'        If (rs_item!itetipotope = 1) Or (rs_item!itetipotope = 4) Then
'            Ded_a23 = Ded_a23 - Items_TOPE(rs_item!itenro)
'        Else
'            Deducciones = Deducciones - Items_TOPE(rs_item!itenro)
'        End If
'    End If
'
'
'    rs_item.MoveNext
'Loop
'
'If HACE_TRAZA Then
'    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "Ganancia Neta", Gan_Imponible)
'    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "Total de deducciones", Deducciones)
'End If
'
'
''Calculo el porcentaje de deduccion segun la ganancia neta.
'Gan_Neta = Gan_Imponible
'Imp_Neto = Imp_Neto - Ajustecargas
'
'
'StrSql = "SELECT * FROM escala_ded " & _
'         " WHERE esd_topeinf <= " & (Gan_Imponible / Ret_mes * 12) & _
'         " AND esd_topesup >=" & (Gan_Imponible / Ret_mes * 12)
'OpenRecordset StrSql, rs_escala_ded
'
'If Not rs_escala_ded.EOF Then
'    Por_Deduccion = rs_escala_ded!esd_porcentaje
'Else
'    If HACE_TRAZA Then
'        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "No hay esc. dedu para", Gan_Imponible)
'    End If
'    ' No se ha encontrado la escala de deduccion para el valor gan_imponible
'End If
'
'If HACE_TRAZA Then
'    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "9- % a tomar deduc.", Por_Deduccion)
'End If
'
''Aplico el porcentaje a las deducciones
'If Ret_ano >= 2000 Then
'    Ded_a23 = Ded_a23 * Por_Deduccion / 100
'End If
'
''Calculo la ganancia imponible
'Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23
'If HACE_TRAZA Then
'    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "9- Ganancia Imponible", Gan_Imponible)
'End If
'
'
'If Gan_Imponible > 0 Then
'    'Entrar en la escala con las ganancias acumuladas
'    StrSql = "SELECT * FROM escala " & _
'             " WHERE escmes =" & Ret_mes & _
'             " AND escano =" & Ret_ano & _
'             " AND escinf <= " & Gan_Imponible & _
'             " AND escsup >= " & Gan_Imponible
'    OpenRecordset StrSql, rs_escala
'
'    If Not rs_escala.EOF Then
'        Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
'    Else
'        Impuesto_Escala = 0
'    End If
'Else
'    Impuesto_Escala = 0
'End If
'
'
'' Calculo las retenciones ya realizadas
'Ret_ano = 0
''For each ficharet where ficharet.empleado = buliq-empleado.ternro
''                    And Month(ficharet.fecha) <= ret-mes
''                    And Year(ficharet.fecha) = ret-ano NO-LOCK:
''    Assign Ret-ant = Ret-Ant + ficharet.importe.
''End.
''como no puede utilizar la funcion month() en sql
''levanto todas las ficharet del tercero y hago la pregunta dentro del loop
'StrSql = "SELECT * FROM ficharet " & _
'         " WHERE empleado =" & buliq_empleado!Ternro
'OpenRecordset StrSql, rs_ficharet
'
'Do While Not rs_ficharet.EOF
'    If (Month(rs_ficharet!Fecha) = Ret_mes) And (Year(rs_ficharet!Fecha) = Ret_ano) Then
'        Ret_Ant = Ret_Ant + rs_ficharet!Importe
'    End If
'    rs_ficharet.MoveNext
'Loop
'
''Calcular la retencion
'Retencion = Impuesto_Escala - Ret_Ant
'
'If Gan_Imponible > 0 Then
'    'Calcular el grossing
'    If rs_escala.EOF Then
'        Topeescala = rs_escala!escsup
'        Porcant = rs_escala!escporexe
'    Else
'        Topeescala = 0
'        Porcant = 0
'    End If
'    Imp_Neto = Imp_Neto - Retencion
'
'
'    ' Empieza el cambio
'        StrSql = "SELECT * FROM escala " & _
'                 " WHERE escmes =" & Ret_mes & _
'                 " AND escano =" & Ret_ano & _
'                 " AND escinf <= " & (Gan_Imponible - Imgrossing) & _
'                 " AND escsup >= " & (Gan_Imponible - Imgrossing)
'        OpenRecordset StrSql, rs_escala
'        If Not rs_escala.EOF Then
'            Impuesto_Escala2 = rs_escala!esccuota + ((Gan_Imponible - Imgrossing - rs_escala!escinf) * rs_escala!escporexe / 100)
'            auxi = rs_escala!escporexe
'        Else
'            Impuesto_Escala2 = 0
'        End If
'
'
'    ' Termina el cambio
'        Grossingup = (Impuesto_Escala - Impuesto_Escala2)
'
'        If HACE_TRAZA Then
'            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "GR - Gcias sin importe", rs_escala!escporexe)
'            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "GR - Escala ded sin importe", Por_Deduccion)
'            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "GR - Importe Grossing", Grossingup)
'            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "GR - Grossing 1", Retencion)
'            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "GR - % aportes", Porcapo)
'        End If
'
'
'        If Items_LIQ(i) >= AMPO Then
'            Ajustecargas = 0
'        Else
'            Difcargas = Grossingup - ((Items_LIQ(i) + Grossingup) - AMPO)
'            Ajustecargas = ((Grossingup - Difcargas) / ((100 / Porcapo) / 100)) - (Grossingup - Difcargas)
'            Retencion = ((Grossingup + Ajustecargas) / ((100 - rs_escala!escporexe) / 100)) - Grossingup
'        End If
'
'        Grossingup = Grossingup + Retencion
'
'        ' Me fijo si se pasa de la escala
'        If Retencion > 0 And (Gan_Imponible + Grossingup) > rs_escala!escsup Then
'            Auxi2 = Gan_Imponible + Grossingup - rs_escala!escsup
'            Auxi2 = rs_escala!escsup - Gan_Imponible
'            Grossingup = Grossingup - Retencion
'
'            StrSql = "SELECT * FROM escala " & _
'                     " WHERE escmes =" & Ret_mes & _
'                     " AND escano =" & Ret_ano & _
'                     " AND escinf <= " & (Gan_Imponible + Grossingup + Retencion) & _
'                     " AND escsup >= " & (Gan_Imponible + Grossingup + Retencion)
'            OpenRecordset StrSql, rs_escala
'
'            Retencion = ((Grossingup + Ajustecargas) / ((100 - rs_escala!escporexe) / 100)) - (Grossingup + Ajustecargas)
'
'            Auxi2 = (((Auxi2) / ((100 - rs_escala!escporexe + auxi) / 100)) - Auxi2)
'
'            If Gan_Imponible < rs_escala!escinf Then
'                Retencion = Retencion - ((Auxi2) / ((100 - rs_escala!escporexe) / 100))
'            End If
'        Else
'            Retencion = 0
'        End If
'
'        If HACE_TRAZA Then
'            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "GR - Grossing 2", Retencion)
'            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "GR - % impto", rs_escala!escporexe)
'        End If
'
'
'        ' esto es para compenzar los cambios en los porcentajes de deducciones
'        If Grossingup + Retencion > 0 Then
'            StrSql = "SELECT * FROM escala_ded " & _
'                     " WHERE esd_topeinf <= " & ((Gan_Neta + Grossingup + Retencion) / Ret_mes * 12) & _
'                     " AND esd_topesup >=" & ((Gan_Neta + Grossingup + Retencion) / Ret_mes * 12)
'            OpenRecordset StrSql, rs_escala_ded
'
'            If Not rs_escala_ded.EOF Then
'                If Por_Deduccion <> rs_escala_ded!esd_porcentaje Then
'                    Ded_a23 = Ded_a23 / Por_Deduccion * 100
'                    baja_ded = Ded_a23 * (Por_Deduccion - rs_escala_ded!esd_porcentaje) / 100
'                    ' lo que queda por encima del tope
'                    baja2 = IIf((Gan_Imponible + Retencion + baja_ded > Topeescala) And (Gan_Imponible + Retencion < Topeescala), Gan_Imponible + Retencion + baja_ded - Topeescala, 0)
'                    baja1 = (baja_ded - baja2)
'                    baja2 = baja2 * (1 / (1 - (Porcant / 100)) - 1)
'                    baja1 = baja1 * (1 / (1 - (rs_escala!escporexe / 100)) - 1)
'                    Ajustededucc = baja1 + baja2
'                End If
'            Else
'                ' No se ha encontrado la escala de deduccion para el valor (gan_neta + retencion)
'                If HACE_TRAZA Then
'                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "No Hay esc. Dedu para ", Gan_Neta)
'                End If
'            End If
'
'            If HACE_TRAZA Then
'                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "9 - dif. por camb. deduc.", baja_ded)
'                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "9 - dif. por camb. deduc. % ", rs_escala_ded!esd_porcentaje)
'                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "9 - Ajuste por ded. 1", baja1)
'                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "9 - Ajuste por ded. 2", baja2)
'            End If
'
'        End If 'If Grossingup + Retencion > 0 Then
'        ' fin de compenzacion por los cambios en los porcentajes de deducciones
'
'    End If 'If Gan_Imponible > 0 Then
'
'    If HACE_TRAZA Then
'        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "Retenciones anteriores", Ret_Ant)
'        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "Impuesto por Escala", Impuesto_Escala)
'        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_concepto(concepto_actual).concnro, 0, "A Retener/Devolver", Retencion)
'    End If
'
'    If (Grossingup + Retencion + Ajustecargas) > 0 Then
'        Monto = ((Grossingup + Retencion + Ajustecargas) * (Porc_Gross / 100)) - Ajustededucc
'    End If
'    Bien = True
'End Function

Public Function for_Grossing5() As Single
' _____________________________________________________________________________________________
' Descripcion           : Programa para el clculo de la Retenci¢n de Ganancias (Grossing)
' Autor                 :
' Fecha                 : Version Nueva 26_11_2002
' Traducido por         : FGZ
' Parametros de entrada :
' Ultima Mod.           :
' Descripcion           :
' _____________________________________________________________________________________________

' Tipos de parametros usados

Dim p_Ampo        As Integer ' 1000   /* Valor del tope ampo */
Dim p_Devuelve    As Integer ' 1001   /* Si devuelve ganancias o no */
Dim p_Tope_Gral   As Integer ' 1002   /* Tope Gral de Retenci¢n */
Dim p_Neto        As Integer ' 1003   /* Base para el tope */
Dim p_Porcapo     As Integer ' 1004   /* Porcentaje de aporte de cargas sociales */
Dim p_Diciembre   As Integer ' 1005   /* Toma la escala de diciembre */
Dim p_Porcgross   As Integer ' 1010   /* Porcentaje de grossing a efectuar */
Dim p_Imgrossing  As Integer ' 51     /* Sueldo Neto pactado */
Dim p_prorratea As Integer  'Si prorratea o no para liq. finales'

'/*  Variables Locales */
Dim AMPO             As Single ' 4800
Dim Porcapo          As Single ' 0
Dim Devuelve         As Single ' 1
Dim Tope_Gral        As Single ' 100
'
Dim prorratea        As Single
Dim Imp_Neto         As Single ' 0
Dim Neto             As Single ' 999999
Dim Retencion        As Single ' 0
Dim Gan_Imponible    As Single ' 0
Dim Gan_Neta         As Single ' 0
Dim Deducciones      As Single ' 0
Dim Ded_a23          As Single ' 0
Dim Por_Deduccion    As Single ' 0
Dim Impuesto_Escala As Single  ' 0
Dim Ret_Ant          As Single ' 0
Dim Porc_Gross       As Single ' 100
Dim Ajustecargas     As Single ' 0
Dim Difcargas     As Single ' 0
Dim Ajustededucc     As Single ' 0
Dim Topeescala       As Single ' 0
Dim Masgross         As Single ' 0
Dim M2asgross        As Single ' 0
Dim M3asgross        As Single ' 0
Dim M4asgross        As Single ' 0
Dim M5asgross        As Single ' 0
Dim Porcant          As Single ' 0
Dim Diciembre        As Single ' 0
Dim Imgrossing       As Single ' 0
Dim Grossingup       As Single ' 0

Dim Ret_mes          As Integer
Dim Ret_ano          As Integer
Dim fin_mes_ret As Date
Dim ini_anyo_ret As Date
Dim Aux_Fin_Ret As Date
Dim Aux_Inicio_Ret As Date
Dim Con_liquid       As Integer
Dim i                As Integer
Dim Texto            As String

'/* vectores para mejorar el proceso */
Dim Items_DDJJ(50)       As Single
Dim Items_LIQ(50)        As Single
Dim Items_OLD_LIQ(50)    As Single
Dim Items_TOPE(50)       As Single
Dim Items_PRORR(50)      As Single


Dim Impuesto_Escala2 As Single  ' 0
Dim Diferencia As Single
Dim auxi As Single
Dim Auxi2 As Single

Dim baja_ded As Single
Dim baja1 As Single
Dim baja2 As Single

'Recorsets Auxiliares
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Ampo As New ADODB.Recordset
Dim rs_WF_impproarg As New ADODB.Recordset
Dim rs_ImpMesArg As New ADODB.Recordset
Dim rs_AmpoConTpa As New ADODB.Recordset
Dim rs_item As New ADODB.Recordset
Dim rs_AcuLiq As New ADODB.Recordset
Dim rs_valitem As New ADODB.Recordset
Dim rs_desmen As New ADODB.Recordset
Dim rs_desliq As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_escala_ded As New ADODB.Recordset
Dim rs_escala As New ADODB.Recordset
Dim rs_ficharet As New ADODB.Recordset
Dim Hasta As Integer
' FGZ - 27/02/2004
Dim Terminar As Boolean
Dim pos1
Dim pos2
' FGZ - 27/02/2004

' FGZ - 12/02/2004
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Single
' FGZ - 12/02/2004

'Inicializaciones
p_Ampo = 1000
p_Devuelve = 1001
p_Tope_Gral = 1002
p_Neto = 1003
p_Porcapo = 1004
p_Diciembre = 1005
p_Porcgross = 1010
p_Imgrossing = 51
p_prorratea = 1005

AMPO = 4800
Porcapo = 0
Devuelve = 1
Tope_Gral = 100

Imp_Neto = 0
Neto = 999999
Retencion = 0
Gan_Imponible = 0
Gan_Neta = 0
Deducciones = 0
Ded_a23 = 0
Por_Deduccion = 0
Impuesto_Escala = 0
Ret_Ant = 0
Porc_Gross = 100
Ajustecargas = 0
Difcargas = 0
Ajustededucc = 0
Topeescala = 0
Masgross = 0
M2asgross = 0
M3asgross = 0
M4asgross = 0
M5asgross = 0
Porcant = 0
Diciembre = 0
Imgrossing = 0
Grossingup = 0

Impuesto_Escala2 = 0
Diferencia = 0

' Comienzo
Bien = False
exito = False

StrSql = "SELECT * FROM acu_liq WHERE acunro  = 6 "
OpenRecordset StrSql, rs_AcuLiq
If Not rs_AcuLiq.EOF Then
    Imp_Neto = rs_AcuLiq!almonto
Else
    Imp_Neto = 0
End If

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

Ret_mes = Month(buliq_proceso!profecpago)
Ret_ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_mes = 12, CDate("01/01/" & Ret_ano + 1) - 1, CDate("01/" & Ret_mes + 1 & "/" & Ret_ano) - 1)
ini_anyo_ret = CDate("01/01/" & Ret_ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).concnro

'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha =" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case p_Ampo:
        AMPO = rs_wf_tpa!Valor
    Case p_Devuelve:
        Devuelve = rs_wf_tpa!Valor
    Case p_Tope_Gral:
        Tope_Gral = rs_wf_tpa!Valor
    Case p_Neto:
        Neto = rs_wf_tpa!Valor
    Case p_Porcgross:
        Porc_Gross = rs_wf_tpa!Valor
    Case p_Diciembre:
        Diciembre = rs_wf_tpa!Valor
    Case p_Imgrossing:
        Imgrossing = rs_wf_tpa!Valor
    Case p_Porcapo:
        Porcapo = rs_wf_tpa!Valor
    Case p_prorratea:
        prorratea = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If Diciembre <> 1 Then
    Ret_mes = 12
    'FGZ - 27/09/2004
    fin_mes_ret = IIf(Ret_mes = 12, CDate("01/01/" & Ret_ano + 1) - 1, CDate("01/" & Ret_mes + 1 & "/" & Ret_ano) - 1)
End If

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Mes de Retencion " & Ret_mes
    Flog.writeline Espacios(Tabulador * 1) & "Año de Retencion " & Ret_ano
    Flog.writeline Espacios(Tabulador * 1) & "Fin mes de Retencion " & fin_mes_ret
    
    Flog.writeline Espacios(Tabulador * 1) & "Máxima Ret. en %" & Tope_Gral
    Flog.writeline Espacios(Tabulador * 1) & "Neto del Mes" & Neto
End If

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Tope_Gral, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Neto, "Neto del Mes", Neto)
End If


' Recorro todos los items de Ganancias
StrSql = "SELECT * FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_item

Do While Not rs_item.EOF
    Select Case rs_item!itetipotope
    Case 1: ' el valor a tomar es lo que dice la escala
    
        StrSql = "SELECT * FROM valitem WHERE viano =" & Ret_ano & _
                 " AND itenro=" & rs_item!itenro & _
                 " AND vimes =" & Ret_mes
        OpenRecordset StrSql, rs_valitem
        
        Do While Not rs_valitem.EOF
            Items_DDJJ(rs_valitem!itenro) = rs_valitem!vimonto
            Items_TOPE(rs_valitem!itenro) = rs_valitem!vimonto
            
            rs_valitem.MoveNext
        Loop
    ' End case 1
    ' ------------------------------------------------------------------------
    
    Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
        ' Busco la declaracion jurada
        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!ternro & _
                 " AND desano=" & Ret_ano & _
                 " AND itenro = " & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
        
        Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_item!itenro = 3 Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                Else
                    If rs_desmen!desmenprorra = 0 Then 'no es parejito
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + rs_desmen!desmondec
                    Else
                        Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                    End If
                End If
            End If
            
            
            rs_desmen.MoveNext
        Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto
            'Si el desliq prorratea debo proporcionarlo
            Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 Or CBool(rs_desliq!dlprorratea)) And (prorratea = 1), rs_desliq!dlmonto / (13 - Month(rs_desliq!dlfecha)) * (Ret_mes - Month(rs_desliq!dlfecha) + 1), rs_desliq!dlmonto)

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itaprorratea) Then
                    If CBool(rs_itemacum!itasigno) Then
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    End If
                Else
                    If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    Else
                        Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                    End If
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) Then
                If CBool(rs_itemconc!itcsigno) Then
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_PRORR(rs_item!itenro) = Items_PRORR(rs_item!itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                End If
            Else
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                End If
            End If
        
            rs_itemconc.MoveNext
        Loop
    ' End case 2
    ' ------------------------------------------------------------------------
        
' ****************************************************************************
' *  OJO QUEDA PENDIENTE EL PRORRATEO PARA LOS ITEMS DE TIPO 3 Y 5           *
' ****************************************************************************


     Case 3: 'TOMO LOS VALORES DE LA DDJJ Y LIQUIDACION Y EL TOPE PARA APLICARLO
     
        StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                 " AND vimes = " & Ret_mes & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_valitem
         Do While Not rs_valitem.EOF
            Items_TOPE(rs_item!itenro) = rs_valitem!vimonto
         
            rs_valitem.MoveNext
         Loop
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
         
            rs_desmen.MoveNext
         Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
        
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_item!itesigno) Then
            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
        End If
        
    ' End case 3
    ' ------------------------------------------------------------------------
   
    Case 4:
        ' Tomo los valores de la DDJJ y el valor de la escala (cargas de familia)
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfechas) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Month(rs_desmen!desfechas) - Month(rs_desmen!desfecdes) + 1)
            Else
                If Month(rs_desmen!desfecdes) <= Ret_mes Then
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec * (Ret_mes - Month(rs_desmen!desfecdes) + 1)
                End If
            End If
        
            rs_desmen.MoveNext
         Loop
        
        If Items_DDJJ(rs_item!itenro) > 0 Then
            StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_ano & _
                     " AND vimes = " & Ret_mes & _
                     " AND itenro =" & rs_item!itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_item!itenro) = rs_valitem!vimonto / Ret_mes * Items_DDJJ(rs_item!itenro)
             
                rs_valitem.MoveNext
             Loop
        End If
    ' End case 4
    ' ------------------------------------------------------------------------
        
    Case 5:
        i = 1
        Hasta = IIf(50 > Len(rs_item!iteitemstope), 50, rs_item!iteitemstope)
        Terminar = False
        Do While i <= Hasta And Not Terminar
            pos1 = i
            pos2 = InStr(i, rs_item!iteitemstope, ",") - 1
            If pos2 > 0 Then
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
            Else
                pos2 = Len(rs_item!iteitemstope)
                Texto = Mid(rs_item!iteitemstope, pos1, pos2 - pos1 + 1)
                Terminar = True
            End If
            
            If Texto <> "" Then
                If Mid(Texto, 1, 1) = "-" Then
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                Else
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                End If
            End If
            i = pos2 + 2
        Loop
        
        Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) * rs_item!iteporctope / 100
    
    
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_ano & _
                 " AND itenro =" & rs_item!itenro
        OpenRecordset StrSql, rs_desmen
         Do While Not rs_desmen.EOF
            If Month(rs_desmen!desfecdes) <= Ret_mes Then
                If rs_desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
                Else
                    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                End If
            End If
            rs_desmen.MoveNext
         Loop
    
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_item!itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_desliq

        Do While Not rs_desliq.EOF
            Items_OLD_LIQ(rs_item!itenro) = Items_OLD_LIQ(rs_item!itenro) + rs_desliq!dlmonto

            rs_desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro=" & rs_item!itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acunro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - Aux_Acu_Monto
                End If
            End If
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
        ' Busco los conceptos de la liquidacion
        StrSql = "SELECT * FROM itemconc " & _
                 " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
                 " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_item!itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_item!itenro) = Items_LIQ(rs_item!itenro) - rs_itemconc!dlimonto
                End If
        
            rs_itemconc.MoveNext
        Loop
' FGZ - 22/06/2004
'        'TOPEO LOS VALORES
'        If Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro) < Items_TOPE(rs_item!itenro) Then
'            Items_TOPE(rs_item!itenro) = Items_LIQ(rs_item!itenro) + Items_OLD_LIQ(rs_item!itenro) + Items_DDJJ(rs_item!itenro)
'        End If

' FGZ - 22/06/2004
' puse lo mismo que para el itemtope 3
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro)) < Items_TOPE(rs_item!itenro) Then
            Items_TOPE(rs_item!itenro) = Abs(Items_LIQ(rs_item!itenro)) + Abs(Items_OLD_LIQ(rs_item!itenro)) + Abs(Items_DDJJ(rs_item!itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_item!itesigno) Then
            Items_TOPE(rs_item!itenro) = -Items_TOPE(rs_item!itenro)
        End If

    ' End case 5
    ' ------------------------------------------------------------------------
    Case Else:
    End Select

    If rs_item!itenro > 7 Then
        Items_TOPE(rs_item!itenro) = Abs(Items_TOPE(rs_item!itenro))
    End If
    
    'Armo la traza del item
    If HACE_TRAZA Then
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_DDJJ(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_LIQ(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_OLD_LIQ(rs_item!itenro))
        Texto = CStr(rs_item!itenro) & "-" & rs_item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_TOPE(rs_item!itenro))
    End If
        
    'Calcula la Ganancia Imponible
    If CBool(rs_item!itesigno) Then
        Gan_Imponible = Gan_Imponible + Items_TOPE(rs_item!itenro)
    Else
        If (rs_item!itetipotope = 1) Or (rs_item!itetipotope = 4) Then
            Ded_a23 = Ded_a23 - Items_TOPE(rs_item!itenro)
        Else
            Deducciones = Deducciones - Items_TOPE(rs_item!itenro)
        End If
    End If
            
            
    rs_item.MoveNext
Loop

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Ganancia Neta", Gan_Imponible)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Total de deducciones", Deducciones)
End If


'Calculo el porcentaje de deduccion segun la ganancia neta.
Gan_Neta = Gan_Imponible
Imp_Neto = Imp_Neto - Ajustecargas


StrSql = "SELECT * FROM escala_ded " & _
         " WHERE esd_topeinf <= " & (Gan_Imponible / Ret_mes * 12) & _
         " AND esd_topesup >=" & (Gan_Imponible / Ret_mes * 12)
OpenRecordset StrSql, rs_escala_ded

If Not rs_escala_ded.EOF Then
    Por_Deduccion = rs_escala_ded!esd_porcentaje
Else
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "No hay esc. dedu para", Gan_Imponible)
    End If
    ' No se ha encontrado la escala de deduccion para el valor gan_imponible
End If

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- % a tomar deduc.", Por_Deduccion)
End If

'Aplico el porcentaje a las deducciones
If Ret_ano >= 2000 Then
    Ded_a23 = Ded_a23 * Por_Deduccion / 100
End If

'Calculo la ganancia imponible
Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Imponible", Gan_Imponible)
End If
        
           
If Gan_Imponible > 0 Then
    'Entrar en la escala con las ganancias acumuladas
    StrSql = "SELECT * FROM escala " & _
             " WHERE escmes =" & Ret_mes & _
             " AND escano =" & Ret_ano & _
             " AND escinf <= " & Gan_Imponible & _
             " AND escsup >= " & Gan_Imponible
    OpenRecordset StrSql, rs_escala
    
    If Not rs_escala.EOF Then
        Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
    Else
        Impuesto_Escala = 0
    End If
Else
    Impuesto_Escala = 0
End If
        
        
' Calculo las retenciones ya realizadas
Ret_Ant = 0
'For each ficharet where ficharet.empleado = buliq-empleado.ternro
'                    And Month(ficharet.fecha) <= ret-mes
'                    And Year(ficharet.fecha) = ret-ano NO-LOCK:
'    Assign Ret-ant = Ret-Ant + ficharet.importe.
'End.
'como no puede utilizar la funcion month() en sql
'levanto todas las ficharet del tercero y hago la pregunta dentro del loop
StrSql = "SELECT * FROM ficharet " & _
         " WHERE empleado =" & buliq_empleado!ternro
OpenRecordset StrSql, rs_ficharet

Do While Not rs_ficharet.EOF
    If (Month(rs_ficharet!Fecha) = Ret_mes) And (Year(rs_ficharet!Fecha) = Ret_ano) Then
        Ret_Ant = Ret_Ant + rs_ficharet!Importe
    End If
    rs_ficharet.MoveNext
Loop

'Calcular la retencion
Retencion = Impuesto_Escala - Ret_Ant

If Gan_Imponible > 0 Then
    'Calcular el grossing
    If rs_escala.EOF Then
        Topeescala = rs_escala!escsup
        Porcant = rs_escala!escporexe
    Else
        Topeescala = 0
        Porcant = 0
    End If
    Imp_Neto = Imp_Neto - Retencion
    
    
    ' Empieza el cambio
        StrSql = "SELECT * FROM escala " & _
                 " WHERE escmes =" & Ret_mes & _
                 " AND escano =" & Ret_ano & _
                 " AND escinf <= " & (Gan_Imponible - Imgrossing) & _
                 " AND escsup >= " & (Gan_Imponible - Imgrossing)
        OpenRecordset StrSql, rs_escala
        If Not rs_escala.EOF Then
            Impuesto_Escala2 = rs_escala!esccuota + ((Gan_Imponible - Imgrossing - rs_escala!escinf) * rs_escala!escporexe / 100)
            auxi = rs_escala!escporexe
        Else
            Impuesto_Escala2 = 0
        End If
        
                
    ' Termina el cambio
        Grossingup = (Impuesto_Escala - Impuesto_Escala2)
        
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Gcias sin importe", rs_escala!escporexe)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Escala ded sin importe", Por_Deduccion)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Importe Grossing", Grossingup)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Grossing 1", Retencion)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - % aportes", Porcapo)
        End If
        
        
        If Items_LIQ(i) >= AMPO Then
            Ajustecargas = 0
            Retencion = ((Grossingup + Ajustecargas) / ((100 - rs_escala!escporexe) / 100)) - Grossingup
        Else
            If Items_LIQ(i) > AMPO Then
                Difcargas = Grossingup - ((Items_LIQ(i) + Grossingup) - AMPO)
                Ajustecargas = ((Grossingup - Difcargas) / ((100 / Porcapo) / 100)) - (Grossingup - Difcargas)
                Retencion = ((Grossingup + Ajustecargas) / ((100 - rs_escala!escporexe) / 100)) - Grossingup
            End If
        End If
        
        Grossingup = Grossingup + Retencion
        
        ' Me fijo si se pasa de la escala
        If Retencion > 0 And (Gan_Imponible + Grossingup) > rs_escala!escsup Then
            Auxi2 = Gan_Imponible + Grossingup - rs_escala!escsup
            Auxi2 = rs_escala!escsup - Gan_Imponible
            Grossingup = Grossingup - Retencion
            
            StrSql = "SELECT * FROM escala " & _
                     " WHERE escmes =" & Ret_mes & _
                     " AND escano =" & Ret_ano & _
                     " AND escinf <= " & (Gan_Imponible + Grossingup + Retencion) & _
                     " AND escsup >= " & (Gan_Imponible + Grossingup + Retencion)
            OpenRecordset StrSql, rs_escala
            
            Retencion = ((Grossingup + Ajustecargas) / ((100 - rs_escala!escporexe) / 100)) - (Grossingup + Ajustecargas)
            
            Auxi2 = (((Auxi2) / ((100 - rs_escala!escporexe + auxi) / 100)) - Auxi2)
            
            If Gan_Imponible < rs_escala!escinf Then
                Retencion = Retencion - ((Auxi2) / ((100 - rs_escala!escporexe) / 100))
            End If
        Else
            Retencion = 0
        End If
        
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - Grossing 2", Retencion)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "GR - % impto", rs_escala!escporexe)
        End If
        
        
        ' esto es para compenzar los cambios en los porcentajes de deducciones
        If Grossingup + Retencion > 0 Then
            StrSql = "SELECT * FROM escala_ded " & _
                     " WHERE esd_topeinf <= " & ((Gan_Neta + Grossingup + Retencion) / Ret_mes * 12) & _
                     " AND esd_topesup >=" & ((Gan_Neta + Grossingup + Retencion) / Ret_mes * 12)
            OpenRecordset StrSql, rs_escala_ded
                    
            If Not rs_escala_ded.EOF Then
                If Por_Deduccion <> rs_escala_ded!esd_porcentaje Then
                    Ded_a23 = Ded_a23 / Por_Deduccion * 100
                    baja_ded = Ded_a23 * (Por_Deduccion - rs_escala_ded!esd_porcentaje) / 100
                    ' lo que queda por encima del tope
                    baja2 = IIf((Gan_Imponible + Retencion + baja_ded > Topeescala) And (Gan_Imponible + Retencion < Topeescala), Gan_Imponible + Retencion + baja_ded - Topeescala, 0)
                    baja1 = (baja_ded - baja2)
                    baja2 = baja2 * (1 / (1 - (Porcant / 100)) - 1)
                    baja1 = baja1 * (1 / (1 - (rs_escala!escporexe / 100)) - 1)
                    Ajustededucc = baja1 + baja2
                End If
            Else
                ' No se ha encontrado la escala de deduccion para el valor (gan_neta + retencion)
                If HACE_TRAZA Then
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "No Hay esc. Dedu para ", Gan_Neta)
                End If
            End If
                    
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9 - dif. por camb. deduc.", baja_ded)
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9 - dif. por camb. deduc. % ", rs_escala_ded!esd_porcentaje)
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9 - Ajuste por ded. 1", baja1)
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9 - Ajuste por ded. 2", baja2)
            End If
                   
        End If 'If Grossingup + Retencion > 0 Then
        ' fin de compenzacion por los cambios en los porcentajes de deducciones
        
    End If 'If Gan_Imponible > 0 Then
        
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Retenciones anteriores", Ret_Ant)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Impuesto por Escala", Impuesto_Escala)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "A Retener/Devolver", Retencion)
    End If
    
    If (Grossingup + Retencion + Ajustecargas) > 0 Then
        Monto = ((Grossingup + Retencion + Ajustecargas) * (Porc_Gross / 100)) - Ajustededucc
    End If
    
    Bien = True
    exito = Bien
    for_Grossing5 = Monto
End Function


Public Function For_Baseacci() As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Programa para el calculo de la base para accidentes
'               Modificado por MB y AP agregando calculo de jornales
'               y personal que trabaja en fases separadas  (25/06/2001) para CCU
'               Modificado por Maxi y AP Agregando personal mensual
'               con menos de 12 meses de trabajo (24/10/2001) para  Ayling
' Autor      :
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim p_acumulador     As Integer 'Acumulador para la base de calculo
Dim p_Monto_Manual  As Integer 'Monto para calculo manual. Deshabilita el resto si es <> 0
Dim p_divisor       As Integer 'Divisor para el calculo (puede ser 360 / 365)
Dim p_Multiplicador As Integer 'Multiplicador para la base (30,4)
Dim p_can_meses     As Integer 'Cantidad de meses en el a¤o (12 o 13) en caso de no tener a¤o completo

'Variables Locales
Dim n_acumulador     As Integer
Dim monto_manual     As Single
Dim divisor          As Single
Dim multiplicador    As Single
Dim cant_meses       As Integer
Dim fecha_desde      As Date
Dim fecha_hasta      As Date
Dim Meses            As Integer
Dim Meses_Validos    As Integer
Dim Monto_Acum       As Single
Dim Monto_Mes        As Single
Dim Texto            As String
Dim mestope As Integer
Dim fectope As Date
Dim anio As Integer
Dim mes As Integer
Dim dia As Integer
Dim Resultado As Integer

Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Emp_Lic As New ADODB.Recordset
Dim rs_Accidente As New ADODB.Recordset
Dim rs_Lic_Accid As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset

Dim Encontro As Boolean

'Inicializacion
p_acumulador = 148
p_Monto_Manual = 51
p_divisor = 54
p_Multiplicador = 149
p_can_meses = 150
monto_manual = 0
divisor = 1
multiplicador = 1
cant_meses = 12
Meses = 0
Meses_Validos = 0
Monto_Acum = 0
Monto_Mes = 0

    
    exito = False
    Encontro = False
    
    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case p_acumulador:
            n_acumulador = rs_wf_tpa!Valor
        Case p_Monto_Manual:
            monto_manual = rs_wf_tpa!Valor
        Case p_divisor:
            divisor = rs_wf_tpa!Valor
        Case p_Multiplicador:
            multiplicador = rs_wf_tpa!Valor
        Case p_can_meses:
            cant_meses = rs_wf_tpa!Valor
        End Select
        
        rs_wf_tpa.MoveNext
    Loop

    fecha_hasta = buliq_periodo!pliqhasta
    fecha_desde = IIf(Month(buliq_periodo!pliqdesde) = 1, CDate("01/12/" & Year(buliq_periodo!pliqdesde) - 1), CDate("01/" & Month(buliq_periodo!pliqdesde) - 1 & "/" & Year(buliq_periodo!pliqdesde)))
                            
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "1 - Acumulador " & n_acumulador
        Flog.writeline Espacios(Tabulador * 1) & "1 - Monto manual " & monto_manual
        Flog.writeline Espacios(Tabulador * 1) & "1 - Divisor " & divisor
        Flog.writeline Espacios(Tabulador * 1) & "1 - Multiplicador " & multiplicador
        Flog.writeline Espacios(Tabulador * 1) & "1 - Cant Meses " & cant_meses
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_acumulador, "1- Acumulador", n_acumulador)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Monto_Manual, "1- Monto manual", monto_manual)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_divisor, "1- Divisor", divisor)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Multiplicador, "1- Multiplicador", multiplicador)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_can_meses, "1- Cant meses", cant_meses)
    End If


If monto_manual = 0 Then
    'Busco la licencia del empleado para ver el calculo
    StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & buliq_empleado!ternro & " )" & _
    StrSql = StrSql & " AND ( tdnro = 13 " 'Accidente a cargo de la empresa
    StrSql = StrSql & " OR tdnro = 14) " 'Accidente a cargo de la ART
    StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(fecha_hasta)
    StrSql = StrSql & " AND elfechahasta >= " & ConvFecha(fecha_desde) 'Dentro de los dos meses anteriores
    OpenRecordset StrSql, rs_Emp_Lic
    
    If Not rs_Emp_Lic.EOF Then
        StrSql = "SELECT * FROM lic_accid "
        StrSql = StrSql & " WHERE emp_licnro = " & rs_Emp_Lic!emp_licnro
        OpenRecordset StrSql, rs_Lic_Accid
        If rs_Lic_Accid.EOF Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "No se ha encontrado el complemento de la licencia " & rs_Emp_Lic!emp_licnro & " del empleado " & buliq_empleado!empleg
            End If
        Else
            StrSql = "SELECT * FROM so_accidente "
            StrSql = StrSql & " WHERE rs_accidente.accnro = " & rs_Lic_Accid!accnro
            OpenRecordset StrSql, rs_Accidente
                    
            If rs_Accidente.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 1) & "No se ha encontrado el accidente de la licencia " & rs_Emp_Lic!emp_licnro & " del empleado " & buliq_empleado!empleg
                End If
            Else
                If HACE_TRAZA Then
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 9, "2- Nro accidente ", rs_Accidente!accnro)
                End If

                StrSql = "SELECT * FROM periodo "
                StrSql = StrSql & " WHERE Periodo.pliqdesde <= " & ConvFecha(rs_Accidente!accfecha)
                StrSql = StrSql & " AND Periodo.pliqhasta >=" & ConvFecha(rs_Accidente!accfecha)
                OpenRecordset StrSql, rs_Periodo
                
                If rs_Periodo.EOF Then
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 1) & "No se ha encontrado un periodo de liquidacion para el accidente de la licencia " & rs_Emp_Lic!emp_licnro & " del empleado " & buliq_empleado!empleg
                    End If
                Else
                    'CALCULO DE ACIDENTES PARA JORNALES
                    If buliq_empleado!folinro = 2 Then
                        mestope = IIf(Month(rs_Accidente!accfecha) - 12 < 0, 12 + (Month(rs_Accidente!accfecha) - 12), Month(rs_Accidente!accfecha) - 12)
                        fectope = CDate(Day(rs_Accidente!accfecha) & "/" & mestope & "/" & Year(rs_Accidente!accfecha) - 1)
                            
                        Call bus_Antfases(rs_Accidente!accfecha, fectope, dia, mes, anio)
                        'RUN antfases.p (recid(buliq-empleado), accidente.accfecha, fectope, output dia, output mes, output anio)
                        divisor = (anio * 360) + (mes * 30) + dia
                    End If
                    '*********************************************************************************************

                    Do While Meses < 12 'busco los periodos anteriores para el calculo
                    
                        StrSql = "SELECT * FROM periodo "
                        StrSql = StrSql & " WHERE Periodo.pliqdesde < " & ConvFecha(rs_Accidente!accfecha)
                        StrSql = StrSql & " AND Periodo.pliqhasta <" & ConvFecha(rs_Accidente!accfecha)
                        StrSql = StrSql & " ORDER BY pliqanio, pliqmes"
                        OpenRecordset StrSql, rs_Periodo
                        
                        If rs_Periodo.EOF Then
                            Exit Do
                        Else
                            rs_Periodo.MoveLast
                            
                            Monto_Mes = 0
                            StrSql = "SELECT * FROM proceso "
                            StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
                            StrSql = StrSql & " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro "
                            StrSql = StrSql & " WHERE proceso.pliqnro = " & rs_Periodo!PliqNro
                            StrSql = StrSql & " AND cabliq.empleado = " & buliq_empleado!ternro
                            StrSql = StrSql & " AND acu_liq.acunro = " & n_acumulador
                            OpenRecordset StrSql, rs_Acu_Liq
                            Do While Not rs_Acu_Liq.EOF
                                  Monto_Mes = Monto_Mes + rs_Acu_Liq!almonto
                                rs_Acu_Liq.MoveNext
                            Loop
                            If HACE_TRAZA Then
                                Texto = Meses & "- Monto mes " & rs_Periodo!pliqmes & " del Anio " & rs_Periodo!pliqanio
                                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 9, Texto, Monto_Mes)
                            End If
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 1) & Meses & " Monto mes " & rs_Periodo!pliqmes & " del año " & rs_Periodo!pliqanio & " es : " & Monto_Mes
                            End If
                            
                            If Monto_Mes > 0 Then
                                Monto_Acum = Monto_Acum + Monto_Mes
                                Meses_Validos = Meses_Validos + 1
                                Bien = True
                                Monto = IIf(buliq_empleado!folinro = 2, Monto_Acum / divisor * multiplicador, Monto_Acum / Meses_Validos * cant_meses / divisor * multiplicador)
                            End If
                        End If
                        Meses = Meses + 1
                    Loop
                    
                    If Meses_Validos = 12 Then
                        Monto = Monto_Acum / divisor * multiplicador
                    Else
                        mestope = IIf(Month(rs_Accidente!accfecha) - 12 < 0, 12 + (Month(rs_Accidente!accfecha) - 12), Month(rs_Accidente!accfecha) - 12)
                        fectope = CDate(Day(rs_Accidente!accfecha) & "/" & mestope & "/" & Year(rs_Accidente!accfecha) - 1)
                        
                        Call bus_Antfases(rs_Accidente!accfecha, fectope, dia, mes, anio)
                        'RUN antfases.p (recid(buliq-empleado), accidente.accfecha, fectope, output dia, output mes, output anio).
                        divisor = (anio * 360) + (mes * 30) + dia
                        Monto = Monto_Acum / divisor * multiplicador
                    End If
                    
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 1) & "mto acum " & Monto_Acum & " " & divisor & " " & multiplicador & " " & Meses_Validos
                    End If
                End If 'If rs_Periodo.EOF Then
            End If 'If rs_accidente.EOF Then
        End If 'If rs_lic_accid.EOF Then
    End If 'If Not rs_emp_lic.EOF Then
Else
    'Tomo el monto y lo pongo como base
    exito = True
    For_Baseacci = Monto
End If


' cierro todo y libero
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
If rs_Emp_Lic.State = adStateOpen Then rs_Emp_Lic.Close
If rs_Accidente.State = adStateOpen Then rs_Accidente.Close
If rs_Lic_Accid.State = adStateOpen Then rs_Lic_Accid.Close
If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close

Set rs_wf_tpa = Nothing
Set rs_Emp_Lic = Nothing
Set rs_Accidente = Nothing
Set rs_Lic_Accid = Nothing
Set rs_Periodo = Nothing
Set rs_Acu_Liq = Nothing

End Function


Public Function for_Nivelar(ByVal AFecha As Date) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Programa para el calculo del nivelador a un valor del NETO (como acumulador).
' Autor      :
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Nivel    As Long 'INITIAL 62.
Dim Neto     As Long 'INITIAL 63.

'variables locales
Dim Valor_Nivel   As Single
Dim Valor_Neto    As Single

Dim entero        As Long
Dim Neto_Truncado As Single
Dim Decimales     As Single
Dim v_monto As Single

'Inicializacion
Nivel = 62
Neto = 63

Dim rs_wf_tpa As New ADODB.Recordset
Dim Encontro1 As Boolean
Dim Encontro2 As Boolean

    exito = False
    Encontro1 = False
    Encontro2 = False
    
    'Obtencion de los parametros de WorkFile
    StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
    OpenRecordset StrSql, rs_wf_tpa
    
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case Nivel:
            Valor_Nivel = rs_wf_tpa!Valor
            Encontro1 = True
        Case Neto:
            Valor_Neto = rs_wf_tpa!Valor
            Encontro2 = True
        End Select
        
        rs_wf_tpa.MoveNext
    Loop

    'si no se obtuvieron satisfactoriamente todos los parametros , error
    If Not (Encontro1 And Encontro2) Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "No se encontraron todos los parametros"
        End If
        Exit Function
    End If

    If Valor_Nivel >= 1 Then 'se redondea a cantidades enteras
        entero = Round((Valor_Neto / Valor_Nivel), 0)
        If ((Valor_Neto / Valor_Nivel) - entero) > 0 Then
            entero = entero + 1
        End If
        entero = entero * Valor_Nivel
        v_monto = Round(entero - Valor_Neto, 4)
    End If
    
    
    If (Valor_Nivel > 0 And Valor_Nivel < 1) Then 'se redondea con decimales
        Decimales = (Valor_Neto - Fix(Valor_Neto)) * 100
        Valor_Nivel = Valor_Nivel * 100
        v_monto = (Fix(Decimales / Valor_Nivel)) * Valor_Nivel - Decimales
        v_monto = IIf(v_monto = 0, 0, (Valor_Nivel + v_monto) / 100)
    End If
    
    exito = True
    for_Nivelar = v_monto
    
End Function
