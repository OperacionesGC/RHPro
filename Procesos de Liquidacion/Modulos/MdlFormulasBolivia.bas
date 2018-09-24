Attribute VB_Name = "MdlFormulasBolivia"
Option Explicit

' ---------------------------------------------------------
' Modulo de fórmulas conocidas para Bolivia
' ---------------------------------------------------------
'====================================================================================================================================
'Detalladamente su cálculo es el siguiente:
'A-  Ingresos Gravables:
'Primero calcula el total de los ingresos gravables del mes:
'HABER BASICO
'Bono ANTIGÜEDAD
'Horas EXTRAS
'Otros BONOS
'REINTEGRO
'
'
'b -Deducciones:
'Lo siguiente será determinar las Deducciones legales que se aplicaran al total de los ingresos gravables:
'b)  AFP RENTA DE VEJEZ
'            AFP RIESGO COMUN
'            AFP PROVISION COMISIONES
'            AFP APORTE NACIONAL SOLIDARIO
'            AFP FONDO SOLIDARIO LABORAL
'
'1   Sueldo Neto: Corresponde a la diferencia de (A-B), es decir el Total de Ingresos Gravables menos las Deducciones.
'
'2   Mínimo No Imponible: Es igual al valor de dos (2) Salarios Mínimos Nacionales (SMN) de Bolivia. Actualmente cada SMN asciende a $ 1400
'
'3   Determinación Base Imponible o Diferencia Sujeta a Impuesto: es la diferencia del Sueldo Neto menos el total de Deducciones. Si el resultado es menor a cero, aquí se detiene. Pero si la diferencia es mayor a cero, entonces pasar al punto 5)
'
'4   Impuesto RC-IVA13%: Obtenida la BASE IMPONIBLE O DIFERENCIA SUJETA A IMPUESTO, se procede a calcular el 13% de la misma.
'
'BASE IMPONIBLE O DIFERENCIA SUJETA A IMPUESTO x 13%= IMPUESTO 13%
'
'5   DDJJ - Formulario 110: Debe traer el ítem cargado como deducible según se especifica mas abajo.
'
'6   13% Sobre Mínimo No Imponible: Se debe calcular el 13% del Mínimo No Imponible
'
'MINIMO NO IMPONIBLE x 13%= 13% S/MINIMO NO IMPONIBLE
'2400                X 13 = 312, 0
'
'
'SI      Impuesto RC-IVA
'menos   DDJJ - Formulario 110
'menos   13% sobre Minimo No imponibe
'da resultado mayor a cero, entonces ítem 7
'da resultado menor a cero, entonces ítem 8
'
'
'
'7   Saldo a Favor Fisco: Surge como resultado de la operación anterior, siempre es un valor positivo.
'
'8   Saldo a Favor Dependiente: Surge como resultado de la operación anterior, siempre es un valor negativo, al cual debe multiplicarse por (-1).
'
'9   Saldo Anterior a Favor del Dependiente Mes Anterior: debe trasladar el saldo del mes anterior.
'
'10  Saldo Anterior a Favor del Dependiente Actualizado: surge de multiplicar el ITEM 9 por un factor. Dicho factor se calcula con la siguiente formula:
'
'Factor: (UFV Actual/UFV Anterior)/ UFV Anterior
'UFV: es el Indice llamado Unidad de Fomento de la Vivienda, el mismo es publicado mensualmente por el Banco Central de Bolivia.
'
'11  Saldo Anterior a Favor Total: Surge de la suma de los Items 9 + 10. Representa el Saldo Anterior a Favor mas su Actualización.
'
'12  Saldo Total a Favor del Dependiente: Surge de la suma de los Items 8+ 11.
'
'13  Saldo Utilizado: Representa el monto de Impuesto que se cancela con el Saldo a Favor.
'
'14  Impuesto Retenido a Pagar: surge de la diferencia entre Saldo a Favor del Fisco menos Saldo Utilizado.
'
'15  Saldo a Favor Próximo Mes: Resulta de la siguiente operación: ITEM 12 menos ITEM 13 menos ITEM 14.
'====================================================================================================================================




Public Function for_RC_IVA(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias para Bolivia.
' Autor      : FGZ
' Fecha      : 10/04/2014
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim I As Long
Dim Ret_Mes As Integer
Dim Ret_Ano As Long


Dim c_Acunro_Mensual    As Integer  'Acumulador
Dim c_Pension_Bruto     As Integer  'Pension es sobre bruto o liquido
Dim c_CP                As Integer  'Acumulador
Dim c_D                 As Integer  'Acumulador

Dim v_Pension_Bruto     As Boolean
Dim v_Acunro_Mensual    As Long
Dim v_CP                As Long
Dim v_D                 As Long

Dim P               As Double   'valor da pensão a ser paga
Dim RB              As Double   'rendimento bruto
Dim CP              As Double   'Contribuição Previdenciária
Dim T               As Double   'alíquota da faixa da tabela progressiva a que pertencer o rendimento bruto;
Dim D               As Double   'dedução de dependentes, caso o contribuinte tenha outros dependentes sob sua guarda, que não o beneficiário da pensão;
Dim PD              As Double   'parcela a deduzir correspondente à faixa da base de cálculo (da tabela progressiva) a que pertencer o rendimento bruto
Dim PA              As Double   'percentagem da pensão alimentícia fixada em juizo


'Dim rs_wf_tpa As New ADODB.Recordset

'Inicializo
Bien = False

c_Pension_Bruto = 1036
c_Acunro_Mensual = 75
c_CP = 99
c_D = 1025
v_Pension_Bruto = True  'Pension calculada sobre el bruto



''Primero limpio la traza
'StrSql = "DELETE FROM traza_gan WHERE "
'StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
'StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
'StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
'StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
'objConn.Execute StrSql, , adExecuteNoRecords

If HACE_TRAZA Then
    Call LimpiarTrazaConcepto(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro)
End If


'Obtencion de los parametros de WorkFile
For I = LI_WF_Tpa To LS_WF_Tpa
    Select Case Arr_WF_TPA(I).tipoparam
    Case c_Acunro_Mensual:
        v_Acunro_Mensual = CBool(Arr_WF_TPA(I).Valor)
    Case c_Pension_Bruto:
        v_Pension_Bruto = CBool(Arr_WF_TPA(I).Valor)
    Case c_CP:
        v_CP = CBool(Arr_WF_TPA(I).Valor)
    Case c_D:
        v_D = CBool(Arr_WF_TPA(I).Valor)
    'Case c_Escala:
    '    v_Escala = Arr_WF_TPA(I).Valor
    Case Else
    End Select
Next I

If CBool(USA_DEBUG) Then
    If v_Pension_Bruto Then
        Flog.writeline Espacios(Tabulador * 1) & "Pensão Alimenticia Bruto"
    Else
        Flog.writeline Espacios(Tabulador * 1) & "Pensão Alimenticia Líquido"
    End If
    'Flog.writeline Espacios(Tabulador * 1) & "Monto Directo " & v_mon_dir
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)

'Busco los parametros
RB = BuscarAcumulador(v_Acunro_Mensual, Ret_Ano, Ret_Mes)
Call Buscar_EscalaIRRF(Ret_Mes, Ret_Ano, RB, T, PD)

'Calculo basico
P = (RB * T) - PD

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 3) & "10 - Base Imponible: " & (RB)
    Flog.writeline Espacios(Tabulador * 3) & "10 - ALÍQUOTA: " & T
    Flog.writeline Espacios(Tabulador * 3) & "10 - PARCELA A DEDUZIR: " & PD
    Flog.writeline Espacios(Tabulador * 3) & "10 - IR: " & (P)
End If

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "10 - Base Imponible ", RB)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "10 - ALÍQUOTA ", T)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "10 - PARCELA A DEDUZIR ", PD)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "10 - IR ", P)
End If



'pension alimentaria
If v_Pension_Bruto Then 'Pension calculada sobre el bruto
    'En teorica no debo hacer nada

Else 'Pension calculada sobre el liquido (neto)
    CP = BuscarAcumulador(v_CP, Ret_Ano, Ret_Mes)
    D = BuscarAcumulador(v_D, Ret_Ano, Ret_Mes)
    
    PA = PensionAlimenticia()

    P = (RB - CP - ((T / 100) * (RB - CP - D - P)) + PD) * (PA / 100)
    
End If

for_RC_IVA = P
exito = Bien

'Cierro y libero todo
'If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
'Set rs_wf_tpa = Nothing

End Function




Public Function BuscarAcumulador(ByVal BrutoMensual As Long, ByVal Anio As Long, ByVal Mes As Integer) As Double
Dim rs_Acum As New ADODB.Recordset
Dim Acum_Maximo As Double

    StrSql = "SELECT acu_mes.ammonto monto FROM acu_mes "
    StrSql = StrSql & " WHERE acunro =" & BrutoMensual
    StrSql = StrSql & " AND ternro =  " & buliq_empleado!Ternro
    StrSql = StrSql & " AND amanio =  " & Anio & " AND ammes = " & Mes
    OpenRecordset StrSql, rs_Acum
    If Not rs_Acum.EOF Then
        If Not IsNull(rs_Acum!Monto) Then
            Acum_Maximo = rs_Acum!Monto
        Else
            Acum_Maximo = 0
        End If
    Else
        Acum_Maximo = 0
    End If

    'Si no tiene acum porque ingreso este mes debe tomar el sueldo actual
    If Acum_Maximo = 0 Then
        'busco los acu_liq del periodo actual
        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(BrutoMensual)) Then
            Acum_Maximo = objCache_Acu_Liq_Monto.Valor(CStr(BrutoMensual))
        End If
    End If

    BuscarAcumulador = Acum_Maximo

If rs_Acum.State = adStateOpen Then rs_Acum.Close
Set rs_Acum = Nothing
End Function



Public Sub Buscar_EscalaIRRF(ByVal Mes As Long, ByVal Anio As Long, ByVal Monto, ByRef Alicuota As Double, ByRef Parcela As Double)
' ---------------------------------------------------------------------------------------------
' Descripcion: Buscar Escalar .
' Autor      : FGZ
' Fecha      : 29/11/2014
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_escala As New ADODB.Recordset
Dim v_factor As Double
Dim v_rebaja As Double
    
v_factor = 0
v_rebaja = 0
    
    StrSql = "SELECT * FROM escala "
    StrSql = StrSql & " WHERE escano = " & Anio
    StrSql = StrSql & " AND escmes = " & Mes
    StrSql = StrSql & " AND escinf <= " & Monto & " AND escsup >= " & Monto
    StrSql = StrSql & " Order BY escano, escmes desc, escporexe"
    OpenRecordset StrSql, rs_escala
    If Not rs_escala.EOF Then
        v_factor = rs_escala!escporexe
        v_rebaja = rs_escala!esccuota
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "No se encontró escala para Año " & Anio & " y mes " & Mes
        End If
        
        'Busco para el año
        StrSql = "SELECT * FROM escala "
        StrSql = StrSql & " WHERE escano = " & Anio
        StrSql = StrSql & " AND escinf <= " & Monto & " AND escsup >= " & Monto
        StrSql = StrSql & " Order BY escano, escporexe"
        OpenRecordset StrSql, rs_escala
        If Not rs_escala.EOF Then
            v_factor = rs_escala!escporexe
            v_rebaja = rs_escala!esccuota
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "No se encontró escala para Año " & Anio
            End If
            
            StrSql = "SELECT * FROM escala "
            StrSql = StrSql & " WHERE escinf <= " & Monto & " AND escsup >= " & Monto
            StrSql = StrSql & " Order BY escporexe"
            OpenRecordset StrSql, rs_escala
            If Not rs_escala.EOF Then
                v_factor = rs_escala!escporexe
                v_rebaja = rs_escala!esccuota
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 1) & "No se encontró escala "
                End If
            End If
        End If
    End If
    
    Alicuota = v_factor
    Parcela = v_rebaja
    
    If rs_escala.State = adStateOpen Then rs_escala.Close
    Set rs_escala = Nothing
End Sub



