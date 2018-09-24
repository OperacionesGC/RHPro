Attribute VB_Name = "MdlFormulasBrasil"
Option Explicit

' ---------------------------------------------------------
' Modulo de fórmulas conocidas para Brasil
' ---------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------
'Cálculo mensual
'a)   Acumulador: T-Rend.Trib. ( mensual)
'b)  Acumulador: T. Deducciones.
'c)  Acumulador Fliares: (importe de Deducción *Cantidad cargada em ADP)
'
'                   Base Liq Ret: (A-(B+C))
'                   ((Base liq de retención * alícuota de escala)- parcela a deducir de escala)
'
'Calculo 13o. Salário
'd)  Acumulador: T-Rend.Trib.
'e)  Acumulador: T. Deducciones.
'f)  Acumulador Fliares: (importe de Deducción *Cantidad cargada em ADP)

'                   Base Liq Ret: (d-(e+f))
'                   ((Base liq de retención * alícuota de escala)- parcela a deducir de escala)
'
'
'Cálculo Anual:
'g)  Acumulador: T-Rend.Trib. ( Anual)
'h)  Acumulador: T. Deducciones. (Anual)
'i)  Acumulador Fliares anual: (importe de Deducción *Cantidad cargada em ADP)*12.
'
'                   Base Liq Ret: (G-(H+I))
'                   ((Base liq de retención * alícuota de escala)- parcela a deducir de escala)
'
'------------------------------
'Importante: el desglose de cálculo mensual y anual de IRRF se deberá detallar en el informe de rendimientos / DIRF.
'-----------
'Adicionalmente se debe tener en cuenta que el cuándo el empleado posee un embargo de pensión alimenticia
'   sobre los ingresos netos mensuales (también existe sobre bruto, pero se realiza de la manera ordinaria de rhpro), el cálculo se realiza en etapas.
'Primeramente se calcula el porcentaje y base a deducir que aplicaría tomando la base de renta, la cual se utiliza para el cálculo de la pensión alimenticia.
'
'Luego se efectúa el cálculo de IRRF normalmente teniendo en cuenta como retención de la base deducciones la pensión calculada.
'
'Y por último, un recalculo de IRRF únicamente si el rendimiento bruto (RB) menos (valor dependientes -/Contribuição Previdenciária /pensão alimentícia)
'   resulte inferior al tenido en cuenta en el primer paso.
'
'A continuación se detalla todo el cálculo mencionado:
'
'           P = {RB - CP - [(T/100) * (RB - CP - D - P)] + PD} * (PA / 100)
'
'Donde:
'   P = valor da pensão a ser paga;
'   RB = rendimento bruto;
'   CP = Contribuição Previdenciária;
'   T = alíquota da faixa da tabela progressiva a que pertencer o rendimento bruto;
'   D = dedução de dependentes, caso o contribuinte tenha outros dependentes sob sua guarda, que não o beneficiário da pensão;
'   PD = parcela a deduzir correspondente à faixa da base de cálculo (da tabela progressiva) a que pertencer o rendimento bruto;
'   PA = percentagem da pensão alimentícia fixada em juizo.
'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------

Public Function for_IRRF(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias para Brasil.
' Autor      : FGZ
' Fecha      : 18/09/2014
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim I As Long
Dim Ret_Mes As Integer
Dim Ret_Ano As Long


Dim c_Acunro_Mensual    As Integer  'Acumulador
Dim c_Pension_Bruto     As Integer  'Pension es sobre bruto o liquido
Dim c_CP                As Integer  'Acumulador
Dim c_D                 As Integer  'Acumulador
'Dim c_Escala            As Integer  'Escala IRRF

Dim v_Pension_Bruto     As Boolean
'Dim v_Escala           As Long
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

for_IRRF = P
exito = Bien

'Cierro y libero todo
'If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
'Set rs_wf_tpa = Nothing

End Function




Public Function BuscarAcumulador(ByVal BrutoMensual As Long, ByVal Anio As Long, ByVal Mes As Integer) As Double
Dim rs_Acum As New ADODB.Recordset
Dim Acum_Maximo As Double

    StrSql = "SELECT acu_mes.ammonto monto FROM sim_acu_mes "
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


Public Function PensionAlimenticia()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la pension alimenticia .
' Autor      : FGZ
' Fecha      : 29/11/2014
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_Pension As New ADODB.Recordset

Dim Aux_Pension As Double
    
    Aux_Pension = 0
    
'    StrSql = "SELECT * FROM embargo "
'    StrSql = StrSql & " WHERE escano = " & Anio
'    StrSql = StrSql & " AND escmes = " & Mes
'    StrSql = StrSql & " AND escinf <= " & Monto & " AND escsup >= " & Monto
'    StrSql = StrSql & " Order BY escano, escmes desc, escporexe"
'    OpenRecordset StrSql, rs_Pension
'    If Not rs_Pension.EOF Then
'        v_factor = rs_Pension!Campo
'    Else
'        If CBool(USA_DEBUG) Then
'            Flog.writeline Espacios(Tabulador * 1) & "No se encontró "
'        End If
'    End If
    
    
'        StrSql = "SELECT embargo.embnro,embargo.tpenro,embargo.embimp,embargo.retley,embargo.embaniofin,embargo.embmesfin, tipoemb.tpeprioridad, tipoemb.tpefordesc, tipoemb.tpehabsig, tipoemb.tpeton, tipoemb.tpecuosoc, embargo.monnro FROM embargo,tipoemb "
'        StrSql = StrSql & " WHERE embargo.ternro = " & buliq_empleado!Ternro
'        StrSql = StrSql & " AND  embargo.embest = 'A' "
'        StrSql = StrSql & " AND embargo.tpenro = tipoemb.tpenro"
'        StrSql = StrSql & " ORDER BY tipoemb.tpeprioridad "
'        OpenRecordset StrSql, rs_Embargo
'        If rs_Embargo.EOF Then
'            If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 4) & "No se encontraron embargos "
'        Else
'            'Busco la moneda origen del sistema
'            If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 4) & "Buscando Moneda del Sistema."
'        End If
    
    PensionAlimenticia = Aux_Pension
    
    If rs_Pension.State = adStateOpen Then rs_Pension.Close
    Set rs_Pension = Nothing
End Function

