Attribute VB_Name = "MdlFormulasUruguay"
Option Explicit

' ---------------------------------------------------------
' Modulo de fórmulas conocidas para Uruguay
' ---------------------------------------------------------

Public Function for_irp(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias para uruguay.
' Autor      :
' Fecha      :
' Ultima Mod.: FGZ - 11/04/2005
' Traducccion: FGZ - 13/09/2004
' Ultima Mod.: FGZ - 08/02/2007
' ---------------------------------------------------------------------------------------------
Dim c_mon_dir As Integer
Dim c_mon_mes As Integer
Dim c_porcentaje As Integer
Dim c_Acunro_Mensual As Integer
Dim c_acunro_directo As Integer

Dim v_mon_dir As Double
Dim v_mon_mes As Double
Dim v_porcentaje As Double
Dim v_Acunro_Mensual As Long
Dim v_acunro_Directo As Long

Dim imp_original As Double
Dim nivel_original As Integer
Dim neto_original As Double
Dim porc_original As Double
Dim imp_original_dir As Double
Dim nivel_original_dir As Integer
Dim neto_original_dir As Double
Dim porc_original_dir As Double
Dim porc_secund As Double
Dim porc_secund_dir As Double
Dim descuento_mes As Double
Dim descuento_fin As Double
Dim descuento_dir As Double
Dim neto_secund As Double
Dim neto_secund_dir As Double
Dim porcentaje_mes As Double
Dim porcentaje_fin As Double
Dim porcentaje_dir As Double

Dim imponible_mensual As Double
Dim irp_ya_dto As Double
Dim p_irp_ya_dto As Double

Dim monto_ya_des_directo As Double  'DEC INITIAL 0.
Dim cant_ya_des_directo As Double 'DEC INITIAL 0.
Dim monto_ya_des_mensual As Double 'DEC INITIAL 0.
Dim cant_ya_des_mensual As Double 'DEC INITIAL 0.

Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_uru_irpcab As New ADODB.Recordset
Dim rs_uru_irpdet As New ADODB.Recordset
Dim rs_ya_desc As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim rs_Acu_Liq2 As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset

'Inicializo
Bien = False

c_mon_dir = 1021
c_mon_mes = 1020
c_porcentaje = 35
c_Acunro_Mensual = 148
c_acunro_directo = 149

v_Acunro_Mensual = 0
v_acunro_Directo = 0
imp_original = 0
nivel_original = 0
neto_original = 0
porc_original = 0
imp_original_dir = 0
nivel_original_dir = 0
neto_original_dir = 0
porc_original_dir = 0
porc_secund = 0
porc_secund_dir = 0
descuento_mes = 0
descuento_fin = 0
descuento_dir = 0
neto_secund = 0
neto_secund_dir = 0
porcentaje_mes = 0
porcentaje_fin = 0
porcentaje_dir = 0

imponible_mensual = 0
irp_ya_dto = 0
p_irp_ya_dto = 0

monto_ya_des_directo = 0
cant_ya_des_directo = 0
monto_ya_des_mensual = 0
cant_ya_des_mensual = 0

' Primero limpio la traza
StrSql = "DELETE FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords


If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).ConcNro)
End If

'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case c_mon_dir:
        v_mon_dir = rs_wf_tpa!Valor
    Case c_mon_mes:
        v_mon_mes = rs_wf_tpa!Valor
    Case c_Acunro_Mensual:
        v_Acunro_Mensual = rs_wf_tpa!Valor
    Case c_acunro_directo:
        v_acunro_Directo = rs_wf_tpa!Valor
    Case c_porcentaje:
        v_porcentaje = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Monto Mensual " & v_mon_mes
    Flog.writeline Espacios(Tabulador * 1) & "Monto Directo " & v_mon_dir
    Flog.writeline Espacios(Tabulador * 1) & "Acumulador mensual " & v_Acunro_Mensual
    Flog.writeline Espacios(Tabulador * 1) & "Acumulador Directo " & v_acunro_Directo
    Flog.writeline Espacios(Tabulador * 1) & "Porcentaje " & v_porcentaje
End If

'Si el imponible del impuesto no es positivo, entonces dejo la formula
If v_acunro_Directo = 0 And v_acunro_Directo = 0 Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "Acumulador Mensual " & c_Acunro_Mensual
        Flog.writeline Espacios(Tabulador * 1) & "Acumulador " & c_acunro_directo
    End If
End If

'Si los acumuladores mensuales y directo no estan configurados no va a descontar el impuesto ya descontado de forma directa
If v_mon_dir <= 0 And v_mon_mes <= 0 Then
    Exit Function
End If

'Busco la cabecera del impuesto
StrSql = "SELECT * FROM uru_irpcab "
StrSql = StrSql & " WHERE irpcabfechasta >= " & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND irpcabfecdesde <= " & ConvFecha(buliq_proceso!profecpago)
OpenRecordset StrSql, rs_uru_irpcab

If rs_uru_irpcab.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe escala para la fecha de pago"
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 1, "No existe escala para la fecha de pago", 0)
    End If
    Exit Function
End If
If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Nro. escala encontrada " & rs_uru_irpcab!irpcabnro
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 1, "Nro. escala encontrada ", rs_uru_irpcab!irpcabnro)
End If
'Fin cabecera del impuesto


If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "busco lo ya descontado en forma directa"
End If
'busco lo ya descontado en forma directa
'Impuesto Ya descontado
StrSql = "SELECT * FROM sim_proceso "
StrSql = StrSql & " INNER JOIN sim_cabliq ON sim_proceso.pronro = sim_cabliq.pronro AND proceso.pronro <> " & buliq_proceso!pronro
StrSql = StrSql & " WHERE sim_proceso.pliqnro = " & buliq_proceso!PliqNro
StrSql = StrSql & " AND sim_cabliq.empleado = " & buliq_cabliq!Empleado
OpenRecordset StrSql, rs_ya_desc

If rs_ya_desc.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 2) & "No hay otros procesos liuqidados en el periodo"
    End If
End If
Do While Not rs_ya_desc.EOF
    monto_ya_des_directo = 0
    cant_ya_des_directo = 0
    monto_ya_des_mensual = 0
    cant_ya_des_mensual = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 2) & "BUSCO EL ACUMULADOR DE MENSUAL PARA SABER SI SON LIQUIDACIONES MENSUALES, SI SON DIRECTAS, NO HAY QUE TOMAR NADA"
    End If
    'BUSCO EL ACUMULADOR DE MENSUAL PARA SABER SI SON LIQUIDACIONES MENSUALES, SI SON DIRECTAS, NO HAY QUE TOMAR NADA
    StrSql = "SELECT * FROM sim_acu_liq "
    'FGZ - 09/02/2007 ---------------
    'StrSql = StrSql & " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro
    StrSql = StrSql & " WHERE sim_acu_liq.cliqnro = " & rs_ya_desc!cliqnro
    'FGZ - 09/02/2007 ---------------
    StrSql = StrSql & " AND sim_acu_liq.acunro = " & v_Acunro_Mensual
    OpenRecordset StrSql, rs_Acu_Liq
    If Not rs_Acu_Liq.EOF Then
        'buscar la proporcion de IMPONIBLE DIRECTO E IMPONIBLE MENSUAL, PARA DESCONTAR SOLO LA PROPORCION DEL MENSUAL
        StrSql = "SELECT * FROM sim_detliq "
        'FGZ - 08/02/2007 ---------------
        'StrSql = StrSql & " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro
        StrSql = StrSql & " WHERE sim_detliq.cliqnro = " & rs_Acu_Liq!cliqnro
        'FGZ - 08/02/2007 ---------------
        StrSql = StrSql & " AND sim_detliq.concnro = " & Buliq_Concepto(Concepto_Actual).ConcNro
        OpenRecordset StrSql, rs_Detliq
        If Not rs_Detliq.EOF Then
            irp_ya_dto = irp_ya_dto + rs_Detliq!dlimonto 'negativo
            p_irp_ya_dto = p_irp_ya_dto + rs_Detliq!dlicant 'positivo
        End If
        monto_ya_des_mensual = rs_Acu_Liq!almonto
        cant_ya_des_mensual = rs_Acu_Liq!alcant
        
        'VERIFICO SI HABIA ALGO DESCONTADO, SI HABIA ALGO, BUSCAMOS SI HAY DIRECTO, ACU=20, PARA PROPORCIONAR
        StrSql = "SELECT * FROM sim_acu_liq "
        'FGZ - 08/02/2007 ---------------
        'StrSql = StrSql & " WHERE acu_liq.cliqnro = " & buliq_cabliq!cliqnro
        StrSql = StrSql & " WHERE sim_acu_liq.cliqnro = " & rs_Acu_Liq!cliqnro
        'FGZ - 08/02/2007 ---------------
        StrSql = StrSql & " AND sim_acu_liq.acunro = " & v_acunro_Directo
        OpenRecordset StrSql, rs_Acu_Liq2
        
        If Not rs_Acu_Liq2.EOF Then
            monto_ya_des_directo = rs_Acu_Liq2!almonto
            cant_ya_des_directo = rs_Acu_Liq2!alcant
            'proporcion de regla de 3 directo para obtener el % del descuento que corresponde al directo
            irp_ya_dto = irp_ya_dto + (monto_ya_des_mensual / (monto_ya_des_directo + monto_ya_des_mensual))
            p_irp_ya_dto = p_irp_ya_dto + (monto_ya_des_mensual / (monto_ya_des_directo + monto_ya_des_mensual))

        End If  'existia acumulador mensual DE DIRECTO
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 2) & "No hay otros procesos liuqidados en el periodo"
        End If
    End If 'existia acumulador mensual DE MENSUAL
    
    rs_ya_desc.MoveNext
Loop

'Busco detalle del impuesto: Rangos
If v_mon_mes > 0 Then
    StrSql = "SELECT * FROM uru_irpdet "
    StrSql = StrSql & " WHERE irpcabnro = " & rs_uru_irpcab!irpcabnro
    StrSql = StrSql & " AND irpdetvaldesde < " & v_mon_mes
    StrSql = StrSql & " AND irpdetvalhasta >= " & v_mon_mes
    OpenRecordset StrSql, rs_uru_irpdet
    If rs_uru_irpdet.EOF Then
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 2, "No existe Franja escala monto", v_mon_mes)
            End If
            Exit Function
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 2, "Franja encontrada ", rs_uru_irpdet!irpdetnivel)
    End If

    imp_original = v_mon_mes * rs_uru_irpdet!irpdetporc / 100
    nivel_original = rs_uru_irpdet!irpdetnivel
    neto_original = v_mon_mes * (100 - rs_uru_irpdet!irpdetporc) / 100
    porc_original = rs_uru_irpdet!irpdetporc

    descuento_mes = -imp_original
    If rs_uru_irpdet.State = adStateOpen Then rs_uru_irpdet.Close

    'Busco la franja anterior de la misma escala
    If nivel_original <> 1 Then 'Cuando no es la primer franja de la escala, busca la franja anterior
        StrSql = "SELECT * FROM uru_irpdet "
        StrSql = StrSql & " WHERE irpcabnro = " & rs_uru_irpcab!irpcabnro
        StrSql = StrSql & " AND irpdetnivel = " & (nivel_original - 1)
        OpenRecordset StrSql, rs_uru_irpdet
        
        If Not rs_uru_irpdet.EOF Then
            neto_secund = rs_uru_irpdet!irpdetvalhasta * (100 - rs_uru_irpdet!irpdetporc) / 100
            porc_secund = rs_uru_irpdet!irpdetporc
    
            If neto_secund > neto_original Then
                descuento_mes = -v_mon_mes * (100 - (neto_secund * 100 / v_mon_mes)) / 100
            End If
       End If
    End If
    porcentaje_mes = -(descuento_mes / v_mon_mes * 100)
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 3, "Imponible entrante Mens.", v_mon_mes)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 4, "Porc Original Mes", porc_original)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 5, "Neto Original Mes", neto_original)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 6, "Porc.F.anterior Mes", porc_secund)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 7, "Neto.F.anterior Mes", neto_secund)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 8, "Descuento Final Mes", descuento_mes)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 9, "Porc. Final Mes", porcentaje_mes)
    End If
    If rs_uru_irpdet.State = adStateOpen Then rs_uru_irpdet.Close
End If 'monto mensual


'Busco impuesto directo
If v_mon_dir > 0 Then
    StrSql = "SELECT * FROM uru_irpdet "
    StrSql = StrSql & " WHERE irpcabnro = " & rs_uru_irpcab!irpcabnro
    StrSql = StrSql & " AND irpdetvaldesde < " & v_mon_dir
    StrSql = StrSql & " AND irpdetvalhasta >= " & v_mon_dir
    OpenRecordset StrSql, rs_uru_irpdet
    If rs_uru_irpdet.EOF Then
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 2, "No existe Franja escala monto", v_mon_dir)
            End If
            Exit Function
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 2, "Franja encontrada ", rs_uru_irpdet!irpdetnivel)
    End If

    imp_original_dir = v_mon_dir * rs_uru_irpdet!irpdetporc / 100
    nivel_original_dir = rs_uru_irpdet!irpdetnivel
    neto_original_dir = v_mon_dir * (100 - rs_uru_irpdet!irpdetporc) / 100
    porc_original_dir = rs_uru_irpdet!irpdetporc

    descuento_dir = -imp_original_dir
    If rs_uru_irpdet.State = adStateOpen Then rs_uru_irpdet.Close

    'Busco la franja anterior de la misma escala
    If nivel_original_dir <> 1 Then 'Cuando no es la primer franja de la escala, busca la franja anterior
        StrSql = "SELECT * FROM uru_irpdet "
        StrSql = StrSql & " WHERE irpcabnro = " & rs_uru_irpcab!irpcabnro
        StrSql = StrSql & " AND irpdetnivel = " & (nivel_original_dir - 1)
        OpenRecordset StrSql, rs_uru_irpdet
        
        If Not rs_uru_irpdet.EOF Then
            neto_secund_dir = rs_uru_irpdet!irpdetvalhasta * (100 - rs_uru_irpdet!irpdetporc) / 100
            porc_secund_dir = rs_uru_irpdet!irpdetporc
    
            If neto_secund_dir > neto_original_dir Then
                descuento_dir = -v_mon_dir * (100 - (neto_secund_dir * 100 / v_mon_dir)) / 100
            End If
       End If
    End If
    porcentaje_dir = -(descuento_dir / v_mon_dir * 100)
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 3, "Imponible entrante Mens.", v_mon_mes)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 4, "Porc Original Mes", porc_original)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 5, "Neto Original Mes", neto_original)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 6, "Porc.F.anterior Mes", porc_secund)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 7, "Neto.F.anterior Mes", neto_secund)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 8, "Descuento Final Mes", descuento_mes)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 9, "Porc. Final Mes", porcentaje_mes)
    End If
    If rs_uru_irpdet.State = adStateOpen Then rs_uru_irpdet.Close
End If 'monto directo

descuento_fin = descuento_mes - irp_ya_dto + descuento_dir
porcentaje_fin = porcentaje_mes - p_irp_ya_dto + porcentaje_dir
'FGZ - 09/02/2007 -----
If porcentaje_fin = 0 Then
    porcentaje_fin = porcentaje_mes
End If
'FGZ - 09/02/2007 -----
Monto = descuento_fin
Bien = True

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 10, "Imponible entrante Dir.", v_mon_dir)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 11, "Porc Original Dir.", porc_original_dir)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 12, "Neto Original Dir.", neto_original_dir)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 13, "Porc.F.anterior Dir.", porc_secund_dir)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 14, "Neto.F.anterior Dir.", neto_secund_dir)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 15, "Descuento Final Dir", descuento_dir)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 16, "Porc. Final Dir.", porcentaje_dir)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 17, "IRP Ya ret.", irp_ya_dto)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 18, "PORC.IRP Ya ret.", p_irp_ya_dto)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 19, "DESC. FINAL", descuento_fin)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 20, "PORC. FINAL", porcentaje_fin)
End If


'actualzo el parametro
StrSql = "UPDATE " & TTempWF_tpa & " SET valor = " & porcentaje_fin
StrSql = StrSql & " WHERE tipoparam =" & c_porcentaje
StrSql = StrSql & " AND fecha = " & ConvFecha(AFecha)
objConn.Execute StrSql, , adExecuteNoRecords
Parametro = porcentaje_fin

for_irp = Monto
exito = Bien

'Cierro y libero todo
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
If rs_uru_irpcab.State = adStateOpen Then rs_uru_irpcab.Close
If rs_uru_irpdet.State = adStateOpen Then rs_uru_irpdet.Close
If rs_ya_desc.State = adStateOpen Then rs_ya_desc.Close
If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close
If rs_Acu_Liq2.State = adStateOpen Then rs_Acu_Liq2.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close

Set rs_wf_tpa = Nothing
Set rs_uru_irpcab = Nothing
Set rs_uru_irpdet = Nothing
Set rs_ya_desc = Nothing
Set rs_Acu_Liq = Nothing
Set rs_Acu_Liq2 = Nothing
Set rs_Detliq = Nothing
End Function

Public Function for_irp_franja(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: calcula la franja de IRP en la que se cae(Uruguay).
' Autor      : DCH
' Fecha      : 06/08/2003
' Ultima Mod.:
' Traducccion: FGZ - 11/04/2005
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim c_msr As Integer    'Porcentaje calculado en el punto anterior
Dim c_franja As Integer 'Nro. de franja contra la que se quiere comparar
Dim c_monto As Integer  'Monto calculado del impuesto previamente

Dim v_msr As Double
Dim v_franja As Double
Dim v_monto As Double

Dim Porcentaje As Double

Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_uru_irpcab As New ADODB.Recordset
Dim rs_uru_irpdet As New ADODB.Recordset

'Inicializo
c_msr = 8
c_franja = 160
c_monto = 51

v_msr = -1
v_franja = -1
v_monto = -1

Bien = False

' Primero limpio la traza
StrSql = "DELETE FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).ConcNro)
End If


'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case c_msr:
        v_msr = rs_wf_tpa!Valor
    Case c_franja:
        v_franja = rs_wf_tpa!Valor
    Case c_monto:
        v_monto = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Porcentaje " & v_msr
    Flog.writeline Espacios(Tabulador * 1) & "Franja " & v_franja
    Flog.writeline Espacios(Tabulador * 1) & "Monto " & v_monto
End If

'Si el imponible del impuesto no es positivo, entonces dejo la formula
If v_msr < 0 Or v_monto < 0 Then
    Exit Function
End If

'Busco la cabecera del impuesto
StrSql = "SELECT * FROM uru_irpcab "
StrSql = StrSql & " WHERE irpcabfechasta >= " & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND irpcabfecdesde <= " & ConvFecha(buliq_proceso!profecpago)
OpenRecordset StrSql, rs_uru_irpcab

If rs_uru_irpcab.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe escala para la fecha de pago"
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 1, "No existe escala para la fecha de pago", 0)
    End If
    Exit Function
End If
If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Nro. escala encontrada " & rs_uru_irpcab!irpcabnro
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 1, "Nro. escala encontrada ", rs_uru_irpcab!irpcabnro)
End If
'Fin cabecera del impuesto

'Busco detalle del impuesto: Rangos
StrSql = "SELECT * FROM uru_irpdet "
StrSql = StrSql & " WHERE irpcabnro = " & rs_uru_irpcab!irpcabnro
StrSql = StrSql & " AND irpdetporc >= " & v_msr
StrSql = StrSql & " AND irpdetnivel = " & CLng(v_franja)
'StrSql = StrSql & " AND irpdetvaldesde < " & v_msr
'StrSql = StrSql & " AND irpdetvalhasta >= " & v_msr
OpenRecordset StrSql, rs_uru_irpdet

If rs_uru_irpdet.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe Franja escala monto " & v_msr
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 2, "No existe Franja escala monto", v_msr)
    End If
    Exit Function
Else
    Monto = v_monto
    Parametro = rs_uru_irpdet!irpdetporc
    Bien = True
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "Franja encontrada " & rs_uru_irpdet!irpdetnivel
        Flog.writeline Espacios(Tabulador * 1) & "Porcentaje " & Parametro
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 2, "Franja encontrada ", rs_uru_irpdet!irpdetnivel)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 2, "Porcentaje ", Parametro)
    End If
End If

for_irp_franja = Monto
exito = Bien

'Cierro y libero todo
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
If rs_uru_irpcab.State = adStateOpen Then rs_uru_irpcab.Close
If rs_uru_irpdet.State = adStateOpen Then rs_uru_irpdet.Close

Set rs_wf_tpa = Nothing
Set rs_uru_irpcab = Nothing
Set rs_uru_irpdet = Nothing
End Function


Public Function for_Ribeteado() As Double
' ---------------------------------------------------------------------------------------------
' Descripcion:  Formula para el calculo Incentivo Ribeteado
'               Reccorre la produccion dia por dia y compara con el indice minimo diario por las horas trabajadas,
'               y multiplica la cantidad de acuerdo a las franjas
' Autor      : CAT
' Fecha      : 11/02/2006
' Ultima Mod.: FGZ - 14/02/2006
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim c_indice_dia_1 As Integer
Dim c_indice_dia_2 As Integer
Dim c_indice_dia_3 As Integer
Dim c_indice_dia_4 As Integer
Dim c_indice_dia_5 As Integer
Dim c_indice_dia_6 As Integer
Dim c_indice_dia_7 As Integer
Dim c_indice_dia_8 As Integer
Dim c_indice_dia_9 As Integer
Dim c_indice_dia_10 As Integer
Dim c_indice_dia_11 As Integer
Dim c_indice_dia_12 As Integer
Dim c_indice_dia_13 As Integer
Dim c_indice_dia_14 As Integer
Dim c_indice_dia_15 As Integer
Dim c_indice_dia_16 As Integer
Dim c_indice_dia_17 As Integer
Dim c_indice_dia_18 As Integer
Dim c_indice_dia_19 As Integer
Dim c_indice_dia_20 As Integer
Dim c_indice_dia_21 As Integer
Dim c_indice_dia_22 As Integer
Dim c_indice_dia_23 As Integer
Dim c_indice_dia_24 As Integer
Dim c_indice_dia_25 As Integer
Dim c_indice_dia_26 As Integer
Dim c_indice_dia_27 As Integer
Dim c_indice_dia_28 As Integer
Dim c_indice_dia_29 As Integer
Dim c_indice_dia_30 As Integer
Dim c_indice_dia_31 As Integer

Dim c_tipohora As Integer 'tipo de hora que contiene las horas diarias
Dim c_coeficiente As Integer ' colchones por hora
Dim c_franja1 As Integer 'segundo ragno rango
Dim c_franja2 As Integer 'segundo ragno rango
Dim c_valorfranja1 As Integer 'valor primer rango
Dim c_valorfranja2 As Integer 'valor segundo ragno rango
Dim c_valorresto As Integer 'Valor Resto

Dim v_indice_dia(31) As Double

Dim v_tipohora      As Double 'tipo de hora que contiene las horas diarias
Dim v_coeficiente   As Double ' colchones por hora
Dim v_franja1       As Double 'Primer rango
Dim v_franja2       As Double 'segundo rango
Dim v_valorfranja1  As Double 'valor primer rango
Dim v_valorfranja2  As Double 'valor segundo ragno rango
Dim v_valorresto    As Double 'Valor Resto

Dim Encontro1 As Boolean
Dim Encontro2 As Boolean

Dim fecha_desde  As Date
Dim fecha_hasta  As Date
Dim descripcion As String
Dim productividad_dia As Double
Dim productividad   As Double
Dim cant_bultos  As Double
Dim indice_diario As Double
Dim cant_total_bultos  As Double
Dim cant_embaladores  As Double
Dim total_colchones_dia As Double
Dim cant_horas_dia As Double

Dim I As Integer

Dim rs_wf_tpa As New ADODB.Recordset
Dim buf_gti_adiario As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_gti_adiario As New ADODB.Recordset

'inicializacion de variables
'c_indice = 166
'c_monto = 51

c_indice_dia_1 = 1101
c_indice_dia_2 = 1102
c_indice_dia_3 = 1103
c_indice_dia_4 = 1104
c_indice_dia_5 = 1105
c_indice_dia_6 = 1106
c_indice_dia_7 = 1107
c_indice_dia_8 = 1108
c_indice_dia_9 = 1109
c_indice_dia_10 = 1110
c_indice_dia_11 = 1111
c_indice_dia_12 = 1112
c_indice_dia_13 = 1113
c_indice_dia_14 = 1114
c_indice_dia_15 = 1115
c_indice_dia_16 = 1116
c_indice_dia_17 = 1117
c_indice_dia_18 = 1118
c_indice_dia_19 = 1119
c_indice_dia_20 = 1120
c_indice_dia_21 = 1121
c_indice_dia_22 = 1122
c_indice_dia_23 = 1123
c_indice_dia_24 = 1124
c_indice_dia_25 = 1125
c_indice_dia_26 = 1126
c_indice_dia_27 = 1127
c_indice_dia_28 = 1128
c_indice_dia_29 = 1129
c_indice_dia_30 = 1130
c_indice_dia_31 = 1131

c_tipohora = 1132
c_coeficiente = 1133
c_franja1 = 1134
c_franja2 = 1135
c_valorfranja1 = 1136
c_valorfranja2 = 1137
c_valorresto = 1138

For I = 1 To 31
    v_indice_dia(I) = 0
Next I

v_tipohora = 0
v_coeficiente = 0
v_franja1 = 0
v_franja2 = 0
v_valorfranja1 = 0
v_valorfranja2 = 0
v_valorresto = 0


    Bien = False
    exito = False
    Encontro1 = False
    Encontro2 = False

    'Leo los parametros de la formula
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Parametros de la formula"
    End If

    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case c_valorresto:
            v_valorresto = rs_wf_tpa!Valor
        Case c_indice_dia_1:
            v_indice_dia(1) = rs_wf_tpa!Valor
        Case c_indice_dia_2:
            v_indice_dia(2) = rs_wf_tpa!Valor
        Case c_indice_dia_3:
            v_indice_dia(3) = rs_wf_tpa!Valor
        Case c_indice_dia_4:
            v_indice_dia(4) = rs_wf_tpa!Valor
        Case c_indice_dia_5:
            v_indice_dia(5) = rs_wf_tpa!Valor
        Case c_indice_dia_6:
            v_indice_dia(6) = rs_wf_tpa!Valor
        Case c_indice_dia_7:
            v_indice_dia(7) = rs_wf_tpa!Valor
        Case c_indice_dia_8:
            v_indice_dia(8) = rs_wf_tpa!Valor
        Case c_indice_dia_9:
            v_indice_dia(9) = rs_wf_tpa!Valor
        Case c_indice_dia_10:
            v_indice_dia(10) = rs_wf_tpa!Valor
        Case c_indice_dia_11:
            v_indice_dia(11) = rs_wf_tpa!Valor
        Case c_indice_dia_12:
            v_indice_dia(12) = rs_wf_tpa!Valor
        Case c_indice_dia_13:
            v_indice_dia(13) = rs_wf_tpa!Valor
        Case c_indice_dia_14:
            v_indice_dia(14) = rs_wf_tpa!Valor
        Case c_indice_dia_15:
            v_indice_dia(15) = rs_wf_tpa!Valor
        Case c_indice_dia_16:
            v_indice_dia(16) = rs_wf_tpa!Valor
        Case c_indice_dia_17:
            v_indice_dia(17) = rs_wf_tpa!Valor
        Case c_indice_dia_18:
            v_indice_dia(18) = rs_wf_tpa!Valor
        Case c_indice_dia_19:
            v_indice_dia(19) = rs_wf_tpa!Valor
        Case c_indice_dia_20:
            v_indice_dia(20) = rs_wf_tpa!Valor
        Case c_indice_dia_21:
            v_indice_dia(21) = rs_wf_tpa!Valor
        Case c_indice_dia_22:
            v_indice_dia(22) = rs_wf_tpa!Valor
        Case c_indice_dia_23:
            v_indice_dia(23) = rs_wf_tpa!Valor
        Case c_indice_dia_24:
            v_indice_dia(24) = rs_wf_tpa!Valor
        Case c_indice_dia_25:
            v_indice_dia(25) = rs_wf_tpa!Valor
        Case c_indice_dia_26:
            v_indice_dia(26) = rs_wf_tpa!Valor
        Case c_indice_dia_27:
            v_indice_dia(27) = rs_wf_tpa!Valor
        Case c_indice_dia_28:
            v_indice_dia(28) = rs_wf_tpa!Valor
        Case c_indice_dia_29:
            v_indice_dia(29) = rs_wf_tpa!Valor
        Case c_indice_dia_30:
            v_indice_dia(30) = rs_wf_tpa!Valor
        Case c_indice_dia_31:
            v_indice_dia(31) = rs_wf_tpa!Valor
        
        Case c_tipohora:
            v_tipohora = rs_wf_tpa!Valor
        Case c_coeficiente:
            v_coeficiente = rs_wf_tpa!Valor
        Case c_franja2:
            v_franja2 = rs_wf_tpa!Valor
        Case c_valorfranja1:
            v_valorfranja1 = rs_wf_tpa!Valor
        Case c_valorfranja2:
            v_valorfranja2 = rs_wf_tpa!Valor
        
        End Select
        
        
        rs_wf_tpa.MoveNext
    Loop

    
'    ' si no se obtuvieron los parametros, ==> Error.
'    If Not Encontro1 Or Not Encontro2 Then
'        Exit Function
'    End If
    
    
'Recorre la Producci¢n Diaria

'Recorriendo el Desglose Diario
fecha_desde = buliq_proceso!profecini
fecha_hasta = buliq_proceso!profecfin
If Not CBool(buliq_empleado!empest) Then
    If fecha_hasta > Empleado_Fecha_Fin Then
        fecha_hasta = Empleado_Fecha_Fin
    End If
End If


'Obtengo las horas trabajadas por dia
If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "Obtengo las horas trabajadas por dia"
End If

StrSql = "SELECT * FROM gti_acumdiario "
StrSql = StrSql & " WHERE ( " & ConvFecha(fecha_desde) & " <= adfecha "
StrSql = StrSql & " AND adfecha <= " & ConvFecha(fecha_hasta) & ")"
StrSql = StrSql & " AND (gti_acumdiario.thnro  = " & v_tipohora & ")"
OpenRecordset StrSql, rs_gti_adiario


Do While Not rs_gti_adiario.EOF
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Dia: " & rs_gti_adiario!adfecha
    End If
    
    cant_horas_dia = rs_gti_adiario!adcanthoras
    v_franja1 = Round(cant_horas_dia * v_coeficiente, 0)
    productividad_dia = 0

    total_colchones_dia = v_indice_dia(Day(rs_gti_adiario!adfecha))
    'Si el total de colchones diarios supera a la cantidad de horas
    'por el coeficiente de colchones minimos por hora
    
    If total_colchones_dia >= v_franja1 Then
        'pago la primer franja
        productividad_dia = v_franja1 * v_valorfranja1
        'le resto para pagar la segunda franja
        total_colchones_dia = total_colchones_dia - v_franja1
        If total_colchones_dia > v_franja2 Then
            'pago la segunda franja
            productividad_dia = productividad_dia + v_franja2 * v_valorfranja2
            'le resto para pagar el resto
            total_colchones_dia = total_colchones_dia - v_franja2
            'pago el resto
            productividad_dia = productividad_dia + total_colchones_dia * v_valorresto
        Else
            productividad_dia = productividad_dia + total_colchones_dia * v_valorfranja2
        End If
    End If
    
    productividad = productividad + productividad_dia

    If HACE_TRAZA Then
        descripcion = Format(rs_gti_adiario!adfecha, "dd/mm/yy")
        descripcion = descripcion + " Ind: " + Format(v_indice_dia(Day(rs_gti_adiario!adfecha)), "0000.00") + " $"
        Call InsertarTraza(NroCab, Buliq_Concepto(Concepto_Actual).ConcNro, 0, descripcion, productividad_dia)
    End If
   
    rs_gti_adiario.MoveNext
Loop

Monto = productividad
for_Ribeteado = productividad
Bien = True
exito = True
End Function

Public Function for_irpf(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias para uruguay.
' Autor      : FGZ
' Fecha      : 05/06/2007
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Sum_Renta As Double
Dim Sum_Deduccion As Double
Dim BPC_Renta As Double
Dim BPC_Deduccion As Double
Dim Por_Renta As Double
Dim Por_Deduccion As Double
Dim Renta As Double
Dim Deduccion As Double
Dim Impuesto As Double
Dim Otras_Ret As Double
'FGZ - 12/05/2014 --------------
'Dim Items_LIQ(50) As Double
'Dim Items_DDJJ(50) As Double
Dim Items_LIQ(100) As Double
Dim Items_DDJJ(100) As Double

'Parametros
Dim c_Aplica_Imp As Long
Dim c_Meses As Long
Dim c_Busq As Long
Dim v_Aplica_Imp As Long
Dim v_Meses As Long
Dim v_Busq As Long

'Auxiliares
Dim Ret_Mes As Integer
Dim Ret_Ano As Integer
Dim Con_liquid As Integer
Dim Acum As Long
Dim Aux_Acu_Monto As Double
Dim Cabecera_Renta As Long
Dim Cabecera_Deduccion As Long
Dim Anual As Boolean
Dim I As Long
Dim Hasta As Long


Dim rs_IrpfCab As New ADODB.Recordset
Dim rs_IrpfDet As New ADODB.Recordset
Dim rs_IrpfDedCab As New ADODB.Recordset
Dim rs_IrpfDedDet As New ADODB.Recordset
Dim rs_Ficharet As New ADODB.Recordset
Dim rs_Item As New ADODB.Recordset
Dim rs_Desmen As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Traza_gan_items_tope As New ADODB.Recordset
Dim rs_Acum As New ADODB.Recordset
Dim Cantidad_BPC_Deducciones As Double


'Inicializo
Bien = False
Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro

'defaults
c_Aplica_Imp = 140
c_Meses = 150
c_Busq = 52
v_Aplica_Imp = 0
v_Meses = 12
v_Busq = 0

'Primero limpio la traza
StrSql = "DELETE FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords
If HACE_TRAZA Then
    'Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).ConcNro)
    Call LimpiarTrazaConcepto(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro)
End If

'Guardo la traza de Ganancia en traza_gan para utilizarla luego cuando se genere el reporte
StrSql = "INSERT INTO sim_traza_gan (pliqnro,pronro,concnro,fecha_pago,empresa,ternro,empleg) VALUES ("
StrSql = StrSql & buliq_periodo!PliqNro & ","
StrSql = StrSql & buliq_proceso!pronro & ","
StrSql = StrSql & Buliq_Concepto(Concepto_Actual).ConcNro & ","
StrSql = StrSql & ConvFecha(buliq_proceso!profecpago) & ","
StrSql = StrSql & NroEmp & ","
StrSql = StrSql & buliq_empleado!Ternro & ","
StrSql = StrSql & buliq_empleado!Empleg
StrSql = StrSql & ")"
objConn.Execute StrSql, , adExecuteNoRecords


'Obtencion de los parametros de WorkFile
'FGZ - 24/05/2011 ------------------------------------
'rs_wf_tpa!valor por     Arr_WF_TPA(I).valor
For I = LI_WF_Tpa To LS_WF_Tpa
    Select Case Arr_WF_TPA(I).tipoparam
    Case c_Aplica_Imp:
        v_Aplica_Imp = Arr_WF_TPA(I).Valor
    Case c_Meses:
        v_Meses = Arr_WF_TPA(I).Valor
    Case c_Busq:
       v_Busq = Arr_WF_TPA(I).Valor
    Case Else
    End Select
Next I


If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Aplica Imponible " & CBool(v_Aplica_Imp)
    Flog.writeline Espacios(Tabulador * 1) & "Meses " & v_Meses
End If
'Genero la traza de los familiares deducidos para IRPF
If HACE_TRAZA Then
    Call Traza_DeduccionFliaresIRPF(v_Busq)
End If
'**************************************************************************

If CBool(v_Aplica_Imp) Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "No se debe Liquidar el concepto."
    End If
    Bien = True
    exito = Bien
    Exit Function
End If

'Busco el periodo que esoy liquidando para ver si tiene la marca de ultimo periodo
Anual = CBool(buliq_periodo!pliqultimo)

If Not Anual Then
    ' Recorro todos los items de Ganancias
    StrSql = "SELECT * FROM item ORDER BY itetipotope"
    OpenRecordset StrSql, rs_Item
    Do While Not rs_Item.EOF
            'Tomo los valores de DDJJ y Liquidacion sin Tope
            'Busco la declaracion jurada
            StrSql = "SELECT * FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                     " AND desano=" & Ret_Ano & _
                     " AND itenro = " & rs_Item!Itenro
            OpenRecordset StrSql, rs_Desmen
            Do While Not rs_Desmen.EOF
                If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                End If
                rs_Desmen.MoveNext
            Loop
            
            StrSql = "SELECT * FROM itemacum " & _
                     " WHERE itenro =" & rs_Item!Itenro & _
                     " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
            OpenRecordset StrSql, rs_itemacum
            Do While Not rs_itemacum.EOF
                Acum = CStr(rs_itemacum!acuNro)
                Aux_Acu_Monto = 0
                'Liquidacion actual
                If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                    Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
                End If
                
                'Mas el acumulador mensual
                StrSql = "SELECT * FROM sim_acu_mes "
                StrSql = StrSql & " WHERE ternro = " & buliq_empleado!Ternro
                StrSql = StrSql & " AND acunro = " & Acum
                StrSql = StrSql & " AND amanio = " & Ret_Ano
                StrSql = StrSql & " AND ammes = " & Ret_Mes
                OpenRecordset StrSql, rs_Acum
                If Not rs_Acum.EOF Then
                    Aux_Acu_Monto = Aux_Acu_Monto + rs_Acum!ammonto
                End If
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
                End If

                rs_itemacum.MoveNext
            Loop


            'Armo la traza del item
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Item " & rs_Item!Itenro
                Flog.writeline Espacios(Tabulador * 4) & "LIQ" & Items_LIQ(rs_Item!Itenro)
                Flog.writeline Espacios(Tabulador * 4) & "DDJJ" & Items_DDJJ(rs_Item!Itenro)
            End If
            If HACE_TRAZA Then
                Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
                Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_LIQ(rs_Item!Itenro))
            End If
            
            'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
            ' "ACHIQUE" DE GANANCIA IMPONIBLE
            If CBool(rs_Item!itesigno) Then
                Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro)
                Sum_Renta = Sum_Renta + Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            Else
                Items_LIQ(rs_Item!Itenro) = -Items_LIQ(rs_Item!Itenro)
                Sum_Deduccion = Sum_Deduccion + Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            End If
                
        rs_Item.MoveNext
    Loop
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Renta        : " & Sum_Renta
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Deducciones  : " & Sum_Deduccion
        Flog.writeline
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Renta", Sum_Renta)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Deducciones", Sum_Deduccion)
    End If
    
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " nogan =" & Sum_Renta
    StrSql = StrSql & " ,otras =" & Sum_Deduccion
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    '**************************************************************************
    
    'Busco en escala para Renta
    'cabecera
    StrSql = "SELECT * FROM uru_irpfcab "
    StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND fecdesde <= " & ConvFecha(buliq_proceso!profecpago)
    OpenRecordset StrSql, rs_IrpfCab
    If rs_IrpfCab.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Renta. No existe escala para la fecha de pago"
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - No existe escala para la fecha de pago", 0)
        End If
        Exit Function
    Else
        BPC_Renta = rs_IrpfCab!val
        Cabecera_Renta = rs_IrpfCab!cabnro
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "BPC para Renta " & BPC_Renta
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - BPC para Renta", rs_IrpfCab!val)
    End If
    'Detalle Renta
    StrSql = "SELECT * FROM uru_irpfdet "
    StrSql = StrSql & " WHERE cabnro = " & Cabecera_Renta
    StrSql = StrSql & " AND valdesde * " & BPC_Renta & " < " & Sum_Renta
    StrSql = StrSql & " ORDER BY valdesde"
    OpenRecordset StrSql, rs_IrpfDet
    If rs_IrpfDet.EOF Then
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - No Entra en Franja de Escala", Sum_Renta)
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "No Entra en Franja de Escala" & Sum_Renta
        End If
    End If
    Do While Not rs_IrpfDet.EOF
        If (rs_IrpfDet!valhasta * BPC_Renta) <= Sum_Renta Then
            Renta = Renta + ((rs_IrpfDet!valhasta * BPC_Renta) - (rs_IrpfDet!valdesde * BPC_Renta)) * rs_IrpfDet!Porc / 100
        Else
            Renta = Renta + ((Sum_Renta) - (rs_IrpfDet!valdesde * BPC_Renta)) * rs_IrpfDet!Porc / 100
        End If
    
        'FGZ - 25/01/2017 -------------------------------
        Cantidad_BPC_Deducciones = rs_IrpfDet!valdesde
        'FGZ - 25/01/2017 -------------------------------
    
        rs_IrpfDet.MoveNext
    Loop
    If HACE_TRAZA Then
        'Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "93 - Porcentaje Renta ", Por_Renta)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Renta", Renta)
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline
        'Flog.writeline Espacios(Tabulador * 3) & "Porcentaje Renta  :" & Por_Renta
        Flog.writeline Espacios(Tabulador * 3) & "Renta             :" & Renta
    End If
    
    
    
    'FGZ - 25/01/2017 ------------------------------------------------
    'PBI 6245 - ITRP - 4631111 - Uruguay -Cambio legal - IRPF 2017
    '
    'Busco en escala para Deduccion
    'cabecera
    StrSql = "SELECT * FROM uru_irpfdedcab "
    StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND fecdesde <= " & ConvFecha(buliq_proceso!profecpago)
    OpenRecordset StrSql, rs_IrpfDedCab
    If rs_IrpfDedCab.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Deducción. No existe escala para la fecha de pago"
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - Deducción. No existe escala para la fecha de pago", 0)
        End If
        Exit Function
    Else
        BPC_Deduccion = rs_IrpfDedCab!val
        Cabecera_Deduccion = rs_IrpfDedCab!cabnro
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "BPC para Deducción " & BPC_Deduccion
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - BPC para Deducción", BPC_Deduccion)
    End If
    
    'Detalle Renta
    StrSql = "SELECT * FROM uru_irpfdeddet "
    StrSql = StrSql & " WHERE cabnro = " & Cabecera_Deduccion
    StrSql = StrSql & " AND valhasta >= " & Cantidad_BPC_Deducciones
    StrSql = StrSql & " ORDER BY valdesde"
    OpenRecordset StrSql, rs_IrpfDedDet
    If rs_IrpfDedDet.EOF Then
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - No Entra en Franja de Escala", Sum_Deduccion)
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "No Entra en Franja de Escala" & Sum_Deduccion
        End If
    End If
    
    If Not rs_IrpfDedDet.EOF Then
        Deduccion = Deduccion + (Sum_Deduccion * rs_IrpfDedDet!Porc / 100)
        Por_Deduccion = rs_IrpfDedDet!Porc
    End If
    
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - Porcentaje Deducción ", Por_Deduccion)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Deducción", Deduccion)
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 3) & "Porcentaje Deducción  :" & Por_Deduccion
        Flog.writeline Espacios(Tabulador * 3) & "Deducción             :" & Deduccion
    End If
    '
    'PBI 6245 - ITRP - 4631111 - Uruguay -Cambio legal - IRPF 2017
    'FGZ - 25/01/2017 ------------------------------------------------

    
    
    
    
    'Busco otras retenciones efectuadas en el mes (en otros procesos)
    Otras_Ret = 0
        
    StrSql = "SELECT * FROM sim_ficharet "
    StrSql = StrSql & " WHERE empleado =" & buliq_empleado!Ternro
    'FGZ - 30/07/2012 --------------------------------------------
    'StrSql = StrSql & " AND month(fecha) = " & Ret_Mes
    'StrSql = StrSql & " AND year(fecha) = " & Ret_Ano
    StrSql = StrSql & " AND fecha >= " & ConvFecha(CDate("01/" & Ret_Mes & "/" & Ret_Ano)) & " AND fecha <= " & ConvFecha(UltimoDiaMes(Ret_Ano, Ret_Mes))
    'FGZ - 30/07/2012 --------------------------------------------
    StrSql = StrSql & " ORDER BY fecha desc"
    OpenRecordset StrSql, rs_Ficharet
    If Not rs_Ficharet.EOF Then
       Otras_Ret = rs_Ficharet!importe
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Otras Retenciones:" & Otras_Ret
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "95 - Otras Retenciones ", Otras_Ret)
    End If
    
    'Resultado del impuesto
    'Impuesto = (Renta - Deduccion - Otras_Ret) * -1
    'No se deben restar otras retenciones
    Impuesto = (Renta - Deduccion) * -1
    Call InsertarFichaRet(buliq_empleado!Ternro, buliq_proceso!profecpago, Abs(Impuesto), buliq_proceso!pronro)
Else

End If
StrSql = "UPDATE sim_traza_gan SET "
StrSql = StrSql & " ganneta =" & Impuesto
StrSql = StrSql & " ,ganimpo =" & Renta
StrSql = StrSql & " ,deducciones =" & Deduccion
StrSql = StrSql & " ,msr =" & BPC_Renta
StrSql = StrSql & " ,nomsr =" & BPC_Deduccion
StrSql = StrSql & " ,noimpo =" & Por_Renta
StrSql = StrSql & " ,porcdeduc =" & Por_Deduccion
StrSql = StrSql & " ,retenciones =" & Otras_Ret
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords


' Grabo todos los items de la liquidacion actual
I = 1
Hasta = 50
Do While I <= Hasta
    
    If Items_LIQ(I) <> 0 Then
       'Busco las liquidaciones anteriores
        StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                 buliq_empleado!Ternro & "," & _
                 ConvFecha(buliq_proceso!profecpago) & "," & _
                 buliq_proceso!pronro & "," & _
                 Items_LIQ(I) & "," & _
                 "0," & _
                 I & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    'guardo los item_ddjj para poder usarlo en el reporte de Ganancias
    If Items_DDJJ(I) <> 0 Then
        StrSql = "SELECT * FROM traza_gan_item_top "
        StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND itenro =" & I
        OpenRecordset StrSql, rs_Traza_gan_items_tope
        If rs_Traza_gan_items_tope.EOF Then
            StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,ddjj,empresa,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_DDJJ(I) & "," & _
                     NroEmp & "," & _
                     I & _
                     ")"
        Else 'Actualizo
            StrSql = "UPDATE traza_gan_item_top SET ddjj =" & Items_DDJJ(I)
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

    'guardo los item_liq para poder usarlo en el reporte de Ganancias
    If Items_LIQ(I) <> 0 Then
        StrSql = "SELECT * FROM traza_gan_item_top "
        StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND itenro =" & I
        OpenRecordset StrSql, rs_Traza_gan_items_tope
        If rs_Traza_gan_items_tope.EOF Then
            StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,liq,empresa,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_LIQ(I) & "," & _
                     NroEmp & "," & _
                     I & _
                     ")"
        Else 'Actualizo
            StrSql = "UPDATE traza_gan_item_top SET liq =" & Items_LIQ(I)
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    I = I + 1
Loop

Monto = Impuesto
Bien = True

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Impuesto: " & Impuesto
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "99 - Impuesto", Impuesto)
End If

for_irpf = Monto
exito = Bien

'Cierro y libero todo
    If rs_IrpfCab.State = adStateOpen Then rs_IrpfCab.Close
    If rs_IrpfDet.State = adStateOpen Then rs_IrpfDet.Close
    If rs_IrpfDedCab.State = adStateOpen Then rs_IrpfDedCab.Close
    If rs_IrpfDedDet.State = adStateOpen Then rs_IrpfDedDet.Close
    If rs_Ficharet.State = adStateOpen Then rs_Ficharet.Close
    If rs_Item.State = adStateOpen Then rs_Item.Close
    If rs_Desmen.State = adStateOpen Then rs_Desmen.Close
    If rs_itemacum.State = adStateOpen Then rs_itemacum.Close
    If rs_itemconc.State = adStateOpen Then rs_itemconc.Close
    If rs_Traza_gan_items_tope.State = adStateOpen Then rs_Traza_gan_items_tope.Close
    If rs_Acum.State = adStateOpen Then rs_Acum.Close
    
    Set rs_IrpfCab = Nothing
    Set rs_IrpfDet = Nothing
    Set rs_IrpfDedCab = Nothing
    Set rs_IrpfDedDet = Nothing
    Set rs_Ficharet = Nothing
    Set rs_Item = Nothing
    Set rs_Desmen = Nothing
    Set rs_itemacum = Nothing
    Set rs_itemconc = Nothing
    Set rs_Traza_gan_items_tope = Nothing
    Set rs_Acum = Nothing
End Function


Public Function for_irpf_OLD(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias para uruguay.
' Autor      : FGZ
' Fecha      : 05/06/2007
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Sum_Renta As Double
Dim Sum_Deduccion As Double
Dim BPC_Renta As Double
Dim BPC_Deduccion As Double
Dim Por_Renta As Double
Dim Por_Deduccion As Double
Dim Renta As Double
Dim Deduccion As Double
Dim Impuesto As Double
Dim Otras_Ret As Double
Dim Items_LIQ(100) As Double
Dim Items_DDJJ(100) As Double

'Parametros
Dim c_Aplica_Imp As Long
Dim c_Meses As Long
Dim c_Busq As Long
Dim v_Aplica_Imp As Long
Dim v_Meses As Long
Dim v_Busq As Long

'Auxiliares
Dim Ret_Mes As Integer
Dim Ret_Ano As Integer
Dim Con_liquid As Integer
Dim Acum As Long
Dim Aux_Acu_Monto As Double
Dim Cabecera_Renta As Long
Dim Cabecera_Deduccion As Long
Dim Anual As Boolean
Dim I As Long
Dim Hasta As Long


Dim rs_IrpfCab As New ADODB.Recordset
Dim rs_IrpfDet As New ADODB.Recordset
Dim rs_IrpfDedCab As New ADODB.Recordset
Dim rs_IrpfDedDet As New ADODB.Recordset
Dim rs_Ficharet As New ADODB.Recordset
Dim rs_Item As New ADODB.Recordset
Dim rs_Desmen As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Traza_gan_items_tope As New ADODB.Recordset
Dim rs_Acum As New ADODB.Recordset

'Inicializo
Bien = False
Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro

'defaults
c_Aplica_Imp = 140
c_Meses = 150
c_Busq = 52
v_Aplica_Imp = 0
v_Meses = 12
v_Busq = 0

'Primero limpio la traza
StrSql = "DELETE FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords
If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).ConcNro)
End If

'Guardo la traza de Ganancia en traza_gan para utilizarla luego cuando se genere el reporte
StrSql = "INSERT INTO sim_traza_gan (pliqnro,pronro,concnro,fecha_pago,empresa,ternro,empleg) VALUES ("
StrSql = StrSql & buliq_periodo!PliqNro & ","
StrSql = StrSql & buliq_proceso!pronro & ","
StrSql = StrSql & Buliq_Concepto(Concepto_Actual).ConcNro & ","
StrSql = StrSql & ConvFecha(buliq_proceso!profecpago) & ","
StrSql = StrSql & NroEmp & ","
StrSql = StrSql & buliq_empleado!Ternro & ","
StrSql = StrSql & buliq_empleado!Empleg
StrSql = StrSql & ")"
objConn.Execute StrSql, , adExecuteNoRecords


'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa
Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case c_Aplica_Imp:
        v_Aplica_Imp = rs_wf_tpa!Valor
    Case c_Meses:
        v_Meses = rs_wf_tpa!Valor
    Case c_Busq:
        v_Busq = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop
If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Aplica Imponible " & CBool(v_Aplica_Imp)
    Flog.writeline Espacios(Tabulador * 1) & "Meses " & v_Meses
End If
'Genero la traza de los familiares deducidos para IRPF
If HACE_TRAZA Then
    Call Traza_DeduccionFliaresIRPF(v_Busq)
End If
'**************************************************************************

If CBool(v_Aplica_Imp) Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "No se debe Liquidar el concepto."
    End If
    Bien = True
    exito = Bien
    Exit Function
End If

'Busco el periodo que esoy liquidando para ver si tiene la marca de ultimo periodo
Anual = CBool(buliq_periodo!pliqultimo)

If Not Anual Then
    ' Recorro todos los items de Ganancias
    StrSql = "SELECT * FROM item ORDER BY itetipotope"
    OpenRecordset StrSql, rs_Item
    Do While Not rs_Item.EOF
            'Tomo los valores de DDJJ y Liquidacion sin Tope
            'Busco la declaracion jurada
            StrSql = "SELECT * FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                     " AND desano=" & Ret_Ano & _
                     " AND itenro = " & rs_Item!Itenro
            OpenRecordset StrSql, rs_Desmen
            Do While Not rs_Desmen.EOF
                If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                End If
                rs_Desmen.MoveNext
            Loop
            
'FGZ - 14/06/2007 - en lugar de buscar acu_liq vamos a buscar acu_mes -------
            'Busco los acumuladores de la liquidacion
'            StrSql = "SELECT * FROM itemacum " & _
'                     " WHERE itenro =" & rs_Item!Itenro & _
'                     " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'            OpenRecordset StrSql, rs_itemacum
'            Do While Not rs_itemacum.EOF
'                Acum = CStr(rs_itemacum!acunro)
'                If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
'                    Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
'
'                    If CBool(rs_itemacum!itasigno) Then
'                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
'                    Else
'                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
'                    End If
'                End If
'                rs_itemacum.MoveNext
'            Loop
            
            StrSql = "SELECT * FROM itemacum " & _
                     " WHERE itenro =" & rs_Item!Itenro & _
                     " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
            OpenRecordset StrSql, rs_itemacum
            Do While Not rs_itemacum.EOF
                Acum = CStr(rs_itemacum!acuNro)
                Aux_Acu_Monto = 0
                'Liquidacion actual
                If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                    Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
                End If
                
                'Mas el acumulador mensual
                StrSql = "SELECT * FROM sim_acu_mes "
                StrSql = StrSql & " WHERE ternro = " & buliq_empleado!Ternro
                StrSql = StrSql & " AND acunro = " & Acum
                StrSql = StrSql & " AND amanio = " & Ret_Ano
                StrSql = StrSql & " AND ammes = " & Ret_Mes
                OpenRecordset StrSql, rs_Acum
                If Not rs_Acum.EOF Then
                    Aux_Acu_Monto = Aux_Acu_Monto + rs_Acum!ammonto
                End If
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
                End If

                rs_itemacum.MoveNext
            Loop




'FGZ - 14/06/2007 - en lugar de buscar acu_liq vamos a buscar acu_mes -------

'FGZ - 14/06/2007 - No se pueden configurar conceptos -------
'            ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
'            ' Busco los conceptos de la liquidacion
'            StrSql = "SELECT * FROM itemconc " & _
'                     " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
'                     " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
'                     " AND itemconc.itenro =" & rs_Item!Itenro & _
'                     " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
'            OpenRecordset StrSql, rs_itemconc
'            Do While Not rs_itemconc.EOF
'                    If CBool(rs_itemconc!itcsigno) Then
'                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + rs_itemconc!dlimonto
'                    Else
'                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - rs_itemconc!dlimonto
'                    End If
'                rs_itemconc.MoveNext
'            Loop
'FGZ - 14/06/2007 - No se pueden configurar conceptos -------

            'Armo la traza del item
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Item " & rs_Item!Itenro
                Flog.writeline Espacios(Tabulador * 4) & "LIQ" & Items_LIQ(rs_Item!Itenro)
                Flog.writeline Espacios(Tabulador * 4) & "DDJJ" & Items_DDJJ(rs_Item!Itenro)
            End If
            If HACE_TRAZA Then
                Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
                Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_LIQ(rs_Item!Itenro))
            End If
            
            'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
            ' "ACHIQUE" DE GANANCIA IMPONIBLE
            If CBool(rs_Item!itesigno) Then
                Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro)
                Sum_Renta = Sum_Renta + Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            Else
                Items_LIQ(rs_Item!Itenro) = -Items_LIQ(rs_Item!Itenro)
                Sum_Deduccion = Sum_Deduccion + Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            End If
                
        rs_Item.MoveNext
    Loop
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Renta        : " & Sum_Renta
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Deducciones  : " & Sum_Deduccion
        Flog.writeline
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Renta", Sum_Renta)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Deducciones", Sum_Deduccion)
    End If
    
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " nogan =" & Sum_Renta
    StrSql = StrSql & " ,otras =" & Sum_Deduccion
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    '**************************************************************************
    
    'Busco en escala para Renta
    'cabecera
    StrSql = "SELECT * FROM uru_irpfcab "
    StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND fecdesde <= " & ConvFecha(buliq_proceso!profecpago)
    OpenRecordset StrSql, rs_IrpfCab
    If rs_IrpfCab.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Renta. No existe escala para la fecha de pago"
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - No existe escala para la fecha de pago", 0)
        End If
        Exit Function
    Else
        BPC_Renta = rs_IrpfCab!val
        Cabecera_Renta = rs_IrpfCab!cabnro
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "BPC para Renta " & BPC_Renta
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - BPC para Renta", rs_IrpfCab!val)
    End If
    'Detalle Renta
    StrSql = "SELECT * FROM uru_irpfdet "
    StrSql = StrSql & " WHERE cabnro = " & Cabecera_Renta
    StrSql = StrSql & " AND valdesde * " & BPC_Renta & " < " & Sum_Renta
    StrSql = StrSql & " ORDER BY valdesde"
    OpenRecordset StrSql, rs_IrpfDet
    If rs_IrpfDet.EOF Then
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - No Entra en Franja de Escala", Sum_Renta)
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "No Entra en Franja de Escala" & Sum_Renta
        End If
    End If
    Do While Not rs_IrpfDet.EOF
        If (rs_IrpfDet!valhasta * BPC_Renta) <= Sum_Renta Then
            Renta = Renta + ((rs_IrpfDet!valhasta * BPC_Renta) - (rs_IrpfDet!valdesde * BPC_Renta)) * rs_IrpfDet!Porc / 100
        Else
            Renta = Renta + ((Sum_Renta) - (rs_IrpfDet!valdesde * BPC_Renta)) * rs_IrpfDet!Porc / 100
        End If
    
        rs_IrpfDet.MoveNext
    Loop
    If HACE_TRAZA Then
        'Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "93 - Porcentaje Renta ", Por_Renta)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Renta", Renta)
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline
        'Flog.writeline Espacios(Tabulador * 3) & "Porcentaje Renta  :" & Por_Renta
        Flog.writeline Espacios(Tabulador * 3) & "Renta             :" & Renta
    End If
    
    'Busco en escala para Deduccion
    'cabecera
    StrSql = "SELECT * FROM uru_irpfdedcab "
    StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND fecdesde <= " & ConvFecha(buliq_proceso!profecpago)
    OpenRecordset StrSql, rs_IrpfDedCab
    If rs_IrpfDedCab.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Deducción. No existe escala para la fecha de pago"
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - Deducción. No existe escala para la fecha de pago", 0)
        End If
        Exit Function
    Else
        BPC_Deduccion = rs_IrpfDedCab!val
        Cabecera_Deduccion = rs_IrpfDedCab!cabnro
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "BPC para Deducción " & BPC_Deduccion
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - BPC para Deducción", BPC_Deduccion)
    End If
    'Detalle Renta
    StrSql = "SELECT * FROM uru_irpfdeddet "
    StrSql = StrSql & " WHERE cabnro = " & Cabecera_Deduccion
    StrSql = StrSql & " AND valdesde * " & BPC_Deduccion & " < " & Sum_Deduccion
    StrSql = StrSql & " ORDER BY valdesde"
    OpenRecordset StrSql, rs_IrpfDedDet
    If rs_IrpfDedDet.EOF Then
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - No Entra en Franja de Escala", Sum_Deduccion)
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "No Entra en Franja de Escala" & Sum_Deduccion
        End If
    End If
    Do While Not rs_IrpfDedDet.EOF
        If (rs_IrpfDedDet!valhasta * BPC_Deduccion) <= Sum_Deduccion Then
            Deduccion = Deduccion + ((rs_IrpfDedDet!valhasta * BPC_Deduccion) - (rs_IrpfDedDet!valdesde * BPC_Deduccion)) * rs_IrpfDedDet!Porc / 100
        Else
            Deduccion = Deduccion + (Sum_Deduccion - (rs_IrpfDedDet!valdesde * BPC_Deduccion)) * rs_IrpfDedDet!Porc / 100
        End If
    
        rs_IrpfDedDet.MoveNext
    Loop
    If HACE_TRAZA Then
        'Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "93 - Porcentaje Deducción ", Por_Deduccion)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Deducción", Deduccion)
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline
        'Flog.writeline Espacios(Tabulador * 3) & "Porcentaje Deducción  :" & Por_Deduccion
        Flog.writeline Espacios(Tabulador * 3) & "Deducción             :" & Deduccion
    End If
    
    'Busco otras retenciones efectuadas en el mes (en otros procesos)
    Otras_Ret = 0
        
    StrSql = "SELECT * FROM sim_ficharet "
    StrSql = StrSql & " WHERE empleado =" & buliq_empleado!Ternro
    'FGZ - 30/07/2012 --------------------------------------------
    'StrSql = StrSql & " AND month(fecha) = " & Ret_Mes
    'StrSql = StrSql & " AND year(fecha) = " & Ret_Ano
    StrSql = StrSql & " AND fecha >= " & ConvFecha(CDate("01/" & Ret_Mes & "/" & Ret_Ano)) & " AND fecha <= " & ConvFecha(UltimoDiaMes(Ret_Ano, Ret_Mes))
    'FGZ - 30/07/2012 --------------------------------------------
    StrSql = StrSql & " ORDER BY fecha desc"
    OpenRecordset StrSql, rs_Ficharet
    If Not rs_Ficharet.EOF Then
       Otras_Ret = rs_Ficharet!importe
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Otras Retenciones:" & Otras_Ret
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "95 - Otras Retenciones ", Otras_Ret)
    End If
    
    'Resultado del impuesto
    'Impuesto = (Renta - Deduccion - Otras_Ret) * -1
    'No se deben restar otras retenciones
    Impuesto = (Renta - Deduccion) * -1
    Call InsertarFichaRet(buliq_empleado!Ternro, buliq_proceso!profecpago, Abs(Impuesto), buliq_proceso!pronro)
Else

End If
StrSql = "UPDATE sim_traza_gan SET "
StrSql = StrSql & " ganneta =" & Impuesto
StrSql = StrSql & " ,ganimpo =" & Renta
StrSql = StrSql & " ,deducciones =" & Deduccion
StrSql = StrSql & " ,msr =" & BPC_Renta
StrSql = StrSql & " ,nomsr =" & BPC_Deduccion
StrSql = StrSql & " ,noimpo =" & Por_Renta
StrSql = StrSql & " ,porcdeduc =" & Por_Deduccion
StrSql = StrSql & " ,retenciones =" & Otras_Ret
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords


' Grabo todos los items de la liquidacion actual
I = 1
Hasta = 50
Do While I <= Hasta
    
    If Items_LIQ(I) <> 0 Then
       'Busco las liquidaciones anteriores
        StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                 buliq_empleado!Ternro & "," & _
                 ConvFecha(buliq_proceso!profecpago) & "," & _
                 buliq_proceso!pronro & "," & _
                 Items_LIQ(I) & "," & _
                 "0," & _
                 I & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    'guardo los item_ddjj para poder usarlo en el reporte de Ganancias
    If Items_DDJJ(I) <> 0 Then
        StrSql = "SELECT * FROM sim_traza_gan_item_top "
        StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND itenro =" & I
        OpenRecordset StrSql, rs_Traza_gan_items_tope
        If rs_Traza_gan_items_tope.EOF Then
            StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,ddjj,empresa,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_DDJJ(I) & "," & _
                     NroEmp & "," & _
                     I & _
                     ")"
        Else 'Actualizo
            StrSql = "UPDATE sim_traza_gan_item_top SET ddjj =" & Items_DDJJ(I)
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

    'guardo los item_liq para poder usarlo en el reporte de Ganancias
    If Items_LIQ(I) <> 0 Then
        StrSql = "SELECT * FROM sim_traza_gan_item_top "
        StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND itenro =" & I
        OpenRecordset StrSql, rs_Traza_gan_items_tope
        If rs_Traza_gan_items_tope.EOF Then
            StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,liq,empresa,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_LIQ(I) & "," & _
                     NroEmp & "," & _
                     I & _
                     ")"
        Else 'Actualizo
            StrSql = "UPDATE sim_traza_gan_item_top SET liq =" & Items_LIQ(I)
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    I = I + 1
Loop

Monto = Impuesto
Bien = True

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Impuesto: " & Impuesto
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "99 - Impuesto", Impuesto)
End If

for_irpf_OLD = Monto
exito = Bien

'Cierro y libero todo
    If rs_IrpfCab.State = adStateOpen Then rs_IrpfCab.Close
    If rs_IrpfDet.State = adStateOpen Then rs_IrpfDet.Close
    If rs_IrpfDedCab.State = adStateOpen Then rs_IrpfDedCab.Close
    If rs_IrpfDedDet.State = adStateOpen Then rs_IrpfDedDet.Close
    If rs_Ficharet.State = adStateOpen Then rs_Ficharet.Close
    If rs_Item.State = adStateOpen Then rs_Item.Close
    If rs_Desmen.State = adStateOpen Then rs_Desmen.Close
    If rs_itemacum.State = adStateOpen Then rs_itemacum.Close
    If rs_itemconc.State = adStateOpen Then rs_itemconc.Close
    If rs_Traza_gan_items_tope.State = adStateOpen Then rs_Traza_gan_items_tope.Close
    If rs_Acum.State = adStateOpen Then rs_Acum.Close
    
    Set rs_IrpfCab = Nothing
    Set rs_IrpfDet = Nothing
    Set rs_IrpfDedCab = Nothing
    Set rs_IrpfDedDet = Nothing
    Set rs_Ficharet = Nothing
    Set rs_Item = Nothing
    Set rs_Desmen = Nothing
    Set rs_itemacum = Nothing
    Set rs_itemconc = Nothing
    Set rs_Traza_gan_items_tope = Nothing
    Set rs_Acum = Nothing
End Function


Public Function for_irpf_diciembre(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias para uruguay.
' Autor      : FGZ
' Fecha      : 05/06/2007
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim Sum_Renta As Double
Dim Sum_Deduccion As Double
Dim BPC_Renta As Double
Dim BPC_Deduccion As Double
Dim Por_Renta As Double
Dim Por_Deduccion As Double
Dim Renta As Double
Dim Deduccion As Double
Dim Impuesto As Double
Dim Otras_Ret As Double
Dim Items_LIQ(100) As Double
Dim Items_DDJJ(100) As Double

'Parametros
Dim c_Aplica_Imp As Long
Dim c_Meses As Long
Dim c_Busq As Long
Dim v_Aplica_Imp As Long
Dim v_Meses As Long
Dim v_Busq As Long

'Auxiliares
Dim Ret_Mes As Integer
Dim Ret_Ano As Integer
Dim Con_liquid As Integer
Dim Acum As Long
Dim Aux_Acu_Monto As Double
Dim Cabecera_Renta As Long
Dim Cabecera_Deduccion As Long
Dim Anual As Boolean
Dim I As Long
Dim Hasta As Long


Dim rs_IrpfCab As New ADODB.Recordset
Dim rs_IrpfDet As New ADODB.Recordset
Dim rs_IrpfDedCab As New ADODB.Recordset
Dim rs_IrpfDedDet As New ADODB.Recordset
Dim rs_Ficharet As New ADODB.Recordset
Dim rs_Item As New ADODB.Recordset
Dim rs_Desmen As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Traza_gan_items_tope As New ADODB.Recordset
Dim rs_Acum As New ADODB.Recordset

'Inicializo
Bien = False
Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro

'defaults
c_Aplica_Imp = 140
c_Meses = 150
c_Busq = 52
v_Aplica_Imp = 0
v_Meses = 12
v_Busq = 0

'Primero limpio la traza
StrSql = "DELETE FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords
If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).ConcNro)
End If

'Guardo la traza de Ganancia en traza_gan para utilizarla luego cuando se genere el reporte
StrSql = "INSERT INTO sim_traza_gan (pliqnro,pronro,concnro,fecha_pago,empresa,ternro,empleg) VALUES ("
StrSql = StrSql & buliq_periodo!PliqNro & ","
StrSql = StrSql & buliq_proceso!pronro & ","
StrSql = StrSql & Buliq_Concepto(Concepto_Actual).ConcNro & ","
StrSql = StrSql & ConvFecha(buliq_proceso!profecpago) & ","
StrSql = StrSql & NroEmp & ","
StrSql = StrSql & buliq_empleado!Ternro & ","
StrSql = StrSql & buliq_empleado!Empleg
StrSql = StrSql & ")"
objConn.Execute StrSql, , adExecuteNoRecords


'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa
Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case c_Aplica_Imp:
        v_Aplica_Imp = rs_wf_tpa!Valor
    Case c_Meses:
        v_Meses = rs_wf_tpa!Valor
    Case c_Busq:
        v_Busq = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop
If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Aplica Imponible " & CBool(v_Aplica_Imp)
    Flog.writeline Espacios(Tabulador * 1) & "Meses " & v_Meses
End If
'Genero la traza de los familiares deducidos para IRPF
If HACE_TRAZA Then
    Call Traza_DeduccionFliaresIRPF(v_Busq)
End If
'**************************************************************************

If CBool(v_Aplica_Imp) Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "No se debe Liquidar el concepto."
    End If
    Bien = True
    exito = Bien
    Exit Function
End If

'Busco el periodo que esoy liquidando para ver si tiene la marca de ultimo periodo
Anual = CBool(buliq_periodo!pliqultimo)

If Not Anual Then
    ' Recorro todos los items de Ganancias
    StrSql = "SELECT * FROM item ORDER BY itetipotope"
    OpenRecordset StrSql, rs_Item
    Do While Not rs_Item.EOF
            'Tomo los valores de DDJJ y Liquidacion sin Tope
            'Busco la declaracion jurada
            StrSql = "SELECT * FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                     " AND desano=" & Ret_Ano & _
                     " AND itenro = " & rs_Item!Itenro
            OpenRecordset StrSql, rs_Desmen
            Do While Not rs_Desmen.EOF
                If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                End If
                rs_Desmen.MoveNext
            Loop
            
'FGZ - 14/06/2007 - en lugar de buscar acu_liq vamos a buscar acu_mes -------
            'Busco los acumuladores de la liquidacion
'            StrSql = "SELECT * FROM itemacum " & _
'                     " WHERE itenro =" & rs_Item!Itenro & _
'                     " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
'            OpenRecordset StrSql, rs_itemacum
'            Do While Not rs_itemacum.EOF
'                Acum = CStr(rs_itemacum!acunro)
'                If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
'                    Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
'
'                    If CBool(rs_itemacum!itasigno) Then
'                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
'                    Else
'                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
'                    End If
'                End If
'                rs_itemacum.MoveNext
'            Loop
            
            StrSql = "SELECT * FROM itemacum " & _
                     " WHERE itenro =" & rs_Item!Itenro & _
                     " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
            OpenRecordset StrSql, rs_itemacum
            Do While Not rs_itemacum.EOF
                Acum = CStr(rs_itemacum!acuNro)
                Aux_Acu_Monto = 0
                'Liquidacion actual
                If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                    Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
                End If
                
                'Mas el acumulador mensual
                StrSql = "SELECT * FROM sim_acu_mes "
                StrSql = StrSql & " WHERE ternro = " & buliq_empleado!Ternro
                StrSql = StrSql & " AND acunro = " & Acum
                StrSql = StrSql & " AND amanio = " & Ret_Ano
                StrSql = StrSql & " AND ammes = " & Ret_Mes
                OpenRecordset StrSql, rs_Acum
                If Not rs_Acum.EOF Then
                    Aux_Acu_Monto = Aux_Acu_Monto + rs_Acum!ammonto
                End If
                If CBool(rs_itemacum!itasigno) Then
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
                Else
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
                End If

                rs_itemacum.MoveNext
            Loop




'FGZ - 14/06/2007 - en lugar de buscar acu_liq vamos a buscar acu_mes -------

'FGZ - 14/06/2007 - No se pueden configurar conceptos -------
'            ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
'            ' Busco los conceptos de la liquidacion
'            StrSql = "SELECT * FROM itemconc " & _
'                     " INNER JOIN detliq ON itemconc.concnro = detliq.concnro " & _
'                     " WHERE detliq.cliqnro = " & buliq_cabliq!cliqnro & _
'                     " AND itemconc.itenro =" & rs_Item!Itenro & _
'                     " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
'            OpenRecordset StrSql, rs_itemconc
'            Do While Not rs_itemconc.EOF
'                    If CBool(rs_itemconc!itcsigno) Then
'                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + rs_itemconc!dlimonto
'                    Else
'                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - rs_itemconc!dlimonto
'                    End If
'                rs_itemconc.MoveNext
'            Loop
'FGZ - 14/06/2007 - No se pueden configurar conceptos -------

            'Armo la traza del item
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Item " & rs_Item!Itenro
                Flog.writeline Espacios(Tabulador * 4) & "LIQ" & Items_LIQ(rs_Item!Itenro)
                Flog.writeline Espacios(Tabulador * 4) & "DDJJ" & Items_DDJJ(rs_Item!Itenro)
            End If
            If HACE_TRAZA Then
                Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
                Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_LIQ(rs_Item!Itenro))
            End If
            
            'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
            ' "ACHIQUE" DE GANANCIA IMPONIBLE
            If CBool(rs_Item!itesigno) Then
                Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro)
                Sum_Renta = Sum_Renta + Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            Else
                Items_LIQ(rs_Item!Itenro) = -Items_LIQ(rs_Item!Itenro)
                Sum_Deduccion = Sum_Deduccion + Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            End If
                
        rs_Item.MoveNext
    Loop
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Renta        : " & Sum_Renta
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Deducciones  : " & Sum_Deduccion
        Flog.writeline
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Renta", Sum_Renta)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Deducciones", Sum_Deduccion)
    End If
    
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " nogan =" & Sum_Renta
    StrSql = StrSql & " ,otras =" & Sum_Deduccion
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    '**************************************************************************
    
    'Busco en escala para Renta
    'cabecera
    StrSql = "SELECT * FROM uru_irpfcab "
    StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(CDate("25/12/" & Year(buliq_proceso!profecpago)))
    StrSql = StrSql & " AND fecdesde <= " & ConvFecha(CDate("25/12/" & Year(buliq_proceso!profecpago)))
    'StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND fecdesde <= " & ConvFecha(buliq_proceso!profecpago)
    OpenRecordset StrSql, rs_IrpfCab
    If rs_IrpfCab.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Renta. No existe escala para la fecha de pago"
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - No existe escala para la fecha de pago", 0)
        End If
        Exit Function
    Else
        BPC_Renta = rs_IrpfCab!val
        Cabecera_Renta = rs_IrpfCab!cabnro
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "BPC para Renta " & BPC_Renta
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - BPC para Renta", rs_IrpfCab!val)
    End If
    'Detalle Renta
    StrSql = "SELECT * FROM uru_irpfdet "
    StrSql = StrSql & " WHERE cabnro = " & Cabecera_Renta
    StrSql = StrSql & " AND valdesde * " & BPC_Renta & " < " & Sum_Renta
    StrSql = StrSql & " ORDER BY valdesde"
    OpenRecordset StrSql, rs_IrpfDet
    If rs_IrpfDet.EOF Then
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - No Entra en Franja de Escala", Sum_Renta)
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "No Entra en Franja de Escala" & Sum_Renta
        End If
    End If
    Do While Not rs_IrpfDet.EOF
        If (rs_IrpfDet!valhasta * BPC_Renta) <= Sum_Renta Then
            Renta = Renta + ((rs_IrpfDet!valhasta * BPC_Renta) - (rs_IrpfDet!valdesde * BPC_Renta)) * rs_IrpfDet!Porc / 100
        Else
            Renta = Renta + ((Sum_Renta) - (rs_IrpfDet!valdesde * BPC_Renta)) * rs_IrpfDet!Porc / 100
        End If
    
        rs_IrpfDet.MoveNext
    Loop
    If HACE_TRAZA Then
        'Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "93 - Porcentaje Renta ", Por_Renta)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Renta", Renta)
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline
        'Flog.writeline Espacios(Tabulador * 3) & "Porcentaje Renta  :" & Por_Renta
        Flog.writeline Espacios(Tabulador * 3) & "Renta             :" & Renta
    End If
    
    'Busco en escala para Deduccion
    'cabecera
    StrSql = "SELECT * FROM uru_irpfdedcab "
    'StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND fecdesde <= " & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(CDate("25/12/" & Year(buliq_proceso!profecpago)))
    StrSql = StrSql & " AND fecdesde <= " & ConvFecha(CDate("25/12/" & Year(buliq_proceso!profecpago)))
    OpenRecordset StrSql, rs_IrpfDedCab
    If rs_IrpfDedCab.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Deducción. No existe escala para la fecha de pago"
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - Deducción. No existe escala para la fecha de pago", 0)
        End If
        Exit Function
    Else
        BPC_Deduccion = rs_IrpfDedCab!val
        Cabecera_Deduccion = rs_IrpfDedCab!cabnro
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "BPC para Deducción " & BPC_Deduccion
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - BPC para Deducción", BPC_Deduccion)
    End If
    'Detalle Renta
    StrSql = "SELECT * FROM uru_irpfdeddet "
    StrSql = StrSql & " WHERE cabnro = " & Cabecera_Deduccion
    StrSql = StrSql & " AND valdesde * " & BPC_Deduccion & " < " & Sum_Deduccion
    StrSql = StrSql & " ORDER BY valdesde"
    OpenRecordset StrSql, rs_IrpfDedDet
    If rs_IrpfDedDet.EOF Then
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - No Entra en Franja de Escala", Sum_Deduccion)
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "No Entra en Franja de Escala" & Sum_Deduccion
        End If
    End If
    Do While Not rs_IrpfDedDet.EOF
        If (rs_IrpfDedDet!valhasta * BPC_Deduccion) <= Sum_Deduccion Then
            Deduccion = Deduccion + ((rs_IrpfDedDet!valhasta * BPC_Deduccion) - (rs_IrpfDedDet!valdesde * BPC_Deduccion)) * rs_IrpfDedDet!Porc / 100
        Else
            Deduccion = Deduccion + (Sum_Deduccion - (rs_IrpfDedDet!valdesde * BPC_Deduccion)) * rs_IrpfDedDet!Porc / 100
        End If
    
        rs_IrpfDedDet.MoveNext
    Loop
    If HACE_TRAZA Then
        'Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "93 - Porcentaje Deducción ", Por_Deduccion)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Deducción", Deduccion)
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline
        'Flog.writeline Espacios(Tabulador * 3) & "Porcentaje Deducción  :" & Por_Deduccion
        Flog.writeline Espacios(Tabulador * 3) & "Deducción             :" & Deduccion
    End If
    
    'Busco otras retenciones efectuadas en el mes (en otros procesos)
    Otras_Ret = 0
        
    StrSql = "SELECT * FROM sim_ficharet "
    StrSql = StrSql & " WHERE empleado =" & buliq_empleado!Ternro
    'FGZ - 30/07/2012 --------------------------------------------
    'StrSql = StrSql & " AND month(fecha) = " & Ret_Mes
    'StrSql = StrSql & " AND year(fecha) = " & Ret_Ano
    StrSql = StrSql & " AND fecha >= " & ConvFecha(CDate("01/" & Ret_Mes & "/" & Ret_Ano)) & " AND fecha <= " & ConvFecha(UltimoDiaMes(Ret_Ano, Ret_Mes))
    'FGZ - 30/07/2012 --------------------------------------------
    StrSql = StrSql & " ORDER BY fecha desc"
    OpenRecordset StrSql, rs_Ficharet
    If Not rs_Ficharet.EOF Then
       Otras_Ret = rs_Ficharet!importe
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Otras Retenciones:" & Otras_Ret
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "95 - Otras Retenciones ", Otras_Ret)
    End If
    
    'Resultado del impuesto
    'Impuesto = (Renta - Deduccion - Otras_Ret) * -1
    'No se deben restar otras retenciones
    Impuesto = (Renta - Deduccion) * -1
    Call InsertarFichaRet(buliq_empleado!Ternro, buliq_proceso!profecpago, Abs(Impuesto), buliq_proceso!pronro)
Else

End If
StrSql = "UPDATE sim_traza_gan SET "
StrSql = StrSql & " ganneta =" & Impuesto
StrSql = StrSql & " ,ganimpo =" & Renta
StrSql = StrSql & " ,deducciones =" & Deduccion
StrSql = StrSql & " ,msr =" & BPC_Renta
StrSql = StrSql & " ,nomsr =" & BPC_Deduccion
StrSql = StrSql & " ,noimpo =" & Por_Renta
StrSql = StrSql & " ,porcdeduc =" & Por_Deduccion
StrSql = StrSql & " ,retenciones =" & Otras_Ret
StrSql = StrSql & " WHERE "
StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords


' Grabo todos los items de la liquidacion actual
I = 1
Hasta = 50
Do While I <= Hasta
    
    If Items_LIQ(I) <> 0 Then
       'Busco las liquidaciones anteriores
        StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                 buliq_empleado!Ternro & "," & _
                 ConvFecha(buliq_proceso!profecpago) & "," & _
                 buliq_proceso!pronro & "," & _
                 Items_LIQ(I) & "," & _
                 "0," & _
                 I & _
                 ")"
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    'guardo los item_ddjj para poder usarlo en el reporte de Ganancias
    If Items_DDJJ(I) <> 0 Then
        StrSql = "SELECT * FROM sim_traza_gan_item_top "
        StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND itenro =" & I
        OpenRecordset StrSql, rs_Traza_gan_items_tope
        If rs_Traza_gan_items_tope.EOF Then
            StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,ddjj,empresa,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_DDJJ(I) & "," & _
                     NroEmp & "," & _
                     I & _
                     ")"
        Else 'Actualizo
            StrSql = "UPDATE sim_traza_gan_item_top SET ddjj =" & Items_DDJJ(I)
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If

    'guardo los item_liq para poder usarlo en el reporte de Ganancias
    If Items_LIQ(I) <> 0 Then
        StrSql = "SELECT * FROM sim_traza_gan_item_top "
        StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND itenro =" & I
        OpenRecordset StrSql, rs_Traza_gan_items_tope
        If rs_Traza_gan_items_tope.EOF Then
            StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,liq,empresa,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_LIQ(I) & "," & _
                     NroEmp & "," & _
                     I & _
                     ")"
        Else 'Actualizo
            StrSql = "UPDATE sim_traza_gan_item_top SET liq =" & Items_LIQ(I)
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
    
    I = I + 1
Loop

Monto = Impuesto
Bien = True

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Impuesto: " & Impuesto
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "99 - Impuesto", Impuesto)
End If

for_irpf_diciembre = Monto
exito = Bien

'Cierro y libero todo
    If rs_IrpfCab.State = adStateOpen Then rs_IrpfCab.Close
    If rs_IrpfDet.State = adStateOpen Then rs_IrpfDet.Close
    If rs_IrpfDedCab.State = adStateOpen Then rs_IrpfDedCab.Close
    If rs_IrpfDedDet.State = adStateOpen Then rs_IrpfDedDet.Close
    If rs_Ficharet.State = adStateOpen Then rs_Ficharet.Close
    If rs_Item.State = adStateOpen Then rs_Item.Close
    If rs_Desmen.State = adStateOpen Then rs_Desmen.Close
    If rs_itemacum.State = adStateOpen Then rs_itemacum.Close
    If rs_itemconc.State = adStateOpen Then rs_itemconc.Close
    If rs_Traza_gan_items_tope.State = adStateOpen Then rs_Traza_gan_items_tope.Close
    If rs_Acum.State = adStateOpen Then rs_Acum.Close
    
    Set rs_IrpfCab = Nothing
    Set rs_IrpfDet = Nothing
    Set rs_IrpfDedCab = Nothing
    Set rs_IrpfDedDet = Nothing
    Set rs_Ficharet = Nothing
    Set rs_Item = Nothing
    Set rs_Desmen = Nothing
    Set rs_itemacum = Nothing
    Set rs_itemconc = Nothing
    Set rs_Traza_gan_items_tope = Nothing
    Set rs_Acum = Nothing
End Function

Public Function for_irpf_simple(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias para uruguay. Calculo simple.
' Autor      : FGZ
' Fecha      : 26/06/2007
' Ultima Mod.: FGZ - 30/06/2015 - modificaciones por Ley No. 19.321 de 29/05/2015. Decreto 154/015 de 1/06/2015.
' ---------------------------------------------------------------------------------------------
Dim Sum_Renta As Double
Dim Sum_Renta_VAC As Double
Dim Sum_Renta_SAC As Double
Dim Sum_Deduccion As Double
Dim BPC_Renta As Double
Dim BPC_Deduccion As Double
Dim Por_Renta As Double
Dim Aux_Por_Renta As Double
Dim Por_Deduccion As Double
Dim Renta As Double
Dim Deduccion As Double
Dim Impuesto As Double
Dim Otras_Ret As Double
'FGZ - 12/05/2014 --------------
'Dim Items_LIQ(50) As Double
'Dim Items_DDJJ(50) As Double
Dim Items_LIQ(100) As Double
Dim Items_DDJJ(100) As Double

'Parametros
Dim c_Imp_Renta As Long
Dim c_Imp_Deduccion As Long
Dim c_Multiplicador As Long
Dim c_Imp_Renta_Vac As Long
Dim c_Imp_Renta_SAC As Long
Dim v_Imp_Renta As Double
Dim v_Imp_Renta_Vac As Double
Dim v_Imp_Renta_Sac As Double
Dim v_Imp_Deduccion As Double
Dim v_Multiplicador As Double

'Auxiliares
Dim Ret_Mes As Integer
Dim Ret_Ano As Integer
Dim Con_liquid As Integer
Dim Acum As Long
Dim Aux_Acu_Monto As Double
Dim Cabecera_Renta As Long
Dim Cabecera_Deduccion As Long
Dim Anual As Boolean
Dim I As Long
Dim Hasta As Long


Dim rs_IrpfCab As New ADODB.Recordset
Dim rs_IrpfDet As New ADODB.Recordset
Dim rs_IrpfDedCab As New ADODB.Recordset
Dim rs_IrpfDedDet As New ADODB.Recordset
Dim rs_Ficharet As New ADODB.Recordset
Dim rs_Item As New ADODB.Recordset
Dim rs_Desmen As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Traza_gan_items_tope As New ADODB.Recordset
Dim rs_Acum As New ADODB.Recordset

'Inicializo
Bien = False
Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro

'defaults
c_Imp_Renta = 1015
c_Imp_Deduccion = 1016
c_Multiplicador = 149

'FGZ - 30/06/2015 ----------
c_Imp_Renta_Vac = 1061
c_Imp_Renta_SAC = 1062
'FGZ - 30/06/2015 ----------

v_Imp_Renta = 0
v_Imp_Deduccion = 0
v_Multiplicador = 1

'FGZ - 30/06/2015 ----------
v_Imp_Renta_Vac = 0
v_Imp_Renta_Sac = 0
'FGZ - 30/06/2015 ----------

'Obtencion de los parametros de WorkFile
'FGZ - 24/05/2011 ------------------------------------
'rs_wf_tpa!valor por     Arr_WF_TPA(I).valor
For I = LI_WF_Tpa To LS_WF_Tpa
    Select Case Arr_WF_TPA(I).tipoparam
    Case c_Imp_Renta:
        v_Imp_Renta = Arr_WF_TPA(I).Valor
    Case c_Imp_Renta_Vac:
        v_Imp_Renta_Vac = Arr_WF_TPA(I).Valor
    Case c_Imp_Renta_SAC:
        v_Imp_Renta_Sac = Arr_WF_TPA(I).Valor
    Case c_Imp_Deduccion:
        v_Imp_Deduccion = Arr_WF_TPA(I).Valor
    Case c_Multiplicador:
        v_Multiplicador = Arr_WF_TPA(I).Valor
    Case Else
    End Select
Next I

'StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
'OpenRecordset StrSql, rs_wf_tpa
'Do While Not rs_wf_tpa.EOF
'    Select Case rs_wf_tpa!tipoparam
'    Case c_Imp_Renta:
'        v_Imp_Renta = rs_wf_tpa!valor
'    Case c_Imp_Deduccion:
'        v_Imp_Deduccion = rs_wf_tpa!valor
'    Case c_Multiplicador:
'        v_Multiplicador = rs_wf_tpa!valor
'    End Select
'
'    rs_wf_tpa.MoveNext
'Loop
'FGZ - 24/05/2011 ------------------------------------

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Imponible para Renta                      " & CBool(v_Imp_Renta)
    Flog.writeline Espacios(Tabulador * 1) & "Imponible para Renta para Vacaciones      " & CBool(v_Imp_Renta_Vac)
    Flog.writeline Espacios(Tabulador * 1) & "Imponible para Renta para Aguinaldo       " & CBool(v_Imp_Renta_Sac)
    Flog.writeline Espacios(Tabulador * 1) & "Imponible para Deducción  " & CBool(v_Imp_Deduccion)
    Flog.writeline Espacios(Tabulador * 1) & "Multiplicador             " & CDbl(v_Multiplicador)
End If

'**************************************************************************

    Sum_Renta = v_Imp_Renta
    Sum_Renta_VAC = v_Imp_Renta_Vac
    Sum_Renta_SAC = v_Imp_Renta_Sac
    Sum_Deduccion = v_Imp_Deduccion
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Renta              : " & Sum_Renta
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Vacaciones         : " & Sum_Renta_VAC
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Aguinaldo          : " & Sum_Renta_SAC
        Flog.writeline Espacios(Tabulador * 3) & " Acumulado Deducciones  : " & Sum_Deduccion
        Flog.writeline
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Renta", Sum_Renta)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Renta para Vac", Sum_Renta_VAC)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Renta para Aguinaldo", Sum_Renta_SAC)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - Base para Deducciones", Sum_Deduccion)
    End If
    
    '**************************************************************************
    
    'Busco en escala para Renta
    'cabecera
    StrSql = "SELECT * FROM uru_irpfcab "
    StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND fecdesde <= " & ConvFecha(buliq_proceso!profecpago)
    OpenRecordset StrSql, rs_IrpfCab
    If rs_IrpfCab.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Renta. No existe escala para la fecha de pago"
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "91 - No existe escala para la fecha de pago", 0)
        End If
        Exit Function
    Else
        BPC_Renta = rs_IrpfCab!val
        Cabecera_Renta = rs_IrpfCab!cabnro
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "BPC para Renta " & BPC_Renta
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - BPC para Renta", rs_IrpfCab!val)
    End If
    'Detalle Renta
    StrSql = "SELECT * FROM uru_irpfdet "
    StrSql = StrSql & " WHERE cabnro = " & Cabecera_Renta
    StrSql = StrSql & " AND valdesde * " & BPC_Renta * v_Multiplicador & " < " & Sum_Renta
    StrSql = StrSql & " ORDER BY valdesde"
    OpenRecordset StrSql, rs_IrpfDet
    If rs_IrpfDet.EOF Then
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - No Entra en Franja de Escala", Sum_Renta)
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "No Entra en Franja de Escala" & Sum_Renta
        End If
    End If
    'FGZ - 30/06/2015 ----------
    Aux_Por_Renta = 0
    'FGZ - 30/06/2015 ----------
    Do While Not rs_IrpfDet.EOF
        If (rs_IrpfDet!valhasta * BPC_Renta * v_Multiplicador) <= Sum_Renta Then
            Renta = Renta + ((rs_IrpfDet!valhasta * BPC_Renta * v_Multiplicador) - (rs_IrpfDet!valdesde * BPC_Renta * v_Multiplicador)) * rs_IrpfDet!Porc / 100
        Else
            Renta = Renta + ((Sum_Renta) - (rs_IrpfDet!valdesde * BPC_Renta * v_Multiplicador)) * rs_IrpfDet!Porc / 100
            'FGZ -30/06/2015 ----------------------------------------------
            If rs_IrpfDet!Porc > Aux_Por_Renta Then
                Aux_Por_Renta = rs_IrpfDet!Porc
            End If
            'FGZ -30/06/2015 ----------------------------------------------
        End If
    
        rs_IrpfDet.MoveNext
    Loop
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - Porcentaje Maximo Alcanzado", Aux_Por_Renta)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Impuesto sobre Renta", Renta)
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline
        'Flog.writeline Espacios(Tabulador * 3) & "Porcentaje Renta  :" & Por_Renta
        Flog.writeline Espacios(Tabulador * 3) & "Renta             :" & Renta
    End If
    
    'FGZ -30/06/2015 -----------------------------------------------------------------
    'Calculo Renta de Vacaciones
    Sum_Renta_VAC = (Sum_Renta_VAC * Aux_Por_Renta) / 100
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 3) & "Renta para VAC             :" & Sum_Renta_VAC
    End If
    'Renta = Renta + Sum_Renta_VAC
    'FGZ -30/06/2015 -----------------------------------------------------------------
    
    
    'FGZ -30/06/2015 -----------------------------------------------------------------
    'Calculo Renta de Aguinaldo
    Sum_Renta_SAC = (Sum_Renta_SAC * Aux_Por_Renta) / 100
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 3) & "Renta para SAC             :" & Sum_Renta_SAC
    End If
    'Renta = Renta + Sum_Renta_SAC
    'FGZ -30/06/2015 -----------------------------------------------------------------
        
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Impuesto sobre Renta para Aguinaldo ", Sum_Renta_SAC)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Impuesto sobre Renta para VAC", Sum_Renta_VAC)
    End If
    
    'Busco en escala para Deduccion
    'cabecera
    StrSql = "SELECT * FROM uru_irpfdedcab "
    StrSql = StrSql & " WHERE fechasta >= " & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND fecdesde <= " & ConvFecha(buliq_proceso!profecpago)
    OpenRecordset StrSql, rs_IrpfDedCab
    If rs_IrpfDedCab.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Deducción. No existe escala para la fecha de pago"
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - Deducción. No existe escala para la fecha de pago", 0)
        End If
        Exit Function
    Else
        BPC_Deduccion = rs_IrpfDedCab!val
        Cabecera_Deduccion = rs_IrpfDedCab!cabnro
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "BPC para Deducción " & BPC_Deduccion
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "92 - BPC para Deducción", BPC_Deduccion)
    End If
    'Detalle Renta
    StrSql = "SELECT * FROM uru_irpfdeddet "
    StrSql = StrSql & " WHERE cabnro = " & Cabecera_Deduccion
    StrSql = StrSql & " AND valdesde * " & BPC_Deduccion * v_Multiplicador & " < " & Sum_Deduccion
    StrSql = StrSql & " ORDER BY valdesde"
    OpenRecordset StrSql, rs_IrpfDedDet
    If rs_IrpfDedDet.EOF Then
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "93 - No Entra en Franja de Escala", Sum_Deduccion)
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "No Entra en Franja de Escala" & Sum_Deduccion
        End If
    End If
    Do While Not rs_IrpfDedDet.EOF
        If (rs_IrpfDedDet!valhasta * BPC_Deduccion * v_Multiplicador) <= Sum_Deduccion Then
            Deduccion = Deduccion + ((rs_IrpfDedDet!valhasta * BPC_Deduccion * v_Multiplicador) - (rs_IrpfDedDet!valdesde * BPC_Deduccion * v_Multiplicador)) * rs_IrpfDedDet!Porc / 100
        Else
            Deduccion = Deduccion + (Sum_Deduccion - (rs_IrpfDedDet!valdesde * BPC_Deduccion * v_Multiplicador)) * rs_IrpfDedDet!Porc / 100
        End If
    
        rs_IrpfDedDet.MoveNext
    Loop
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "94 - Deducción", Deduccion)
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 3) & "Deducción             :" & Deduccion
    End If
    
    'EAM (6.48) - Se suma al impuesto el SAC y Vac si la renta es positiva
    If (Renta - Deduccion) >= 0 Then
        Renta = Renta + Sum_Renta_VAC
        Renta = Renta + Sum_Renta_SAC
    End If

    
    'Resultado del impuesto
    Impuesto = (Renta - Deduccion) * -1

    Monto = Impuesto
    Bien = True

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Impuesto: " & Impuesto
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "99 - Impuesto", Impuesto)
End If

for_irpf_simple = Monto
exito = Bien

'Cierro y libero todo
    If rs_IrpfCab.State = adStateOpen Then rs_IrpfCab.Close
    If rs_IrpfDet.State = adStateOpen Then rs_IrpfDet.Close
    If rs_IrpfDedCab.State = adStateOpen Then rs_IrpfDedCab.Close
    If rs_IrpfDedDet.State = adStateOpen Then rs_IrpfDedDet.Close
    
    Set rs_IrpfCab = Nothing
    Set rs_IrpfDet = Nothing
    Set rs_IrpfDedCab = Nothing
    Set rs_IrpfDedDet = Nothing
End Function





Public Sub Traza_DeduccionFliaresIRPF(ByVal NroProg As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: genera la traza para Deduccion de Asignaciones Familiares para IRFP
' Autor      : FGZ
' Fecha      : 07/06/2007
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Edad As Integer                 'cant de años o nulo o vacio
Dim Parentesco As Integer           'codigo del parentesco
                            
Dim edad_f As Integer
Dim Fecha_Hasta_Edad As Date
Dim Opcion_Fecha_Hasta As Integer
Dim Fecha_vto_asig As Date
Dim Porcentaje As String
Dim Porc As Double
Dim Incapacitado As String
Dim Nrofam As Long

Dim rs_Familiar As New ADODB.Recordset

    If Arr_Programa(NroProg).Prognro <> 0 Then
        Edad = Arr_Programa(NroProg).Auxint1
        Parentesco = Arr_Programa(NroProg).Auxchar4
        Opcion_Fecha_Hasta = IIf(Not EsNulo(Arr_Programa(NroProg).Auxchar), Arr_Programa(NroProg).Auxchar, 5)
        Select Case Opcion_Fecha_Hasta
        Case 1:
            Fecha_Hasta_Edad = buliq_periodo!pliqhasta
        Case 2:
            Fecha_Hasta_Edad = DateAdd("d", -1, buliq_periodo!pliqdesde)
        Case 3: 'a fin de año
            Fecha_Hasta_Edad = C_Date("31/12/" & Year(buliq_periodo!pliqhasta))
        Case 4: 'a principio de año
            Fecha_Hasta_Edad = C_Date("01/01/" & Year(buliq_periodo!pliqhasta))
        Case 5:
            'Si el empleado tiene fecha de baja < a la fecha fin del proceso, se toma la fecha de baja.
            Fecha_Hasta_Edad = Empleado_Fecha_Fin
        Case Else:
            Fecha_Hasta_Edad = Empleado_Fecha_Fin
        End Select
    Else
        Exit Sub
    End If
    Fecha_vto_asig = buliq_proceso!profecfin

    Texto = "000 - Deduccion de Familiares "
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, 0)

    StrSql = "SELECT tercero.ternro, tercero.terape, tercero.ternom, tercero.terfecnac, familiar.parenro, familiar.famcargadgi, familiar.famemergencia, familiar.faminc"
    StrSql = StrSql & " FROM familiar INNER JOIN tercero ON tercero.ternro = familiar.ternro"
    StrSql = StrSql & " WHERE (familiar.empleado =" & buliq_empleado!Ternro
    StrSql = StrSql & " AND familiar.parenro = " & Parentesco
    StrSql = StrSql & " AND famcargadgi = -1)"
    StrSql = StrSql & " AND (familiar.famDGIdesde <= " & ConvFecha(Fecha_vto_asig) & " OR familiar.famDGIdesde is null)"
    StrSql = StrSql & " AND (familiar.famDGIhasta >= " & ConvFecha(Fecha_vto_asig) & " OR familiar.famDGIhasta is null)"
    StrSql = StrSql & " Order by tercero.ternro"
    OpenRecordset StrSql, rs_Familiar
    Nrofam = 0
    Do While Not rs_Familiar.EOF
        Nrofam = Nrofam + 1
        If Not CBool(rs_Familiar!faminc) Then
            edad_f = Calcular_Edad(rs_Familiar!terfecnac, Fecha_Hasta_Edad)
            If edad_f <= Edad Or EsNulo(Edad) Then
                If CBool(rs_Familiar!faminc) Then
                    If Not CBool(rs_Familiar!famemergencia) Then
                        Incapacitado = "(Incapacitado)"
                        Porcentaje = "200%"
                        Porc = 200
                    Else
                        Incapacitado = "(Incapacitado)"
                        Porcentaje = "100%"
                        Porc = 100
                    End If
                    Texto = "000" & Nrofam & " - " & rs_Familiar!terape & ", " & rs_Familiar!ternom & " --> " & Porcentaje & Incapacitado
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Porc)
                
                    'Grabo en traza_gan
                    StrSql = "UPDATE sim_traza_gan SET "
                    StrSql = StrSql & " entidad" & Nrofam & " ='" & rs_Familiar!terape & ", " & rs_Familiar!ternom & "'"
                    StrSql = StrSql & " , cuit_entidad" & Nrofam & " = '" & rs_Familiar!Ternro & "'"
                    StrSql = StrSql & " , monto_entidad" & Nrofam & " = " & Porc
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    If Not CBool(rs_Familiar!famemergencia) Then
                        Incapacitado = ""
                        Porcentaje = "100%"
                        Porc = 100
                    Else
                        Porcentaje = "50%"
                        Incapacitado = ""
                        Porc = 50
                    End If
                    Texto = "000" & Nrofam & " - " & rs_Familiar!terape & ", " & rs_Familiar!ternom & " --> " & Porcentaje & Incapacitado
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Porc)
                    
                    'Grabo en traza_gan
                    StrSql = "UPDATE sim_traza_gan SET "
                    StrSql = StrSql & " entidad" & Nrofam & " ='" & rs_Familiar!terape & ", " & rs_Familiar!ternom & "'"
                    StrSql = StrSql & " , cuit_entidad" & Nrofam & " = '" & rs_Familiar!Ternro & "'"
                    StrSql = StrSql & " , monto_entidad" & Nrofam & " = " & Porc
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        Else
            If Not CBool(rs_Familiar!famemergencia) Then
                Incapacitado = "(Incapacitado)"
                Porcentaje = "200%"
                Porc = 200
            Else
                Incapacitado = "(Incapacitado)"
                Porcentaje = "100%"
                Porc = 100
            End If
            Texto = "000" & Nrofam & " - " & rs_Familiar!terape & ", " & rs_Familiar!ternom & " --> " & Porcentaje & Incapacitado
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Porc)
        
            'Grabo en traza_gan
            StrSql = "UPDATE sim_traza_gan SET "
            StrSql = StrSql & " entidad" & Nrofam & " ='" & rs_Familiar!terape & ", " & rs_Familiar!ternom & "'"
            StrSql = StrSql & " , cuit_entidad" & Nrofam & " = '" & rs_Familiar!Ternro & "'"
            StrSql = StrSql & " , monto_entidad" & Nrofam & " = " & Porc
            StrSql = StrSql & " WHERE "
            StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
            StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        rs_Familiar.MoveNext
    Loop
    
'Cierro todo y libero
    If rs_Familiar.State = adStateOpen Then rs_Familiar.Close
    Set rs_Familiar = Nothing
End Sub



