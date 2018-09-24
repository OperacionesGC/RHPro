Attribute VB_Name = "MdlFormulasUruguay"
Option Explicit

' ---------------------------------------------------------
' Modulo de fórmulas conocidas para Uruguay
' ---------------------------------------------------------

Public Function for_irp(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Single, Bien As Boolean) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias para uruguay.
' Autor      :
' Fecha      :
' Ultima Mod.:
' Traducccion: FGZ - 13/09/2004
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim c_msr As Integer
Dim c_porcentaje As Integer

Dim v_msr As Single
Dim v_porcentaje As Single

Dim imp_original As Single
Dim nivel_original As Integer
Dim neto_original As Single
Dim porc_original As Single
Dim descuento As Single
Dim neto_secund As Single
Dim porcentaje As Single

Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_uru_irpcab As New ADODB.Recordset
Dim rs_uru_irpdet As New ADODB.Recordset

'Inicializo
c_msr = 8
c_porcentaje = 35

Bien = False

' Primero limpio la traza
StrSql = "DELETE FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
objConn.Execute StrSql, , adExecuteNoRecords


If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If


'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case c_msr:
        v_msr = rs_wf_tpa!Valor
    Case c_porcentaje:
        v_porcentaje = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Imponible del impuesto " & v_msr
    Flog.writeline Espacios(Tabulador * 1) & "Porcentaje " & v_porcentaje
End If

'Si el imponible del impuesto no es positivo, entonces dejo la formula
If Not (v_msr > 0) Then
    Exit Function
End If

'Busco la cabecera del impuesto
StrSql = "SELECT * FROM uru_irpcab "
StrSql = StrSql & " WHERE irpcabfechasta >= " & buliq_proceso!profecpago
StrSql = StrSql & " AND irpcabfecdesde <= " & buliq_proceso!profecpago
OpenRecordset StrSql, rs_uru_irpcab

If rs_uru_irpcab.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe escala para la fecha de pago"
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 1, "No existe escala para la fecha de pago", 0)
    End If
    Exit Function
End If
If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Nro. escala encontrada " & rs_uru_irpcab!irpcabnro
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 1, "Nro. escala encontrada ", rs_uru_irpcab!irpcabnro)
End If
'Fin cabecera del impuesto


'Busco detalle del impuesto: Rangos
StrSql = "SELECT * FROM uru_irpdet "
StrSql = StrSql & " WHERE irpcabnro = " & rs_uru_irpcab!irpcabnro
StrSql = StrSql & " AND irpdetvaldesde < " & v_msr
StrSql = StrSql & " AND irpdetvalhasta >= " & v_msr
OpenRecordset StrSql, rs_uru_irpdet

If rs_uru_irpdet.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe Franja escala monto " & v_msr
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 2, "No existe Franja escala monto", v_msr)
    End If
    Exit Function
End If
If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Franja encontrada " & rs_uru_irpdet!irpdetnivel
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 2, "Franja encontrada ", rs_uru_irpdet!irpdetnivel)
End If

imp_original = v_msr * rs_uru_irpdet!irpdetporc / 100
nivel_original = rs_uru_irpdet!irpdetnivel
neto_original = v_msr * (100 - rs_uru_irpdet!irpdetporc) / 100
porc_original = rs_uru_irpdet!irpdetporc

descuento = -imp_original

'Busco la franja anterior de la misma escala
If nivel_original <> 1 Then 'Cuando no es la primer franja de la escala, busca la franja anterior
    StrSql = "SELECT * FROM uru_irpdet "
    StrSql = StrSql & " WHERE irpcabnro = " & rs_uru_irpcab!irpcabnro
    StrSql = StrSql & " AND irpdetnivel = " & (nivel_original - 1)
    OpenRecordset StrSql, rs_uru_irpdet
    
    If Not rs_uru_irpdet.EOF Then
        neto_secund = rs_uru_irpdet!irpdetvalhasta * (100 - rs_uru_irpdet!irpdetporc) / 100
        If neto_secund > neto_original Then
            descuento = -(rs_uru_irpdet!irpdetvalhasta * rs_uru_irpdet!irpdetporc / 100)
        End If
   End If
End If
porcentaje = -(descuento / v_msr * 100)

'actualzo el parametro
StrSql = "UPDATE " & TTempWF_tpa & " SET valor = " & porcentaje
StrSql = StrSql & " WHERE tpanro =" & c_porcentaje
StrSql = StrSql & " AND fecha = " & ConvFecha(AFecha)
objConn.Execute StrSql, , adExecuteNoRecords

Monto = descuento
Bien = True

for_irp = Monto

'Cierro y libero todo
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
If rs_uru_irpcab.State = adStateOpen Then rs_uru_irpcab.Close
If rs_uru_irpdet.State = adStateOpen Then rs_uru_irpdet.Close

Set rs_wf_tpa = Nothing
Set rs_uru_irpcab = Nothing
Set rs_uru_irpdet = Nothing
End Function
