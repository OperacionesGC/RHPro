Attribute VB_Name = "MdlformulasMoñoAzul"
' ---------------------------------------------------------------------------------------------
' Descripcion: Modulo de Formulas para Moño Azul
' Autor      : FGZ
' Fecha      : 05/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Option Explicit


Public Function for_203() As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Formula para el calculo de produccion: VA-Peras-Emb
'               Reccorre la produccion dia por dia y compara con el minimo, aplica el
'               porcentual de dia trabajado y aplica un monto global
' Autor      : FGZ
' Fecha      : 23/12/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim c_prod_minima_de_peras As Integer
Dim v_prod_minima_de_peras As Single
Dim c_monto As Integer
Dim v_monto As Single

Dim Encontro1 As Boolean
Dim Encontro2 As Boolean

Dim productividad   As Single
Dim nro_empaque As Integer
Dim nro_producto    As Integer
Dim hora_produccion  As Integer
Dim bultos           As Integer
Dim fecha_desde  As Date
Dim fecha_hasta  As Date
Dim cant_bultos  As Single
Dim cant_jornadas As Single
Dim id_empaque   As Integer
Dim descripcion As String
Dim productividad_dia As Single

Dim rs_wf_tpa As New ADODB.Recordset
Dim buf_gti_achdiario As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_gti_achdiario As New ADODB.Recordset

' inicializacion de variables
c_prod_minima_de_peras = 167
c_monto = 51

productividad = 0
nro_empaque = 206 'Vista Alegre
nro_producto = 1 'Pera
hora_produccion = 10 'Jornada de producci¢n a analizar
bultos = 51 'tipo de hora donde guarda los bultos

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Parametros Fijos "
        Flog.writeline Espacios(Tabulador * 5) & "Empaque 206 (Vista Alegre)"
        Flog.writeline Espacios(Tabulador * 5) & "Producto 1 (Pera)"
        Flog.writeline Espacios(Tabulador * 5) & "Hora Produccion 10 (thnro de Jornada Produccion)"
        Flog.writeline Espacios(Tabulador * 5) & "Bultos 51 (thnro de donde se guardan los bultos)"
    End If

    exito = False
    Encontro1 = False
    Encontro2 = False
    
    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case c_prod_minima_de_peras:
            v_prod_minima_de_peras = rs_wf_tpa!Valor
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Parametros Produccion Minima de Peras: " & v_prod_minima_de_peras
            End If
            Encontro1 = True
        Case c_monto:
            v_monto = rs_wf_tpa!Valor
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Parametros Monto: " & v_monto
            End If
            Encontro2 = True
        End Select
        
        rs_wf_tpa.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not Encontro1 Or Not Encontro2 Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Parametros obligatorios no configurados"
        End If
        Exit Function
    End If
    
    
    'Recorre la Producci¢n Diaria
'DEF BUFFER buf-gti_achdiario FOR gti_achdiario.

'Recorriendo el Desglose Diario
fecha_desde = buliq_periodo!pliqdesde
fecha_hasta = buliq_periodo!pliqhasta

'FGZ - 19/03/2004
If Not CBool(buliq_empleado!empest) Then
    If fecha_hasta > Empleado_Fecha_Fin Then
        fecha_hasta = Empleado_Fecha_Fin
    End If
End If

StrSql = " SELECT estrnro FROM his_estructura " & _
         " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
         " tenro = " & nro_empaque & " AND " & _
         " (htetdesde <= " & ConvFecha(fecha_desde) & ") AND " & _
         " ((" & ConvFecha(fecha_desde) & " <= htethasta) or (htethasta is null))"
OpenRecordset StrSql, rs_Estructura

If rs_Estructura.EOF Then
    'Flog "No se encuentra la sucursal"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "No se encuentra la sucursal. No se ejecuta la formula."
    End If
    Exit Function
Else
    id_empaque = rs_Estructura!estrnro
End If
    
StrSql = "SELECT * FROM gti_achdiario "
StrSql = " INNER JOIN gti_achdiario_estr ON gti_achdiario.achdnro = gti_achdiario_estr.achdnro AND gti_achdiario_estr.tenro = 1 AND gti_achdiario_estr.estrnro = " & nro_empaque  'sucursal o empaque
StrSql = " INNER JOIN  gti_achdiario_estr ON gti_achdiario.achdnro = gti_achdiario_estr.achdnro AND gti_achdiario_estr.tenro = 38 AND gti_achdiario_estr.estrnro = " & nro_producto  'Producto Pera"
StrSql = StrSql & " WHERE (achdfecha =" & ConvFecha(rs_gti_achdiario!achdfecha)
StrSql = StrSql & ") AND (ternro =" & buliq_cabliq!Empleado
StrSql = StrSql & ") AND (thnro  =" & hora_produccion & ")"
OpenRecordset StrSql, rs_gti_achdiario

Do While Not rs_gti_achdiario.EOF
    cant_jornadas = rs_gti_achdiario!achdcanthoras
    productividad_dia = 0

    'Buscar la cantidad de bultos, para el producto y la fecha */
    StrSql = "SELECT * FROM gti_achdiario "
    StrSql = " INNER JOIN gti_achdiario_estr ON gti_achdiario.achdnro = gti_achdiario_estr.achdnro AND gti_achdiario_estr.tenro = 1 AND gti_achdiario_estr.estrnro = " & nro_empaque  'sucursal o empaque
    StrSql = " INNER JOIN  gti_achdiario_estr ON gti_achdiario.achdnro = gti_achdiario_estr.achdnro AND gti_achdiario_estr.tenro = 38 AND gti_achdiario_estr.estrnro = " & nro_producto  'Producto Pera"
    StrSql = StrSql & " WHERE (achdfecha =" & ConvFecha(rs_gti_achdiario!achdfecha)
    StrSql = StrSql & ") AND (ternro =" & buliq_cabliq!Empleado
    StrSql = StrSql & ") AND (thnro  =" & bultos & ")"
    OpenRecordset StrSql, buf_gti_achdiario

    If Not buf_gti_achdiario.EOF Then
        cant_bultos = buf_gti_achdiario!achdcanthoras
    Else
        cant_bultos = 0
    End If
    
    If (cant_bultos > 0) And (cant_bultos > (v_prod_minima_de_peras * cant_jornadas)) Then
       productividad_dia = (cant_bultos - (v_prod_minima_de_peras * cant_jornadas)) * v_monto
       productividad = productividad + productividad_dia
    End If

    If HACE_TRAZA Then
        descripcion = Format(rs_gti_achdiario!achdfecha, "dd/mm/yy")
        descripcion = descripcion + ") Dias: " + Format(cant_jornadas, "0.0")
        descripcion = descripcion + " Bultos: " + Format(cant_bultos, "0000") + " $: "
        Call InsertarTraza(NroCab, Buliq_Concepto(Concepto_Actual).concnro, 0, descripcion, productividad_dia)
        
        descripcion = String(rs_gti_achdiario!achdfecha, "dd/mm/yy") + ") Productividad Acumulada $:"
        Call InsertarTraza(NroCab, Buliq_Concepto(Concepto_Actual).concnro, 0, descripcion, productividad)
    End If

    rs_gti_achdiario.MoveNext
Loop

Monto = productividad
for_203 = productividad
Bien = True
exito = True
    
End Function

Public Function for_204() As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Formula para el calculo de produccion: VA-Peras-Resto
'               Reccorre la produccion dia por dia y compara con el indice con el promedio, aplica el
'               porcentual de dia trabajado y aplica un monto global
' Autor      : FGZ
' Fecha      : 23/12/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim c_indice As Integer
Dim v_indice As Single
Dim c_monto As Integer
Dim v_monto As Single

Dim Encontro1 As Boolean
Dim Encontro2 As Boolean

Dim fecha_desde  As Date
Dim fecha_hasta  As Date
Dim cant_jornadas As Single
Dim id_empaque   As Integer
Dim descripcion As String
Dim productividad_dia As Single
Dim dias_trabajados As Single
Dim nro_empaque As Integer
Dim nro_producto    As Integer
Dim hora_produccion  As Integer
Dim bultos           As Integer
Dim productividad   As Single
Dim cant_bultos  As Single
Dim indice_diario As Single
Dim cant_total_bultos  As Single
Dim cant_embaladores  As Single
Dim aux_ternro As Long


Dim rs_wf_tpa As New ADODB.Recordset
Dim buf_gti_achdiario As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_gti_achdiario As New ADODB.Recordset

' inicializacion de variables
c_indice = 166
c_monto = 51

dias_trabajados = 0
productividad = 0
nro_empaque = 206 'Vista Alegre
nro_producto = 1 'Pera
hora_produccion = 10 'Jornada de producci¢n a analizar
bultos = 51 'tipo de hora donde guarda los bultos


    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Parametros Fijos "
        Flog.writeline Espacios(Tabulador * 5) & "Empaque 206 (Vista Alegre)"
        Flog.writeline Espacios(Tabulador * 5) & "Producto 1 (Pera)"
        Flog.writeline Espacios(Tabulador * 5) & "Hora Produccion 10 (thnro de Jornada Produccion)"
        Flog.writeline Espacios(Tabulador * 5) & "Bultos 51 (thnro de donde se guardan los bultos)"
    End If

    exito = False
    Encontro1 = False
    Encontro2 = False
    
    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case c_indice:
            v_indice = rs_wf_tpa!Valor
            Encontro1 = True
        Case c_monto:
            v_monto = rs_wf_tpa!Valor
            Encontro2 = True
        End Select
        
        rs_wf_tpa.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not Encontro1 Or Not Encontro2 Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encuentró algun parametro "
        End If
        Exit Function
    End If
    
    
'Recorre la Producci¢n Diaria

'Recorriendo el Desglose Diario
fecha_desde = buliq_periodo!pliqdesde
fecha_hasta = buliq_periodo!pliqhasta

'FGZ - 19/03/2004
If Not CBool(buliq_empleado!empest) Then
    If fecha_hasta > Empleado_Fecha_Fin Then
        fecha_hasta = Empleado_Fecha_Fin
    End If
End If

StrSql = " SELECT estrnro FROM his_estructura " & _
         " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
         " tenro = " & nro_empaque & " AND " & _
         " (htetdesde <= " & ConvFecha(fecha_desde) & ") AND " & _
         " ((" & ConvFecha(fecha_desde) & " <= htethasta) or (htethasta is null))"
OpenRecordset StrSql, rs_Estructura

If rs_Estructura.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "No se encuentra la sucursal"
    End If
    Exit Function
Else
    id_empaque = rs_Estructura!estrnro
End If


StrSql = "SELECT * FROM gti_achdiario "
StrSql = " INNER JOIN gti_achdiario_estr ON gti_achdiario.achdnro = gti_achdiario_estr.achdnro AND gti_achdiario_estr.tenro = 1 AND gti_achdiario_estr.estrnro = " & nro_empaque  'sucursal o empaque
StrSql = " INNER JOIN  gti_achdiario_estr ON gti_achdiario.achdnro = gti_achdiario_estr.achdnro AND gti_achdiario_estr.tenro = 38 AND gti_achdiario_estr.estrnro = " & nro_producto  'Producto Pera"
StrSql = StrSql & " WHERE (achdfecha =" & ConvFecha(rs_gti_achdiario!achdfecha)
StrSql = StrSql & ") AND (ternro =" & buliq_cabliq!Empleado
StrSql = StrSql & ") AND (thnro  =" & hora_produccion & ")"
OpenRecordset StrSql, rs_gti_achdiario

Do While Not rs_gti_achdiario.EOF
    cant_jornadas = rs_gti_achdiario!achdcanthoras
    productividad_dia = 0

    'Buscar el Indice Diario (promedio de bultos por cantidad de embaladores, para el producto y la fecha)
    cant_embaladores = 0
    cant_total_bultos = 0
    indice_diario = 0
    
    StrSql = "SELECT * FROM gti_achdiario "
    StrSql = " INNER JOIN gti_achdiario_estr ON gti_achdiario.achdnro = gti_achdiario_estr.achdnro AND gti_achdiario_estr.tenro = 1 AND gti_achdiario_estr.estrnro = " & nro_empaque  'sucursal o empaque
    StrSql = " INNER JOIN  gti_achdiario_estr ON gti_achdiario.achdnro = gti_achdiario_estr.achdnro AND gti_achdiario_estr.tenro = 38 AND gti_achdiario_estr.estrnro = " & nro_producto  'Producto Pera"
    StrSql = StrSql & " WHERE (achdfecha =" & ConvFecha(rs_gti_achdiario!achdfecha)
    StrSql = StrSql & ") AND (thnro  =" & bultos & ")"
    OpenRecordset StrSql, rs_gti_achdiario
    
    If Not buf_gti_achdiario.EOF Then
        aux_ternro = buf_gti_achdiario!ternro
    
        Do While Not buf_gti_achdiario.EOF
            ' si cambia el tercero ==> es otro embalador
            If aux_ternro <> buf_gti_achdiario!ternro Then
                cant_embaladores = cant_embaladores + 1
            End If
            
            cant_total_bultos = cant_total_bultos + buf_gti_achdiario!achdcanthoras
            
            buf_gti_achdiario.MoveNext
        Loop
        indice_diario = cant_total_bultos / cant_embaladores
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Cantidad de bultos " & cant_total_bultos
            Flog.writeline Espacios(Tabulador * 4) & "Cantidad de embaladores " & cant_embaladores
            Flog.writeline Espacios(Tabulador * 4) & "Indice Diario " & indice_diario
        End If
    End If
    
    If HACE_TRAZA Then
        descripcion = Format(rs_gti_achdiario!achdfecha, "dd/mm/yy")
        descripcion = descripcion + " Bultos: " + Format(cant_total_bultos, "000000") + " "
        descripcion = descripcion + " Emb: " + Format(cant_embaladores, "00000") + " Ind: "
        Call InsertarTraza(NroCab, Buliq_Concepto(Concepto_Actual).concnro, 0, descripcion, indice_diario)
    End If


    'Opcion para el Indice Diario: Simplifica, es mas rapido y evita minimos errores
    '                              individuales de bultos que pueden afectar a todos.

    ' Crear 31 parametros de Novedad Global (no obligatorios, si no estan, ser cero,
    ' y que se cargue manualmente para cada dia el promedio.
    ' Luego se pueden cargar los valores en un arreglo de 1 a 31 y entrar en la posiscion
    ' de acuerdo a la fecha que se esta analizando.
    

    If indice_diario > (v_indice * cant_jornadas) Then
        indice_diario = indice_diario - v_indice
        productividad_dia = (cant_jornadas * indice_diario * v_monto)
        productividad = productividad + productividad_dia
   
        If HACE_TRAZA Then
            descripcion = Format(rs_gti_achdiario!achdfecha, "dd/mm/yy")
            descripcion = descripcion + " Día: " + Format(cant_jornadas, "0.0") + " "
            descripcion = descripcion + " Día: " + Format(productividad_dia, "0000.00") + " $"
            Call InsertarTraza(NroCab, Buliq_Concepto(Concepto_Actual).concnro, 0, descripcion, productividad)
        End If
    End If
    
    rs_gti_achdiario.MoveNext
Loop

Monto = productividad
for_204 = productividad
Bien = True
exito = True
    
End Function


Public Function for_207() As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Formula para el calculo de produccion: VA-Peras-Resto
'               Reccorre la produccion dia por dia y compara con el indice con el promedio, aplica el
'               porcentual de dia trabajado y aplica un monto global
' Autor      : FGZ
' Fecha      : 29/12/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim c_indice As Integer
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
Dim c_monto As Integer

Dim v_indice As Single
Dim v_indice_dia(31) As Single

Dim v_indice_dia_1 As Single
Dim v_indice_dia_2 As Single
Dim v_indice_dia_3 As Single
Dim v_indice_dia_4 As Single
Dim v_indice_dia_5 As Single
Dim v_indice_dia_6 As Single
Dim v_indice_dia_7 As Single
Dim v_indice_dia_8 As Single
Dim v_indice_dia_9 As Single
Dim v_indice_dia_10 As Single
Dim v_indice_dia_11 As Single
Dim v_indice_dia_12 As Single
Dim v_indice_dia_13 As Single
Dim v_indice_dia_14 As Single
Dim v_indice_dia_15 As Single
Dim v_indice_dia_16 As Single
Dim v_indice_dia_17 As Single
Dim v_indice_dia_18 As Single
Dim v_indice_dia_19 As Single
Dim v_indice_dia_20 As Single
Dim v_indice_dia_21 As Single
Dim v_indice_dia_22 As Single
Dim v_indice_dia_23 As Single
Dim v_indice_dia_24 As Single
Dim v_indice_dia_25 As Single
Dim v_indice_dia_26 As Single
Dim v_indice_dia_27 As Single
Dim v_indice_dia_28 As Single
Dim v_indice_dia_29 As Single
Dim v_indice_dia_30 As Single
Dim v_indice_dia_31 As Single
Dim v_monto As Single

Dim Encontro1 As Boolean
Dim Encontro2 As Boolean

Dim fecha_desde  As Date
Dim fecha_hasta  As Date
Dim cant_jornadas As Single
Dim id_empaque   As Integer
Dim descripcion As String
Dim productividad_dia As Single
Dim supera As Single
Dim dias_trabajados As Single
Dim nro_empaque As Integer
Dim nro_producto    As Integer
Dim hora_produccion  As Integer
Dim bultos           As Integer
Dim productividad   As Single
Dim cant_bultos  As Single
Dim indice_diario As Single
Dim cant_total_bultos  As Single
Dim cant_embaladores  As Single

Dim rs_wf_tpa As New ADODB.Recordset
Dim buf_gti_achdiario As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_gti_achdiario As New ADODB.Recordset

' inicializacion de variables
c_indice = 166
c_monto = 51

c_indice_dia_1 = 168
c_indice_dia_2 = 169
c_indice_dia_3 = 170
c_indice_dia_4 = 171
c_indice_dia_5 = 172
c_indice_dia_6 = 173
c_indice_dia_7 = 174
c_indice_dia_8 = 175
c_indice_dia_9 = 176
c_indice_dia_10 = 177
c_indice_dia_11 = 178
c_indice_dia_12 = 179
c_indice_dia_13 = 180
c_indice_dia_14 = 181
c_indice_dia_15 = 182
c_indice_dia_16 = 183
c_indice_dia_17 = 184
c_indice_dia_18 = 185
c_indice_dia_19 = 186
c_indice_dia_20 = 187
c_indice_dia_21 = 188
c_indice_dia_22 = 189
c_indice_dia_23 = 190
c_indice_dia_24 = 191
c_indice_dia_25 = 192
c_indice_dia_26 = 193
c_indice_dia_27 = 194
c_indice_dia_28 = 195
c_indice_dia_29 = 196
c_indice_dia_30 = 197
c_indice_dia_31 = 198

v_indice_dia_1 = 0
v_indice_dia_2 = 0
v_indice_dia_3 = 0
v_indice_dia_4 = 0
v_indice_dia_5 = 0
v_indice_dia_6 = 0
v_indice_dia_7 = 0
v_indice_dia_8 = 0
v_indice_dia_9 = 0
v_indice_dia_10 = 0
v_indice_dia_11 = 0
v_indice_dia_12 = 0
v_indice_dia_13 = 0
v_indice_dia_14 = 0
v_indice_dia_15 = 0
v_indice_dia_16 = 0
v_indice_dia_17 = 0
v_indice_dia_18 = 0
v_indice_dia_19 = 0
v_indice_dia_20 = 0
v_indice_dia_21 = 0
v_indice_dia_22 = 0
v_indice_dia_23 = 0
v_indice_dia_24 = 0
v_indice_dia_25 = 0
v_indice_dia_26 = 0
v_indice_dia_27 = 0
v_indice_dia_28 = 0
v_indice_dia_29 = 0
v_indice_dia_30 = 0
v_indice_dia_31 = 0

dias_trabajados = 0
productividad = 0
nro_empaque = 2 'Vista Alegre
nro_producto = 1 'Pera
hora_produccion = 10 'Jornada de producci¢n a analizar
bultos = 51 'tipo de hora donde guarda los bultos

    Bien = False
    exito = False
    Encontro1 = False
    Encontro2 = False

    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case c_monto:
            v_monto = rs_wf_tpa!Valor
            Encontro2 = True
        Case c_indice:
            v_indice = rs_wf_tpa!Valor
            Encontro1 = True
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
        End Select
        
        rs_wf_tpa.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not Encontro1 Or Not Encontro2 Then
        Exit Function
    End If
    
    
'Recorre la Producci¢n Diaria

'Recorriendo el Desglose Diario
fecha_desde = buliq_periodo!pliqdesde
fecha_hasta = buliq_periodo!pliqhasta

'FGZ - 19/03/2004
If Not CBool(buliq_empleado!empest) Then
    If fecha_hasta > Empleado_Fecha_Fin Then
        fecha_hasta = Empleado_Fecha_Fin
    End If
End If

StrSql = " SELECT estrnro FROM his_estructura " & _
         " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
         " tenro = " & nro_empaque & " AND " & _
         " (htetdesde <= " & ConvFecha(fecha_desde) & ") AND " & _
         " ((" & ConvFecha(fecha_desde) & " <= htethasta) or (htethasta is null))"
OpenRecordset StrSql, rs_Estructura

If rs_Estructura.EOF Then
    'Flog "No se encuentra la sucursal"
    Exit Function
Else
    id_empaque = rs_Estructura!estrnro
End If


StrSql = "SELECT * FROM gti_achdiario WHERE (" & _
         ConvFecha(fecha_desde) & "<= achfecha " & _
         " AND achdfecha <=" & ConvFecha(fecha_hasta) & _
         ") AND (puenro = " & nro_producto & _
         ") AND (sucursal = " & id_empaque & _
         ") AND (ternro = " & buliq_cabliq!Empleado & _
         ") AND (gti_achdiario.thnro  = " & hora_produccion & ")"
OpenRecordset StrSql, rs_gti_achdiario

Do While Not rs_gti_achdiario.EOF
    cant_jornadas = rs_gti_achdiario!achdcanthoras
    productividad_dia = 0

    'Buscar el Indice Diario (promedio de bultos por cantidad de embaladores, para el producto y la fecha)
    cant_embaladores = 0
    cant_total_bultos = 0
    
    supera = v_indice_dia(Day(rs_gti_achdiario!achdfecha)) - v_indice
    productividad_dia = (v_indice_dia(Day(rs_gti_achdiario!achdfecha)) * cant_jornadas * v_monto)
    productividad = productividad + productividad_dia

    
    If HACE_TRAZA Then
        descripcion = Format(rs_gti_achdiario!achdfecha, "dd/mm/yy")
        descripcion = descripcion + " Día: " + Format(cant_jornadas, "0.0") + " "
        descripcion = descripcion + " Ind: " + Format(v_indice_dia(Day(rs_gti_achdiario!achdfecha)), "0000.00") + " $"
        Call InsertarTraza(NroCab, Buliq_Concepto(Concepto_Actual).concnro, 0, descripcion, productividad_dia)
    End If
    
    rs_gti_achdiario.MoveNext
Loop

Monto = productividad
for_207 = productividad
Bien = True
exito = True
    
End Function


Public Function For_Prom_Produccion()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Promedio de Prod.
'              segun Dias Trabajados y Bultos Producidos
' Autor      : FGZ
' Fecha      : 05/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim c_Trabajados As Integer
Dim v_Trabajados As Single
Dim c_Total_de_Bultos As Integer
Dim v_Total_de_Bultos As Single

Dim Encontro1 As Boolean
Dim Encontro2 As Boolean

Dim rs_wf_tpa As New ADODB.Recordset

' inicializacion de variables
c_Trabajados = 2
c_Total_de_Bultos = 163

    Bien = False
    Encontro1 = False
    Encontro2 = False

    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa

    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case c_Trabajados:
            v_Trabajados = rs_wf_tpa!Valor
            Encontro1 = True
        Case c_Total_de_Bultos:
            v_Total_de_Bultos = rs_wf_tpa!Valor
            Encontro2 = True
        End Select

        rs_wf_tpa.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not Encontro1 Or Not Encontro2 Then
        Exit Function
    End If

Bien = True
For_Prom_Produccion = v_Total_de_Bultos / v_Trabajados
exito = True

End Function
