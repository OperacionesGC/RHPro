Attribute VB_Name = "Mdlformulas"
Option Explicit

Public Type TTraza_Gan
    PliqNro As Long
    concnro As Long
    Empresa As Long
    Fecha_pago As Date
    ternro As Long
    Msr As Double
    Nomsr As Double
    Nogan As Double
    Jubilacion As Double
    Osocial As Double
    Cuota_medico As Double
    Prima_seguro As Double
    Sepelio As Double
    Estimados As Double
    Otras As Double
    Donacion As Double
    Dedesp As Double
    Noimpo As Double
    Car_flia As Double
    conyuge As Long
    Hijo   As Long
    Otras_cargas   As Long
    Retenciones As Double
    Promo As Double
    Saldo As Double
    Sindicato As Double
    Ret_Mes As Double
    Mon_conyuge As Double
    Mon_hijo As Double
    Mon_otras As Double
    Viaticos As Double
    Amortizacion As Double
    Entidad1 As String
    Entidad2 As String
    Entidad3 As String
    Entidad4 As String
    Entidad5 As String
    Entidad6 As String
    Entidad7 As String
    Entidad8 As String
    Entidad9 As String
    Entidad10 As String
    Entidad11 As String
    Entidad12 As String
    Entidad13 As String
    Entidad14 As String
    Cuit_entidad1 As String
    Cuit_entidad2 As String
    Cuit_entidad3 As String
    Cuit_entidad4 As String
    Cuit_entidad5 As String
    Cuit_entidad6 As String
    Cuit_entidad7 As String
    Cuit_entidad8 As String
    Cuit_entidad9 As String
    Cuit_entidad10 As String
    Cuit_entidad11 As String
    Cuit_entidad12 As String
    Cuit_entidad13 As String
    Cuit_entidad14 As String
    Monto_entidad1 As Double
    Monto_entidad2 As Double
    Monto_entidad3 As Double
    Monto_entidad4 As Double
    Monto_entidad5 As Double
    Monto_entidad6 As Double
    Monto_entidad7 As Double
    Monto_entidad8 As Double
    Monto_entidad9 As Double
    Monto_entidad10 As Double
    Monto_entidad11 As Double
    Monto_entidad12 As Double
    Monto_entidad13 As Double
    Monto_entidad14 As Double
    Ganimpo As Double
    Ganneta As Double
    Total_entidad1 As Double
    Total_entidad2 As Double
    Total_entidad3 As Double
    Total_entidad4 As Double
    Total_entidad5 As Double
    Total_entidad6 As Double
    Total_entidad7 As Double
    Total_entidad8 As Double
    Total_entidad9 As Double
    Total_entidad10 As Double
    Total_entidad11 As Double
    Total_entidad12 As Double
    Total_entidad13 As Double
    Total_entidad14 As Double
    pronro As Long
    Imp_deter As Double
    Eme_medicas As Double
    Seguro_optativo As Double
    Seguro_retiro As Double
    Tope_os_priv As Double
    Empleg As Long
End Type

Public Type TTraza_Gan_Item_Top
    Itenro  As Long
    ternro  As Long
    pronro  As Long
    Empresa As Long
    Monto As Double
    Ddjj As Double
    Old_liq As Double
    Liq As Double
    Prorr As Double
    'Pror As Double
End Type


Public Type TDesliq
    Itenro As Long
    Empleado As Long
    Dlfecha As Date
    pronro As Long
    Dlmonto As Double
    Dlprorratea As Boolean
End Type

Public Type TFicharet
    Fecha As Date
    importe As Double
    pronro As Long
    liqsistema As Boolean
    Empleado As Long
End Type


Public Function EjecutarFormulaNoConfigurable(ByVal nombre As String) As Double
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
' Inicio Formulas de Glencore
    Case "FOR_PROVSAC":
        EjecutarFormulaNoConfigurable = for_ProvSac(NroCab, fec, Monto, Bien)
    Case "FOR_PROVVAC":
        EjecutarFormulaNoConfigurable = for_ProvVac(NroCab, fec, Monto, Bien)
    Case "FOR_DPROVSAC":
        EjecutarFormulaNoConfigurable = for_DesaProvSac(NroCab, fec, Monto, Bien)
    Case "FOR_DPROVVAC":
        EjecutarFormulaNoConfigurable = for_DesaProvVac(NroCab, fec, Monto, Bien)
    Case "FOR_PORCPRES":
        EjecutarFormulaNoConfigurable = for_PorcPres(NroCab, fec, Monto, Bien)
' Fin Formulas de Glencore
'-------------------------------------------------
    Case Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Formula inexistente " & UCase(nombre)
        End If
    End Select
End Function

Public Function EjecutarFormulaDeSistema(ByVal nombre As String) As Double

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
    Case "FOR_BASEACCI":
        EjecutarFormulaDeSistema = For_Baseacci
    Case "FOR_NIVELAR":
        EjecutarFormulaDeSistema = for_Nivelar(fec)
    Case "FOR_IRP":
        EjecutarFormulaDeSistema = for_irp(NroCab, fec, Monto, Bien)
    Case "FOR_IRP_FRANJA":
        EjecutarFormulaDeSistema = for_irp_franja(NroCab, fec, Monto, Bien)
    Case "FOR_RIBETEADO":
        EjecutarFormulaDeSistema = for_Ribeteado
    Case "FOR_IRPF":
        EjecutarFormulaDeSistema = for_irpf(NroCab, fec, Monto, Bien)
    Case "FOR_IRPF_SIMPLE":
        EjecutarFormulaDeSistema = for_irpf_simple(NroCab, fec, Monto, Bien)
    Case "FOR_IRPF_DICIEMBRE":
        EjecutarFormulaDeSistema = for_irpf_diciembre(NroCab, fec, Monto, Bien)
    Case "FOR_GROSSNEW":
        EjecutarFormulaDeSistema = for_grossnew
    Case "FOR_GANANCIAS_PETROLEROS":
        EjecutarFormulaDeSistema = for_Ganancias_Petroleros(NroCab, fec, Monto, Bien)
    Case "FOR_IMPUESTOUNICO":
        EjecutarFormulaDeSistema = for_ImpuestoUnico(NroCab, fec, Monto, Bien)
    Case "FOR_RECALCCONCEPTO":
        EjecutarFormulaDeSistema = for_RecalcConcepto(NroCab, fec, Monto, Bien)
    Case "FOR_RECALCIMPUESTOUNICO":
        EjecutarFormulaDeSistema = for_RecalcImpuestoUnico(NroCab, fec, Monto, Bien)
    Case "FOR_SAC_NO_REMU":
        EjecutarFormulaDeSistema = For_Sac_No_Remu
    Case Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "Formula inexistente " & UCase(nombre)
        End If
    End Select
   
End Function

' ---------------------------------------------------------
' Modulo de fórmulas conocidas
' ---------------------------------------------------------

Public Function for_neg() As Double
' Monto Negativo
Dim v_monto As Double
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


Public Function for_pos() As Double
' Monto Positivo
Dim v_monto As Double
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

Public Function for_grossnew() As Double
' Grossing nuevo
' Maximiliano Breglia 24/11/2006

Dim v_monto As Double
Dim v_diferencia As Double
Dim rs_wf_tpa As New ADODB.Recordset
Dim Encontro_netoactual As Boolean
Dim Encontro_netopactado As Boolean
Dim Encontro_error As Boolean
Dim Encontro_conccod As Boolean
Dim Encontro_tpanro As Boolean
Dim v_netoactual As Double
Dim v_netopactado As Double
Dim v_error As Double
Dim v_conccod As Double
Dim v_concnro As Double
Dim v_tpanro As Double
Const c_netoactual = 63
Const c_netopactado = 115
Const c_error = 4
Const c_conccod = 7
Const c_tpanro = 51

Dim rs_Conc As New ADODB.Recordset

    
    
    exito = False
    Encontro_netoactual = False
    Encontro_netopactado = False
    Encontro_error = False
    Encontro_conccod = False
    Encontro_tpanro = False
    v_concnro = 0
    
    'Control para no pasar al sgt concepto de grossing hasta que el actual converga
    If (v_nroConce = Arr_conceptos(Concepto_Actual).concnro) Or (v_nroConce = 0) Then
        
        Usa_grossing = False
        Termino_gross = True
        
        StrSql = "SELECT * FROM " & TTempWF_tpa
        OpenRecordset StrSql, rs_wf_tpa
        
        Do While Not rs_wf_tpa.EOF
            Select Case rs_wf_tpa!tipoparam
            Case c_netoactual:
                v_netoactual = rs_wf_tpa!Valor
                Encontro_netoactual = True
            Case c_netopactado:
                v_netopactado = rs_wf_tpa!Valor
                Encontro_netopactado = True
            Case c_error:
                v_error = rs_wf_tpa!Valor
                Encontro_error = True
            Case c_conccod:
                v_conccod = rs_wf_tpa!Valor
                Encontro_conccod = True
            Case Else
            End Select
            
            rs_wf_tpa.MoveNext
        Loop
    
        ' si no se obtuvieron los parametros, ==> Error.
        If Not Encontro_conccod Or Not Encontro_netoactual Or Not Encontro_netopactado Or Not Encontro_error Then
            Exit Function
        End If
        
        ' Buscar concepto VER COMO SACAR EL CAST
         StrSql = "SELECT concnro FROM concepto " & _
                     " WHERE CAST(conccod as int) = " & v_conccod
            OpenRecordset StrSql, rs_Conc
           
        If Not rs_Conc.EOF Then
                v_concnro = rs_Conc!concnro
              Else
                Exit Function
        End If
        
        If v_nroitera = 1 Then
           v_netofijo = v_netopactado
           v_nroConce = Arr_conceptos(Concepto_Actual).concnro
        End If
        
         If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "---------Formula Grossing------------------"
            Flog.writeline Espacios(Tabulador * 3) & "Iteracion " & v_nroitera
            Flog.writeline Espacios(Tabulador * 3) & "Neto pactado " & v_netopactado
            Flog.writeline Espacios(Tabulador * 3) & "Neto Actual " & v_netoactual
        End If
        
        ' Comparo con el error
        If (Abs((v_netofijo - v_netoactual)) < v_error) Or v_nroitera > MaxIteraGross Then
               'Converge la formula de Grossing
               Usa_grossing = False
               Termino_gross = True
               
               'Para pasar al siguiente concepto
               v_nroConce = 0
               v_nroitera = 1
    
        Else
           v_diferencia = v_netofijo - v_netoactual
           ' Inserto novedad
            StrSql = " SELECT * FROM novemp " & _
                     " WHERE concnro =" & v_concnro & _
                     " AND tpanro =" & c_tpanro & _
                     " AND empleado =" & buliq_empleado!ternro & _
                     " AND ((nevigencia = -1 " & _
                     " AND nedesde <= " & ConvFecha(Fecha_Fin) & _
                     " AND (nehasta >= " & ConvFecha(Fecha_Inicio) & _
                     " OR nehasta is null )) " & _
                     " OR nevigencia = 0)"
            OpenRecordset StrSql, rs_Conc
            
            If rs_Conc.EOF Then
               ' Inserto novedad
                 StrSql = "INSERT INTO novemp (empleado, concnro, tpanro, nevalor,nevigencia,nedesde,nehasta,pronro ) VALUES (" & _
                 buliq_empleado!ternro & "," & v_concnro & "," & c_tpanro & "," & v_diferencia & ", -1" & _
                 "," & ConvFecha(Fecha_Inicio) & _
                 "," & ConvFecha(Fecha_Fin) & _
                 "," & buliq_proceso!pronro & _
                 " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Else ' Actualizo la novedad
                 StrSql = "UPDATE novemp SET nevalor = nevalor + " & v_diferencia & _
                          " WHERE concnro = " & v_concnro & _
                          " AND tpanro = " & c_tpanro & _
                          " AND empleado = " & buliq_empleado!ternro & _
                          " AND ((nevigencia = -1 " & _
                          " AND nedesde <= " & ConvFecha(Fecha_Fin) & _
                          " AND (nehasta >= " & ConvFecha(Fecha_Inicio) & _
                          " OR nehasta is null )) " & _
                          " OR nevigencia = 0)"
         
                 objConn.Execute StrSql, , adExecuteNoRecords
            
            End If
            Usa_grossing = True
            Termino_gross = False
            
        End If
                                       
       If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Diferencia " & v_diferencia
        Flog.writeline Espacios(Tabulador * 3) & "Termino " & Termino_gross
       End If
       
       exito = True
       for_grossnew = v_diferencia
       
       If v_nroitera > MaxIteraGross Then
          exito = False
       End If
            
    End If '(v_nroConce = Arr_conceptos(Concepto_Actual).concnro) Or (v_nroConce = 0)
    
End Function

Public Function for_porc_neg() As Double
' Porcentaje Negativo
Const c_porcentaje = 35
Const c_msr = 8

Dim v_porcentaje As Double
Dim v_msr As Double

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


Public Function for_porc_pos() As Double
' Porcentaje Positivo
Const c_porcentaje = 35
Const c_msr = 8

Dim v_porcentaje As Double
Dim v_msr As Double

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


Public Function for_Ganancias(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias.
' Autor      :
' Fecha      :
' Ultima Mod.: 22/07/2005
' Descripcion: se agregó el item 30 (movilidad) y se computa su valor en traza_gan.viaticos.
' Ultima Mod.: D.S. 02/11/2005
' Descripcion: Se agregaron 3 campos nuevos a traza_gan que estan relacionados con el F649.
'               traza_gan.deducciones decimal(19,4)
'               traza_gan.art23 decimal(19,4)
'               traza_gan.porcdeduc decimal(19,4)
' ---------------------------------------------------------------------------------------------
Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim p_prorratea As Integer  'Si prorratea o no para liq. finales

'Variables Locales
Dim Devuelve As Double
Dim Tope_Gral As Double
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
Dim Items_DDJJ(100) As Double
Dim Items_LIQ(100) As Double
Dim Items_PRORR(100) As Double
Dim Items_OLD_LIQ(100) As Double
Dim Items_TOPE(100) As Double
Dim Items_ART_23(100) As Boolean

'Recorsets Auxiliares
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Item As New ADODB.Recordset
Dim rs_valitem As New ADODB.Recordset
Dim rs_Desmen As New ADODB.Recordset
Dim rs_Desliq As New ADODB.Recordset
Dim rs_itemacum As New ADODB.Recordset
Dim rs_itemconc As New ADODB.Recordset
Dim rs_escala_ded As New ADODB.Recordset
Dim rs_escala As New ADODB.Recordset
Dim rs_Ficharet As New ADODB.Recordset
Dim rs_Traza_gan_items_tope As New ADODB.Recordset

Dim Hasta As Integer

' FGZ - 12/02/2004
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Double
Dim val_impdebitos As Double
Dim fechaFichaH As Date
Dim fechaFichaD As Date


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
Dim Total_Empresa As Double
Dim Tope As Integer
'Dim rs_Rep19 As New ADODB.Recordset
Dim rs_Traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Double
Total_Empresa = 0
Tope = 10

Descuentos = 0
' Primero limpio la traza
StrSql = "DELETE FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
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
         buliq_empleado!Empleg & _
         ")"
objConn.Execute StrSql, , adExecuteNoRecords

'FGZ - 22/12/2004
'Leo la tabla
StrSql = "SELECT * FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
OpenRecordset StrSql, rs_Traza_gan
    

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
ini_anyo_ret = C_Date("01/01/" & Ret_Ano)
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
    Ret_Mes = 12
    'FGZ - 27/09/2004
    fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
End If

If Neto < 0 Then
   If CBool(USA_DEBUG) Then
      Flog.writeline Espacios(Tabulador * 3) & "El Neto del mes es negativo, se setea en cero."
   End If
   If HACE_TRAZA Then
      Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Neto, "El Neto del Mes es negativo, se seteara en cero.", Neto)
   End If
   Neto = 0
End If

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 3) & "---------Formula-----------------------------"
    Flog.writeline Espacios(Tabulador * 3) & "Fecha del Proceso " & buliq_proceso!profecpago
    Flog.writeline Espacios(Tabulador * 3) & "Mes de Retencion " & Ret_Mes
    Flog.writeline Espacios(Tabulador * 3) & "Año de Retencion " & Ret_Ano
    Flog.writeline Espacios(Tabulador * 3) & "Fin mes de Retencion " & fin_mes_ret
    
    Flog.writeline Espacios(Tabulador * 3) & "Máxima Ret. en % " & Tope_Gral
    Flog.writeline Espacios(Tabulador * 3) & "Neto del Mes " & Neto
End If
If HACE_TRAZA Then
'    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Tope_Gral, "Máxima Ret. en %", Tope_Gral)
'    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Neto, "Neto del Mes", Neto)
'FGZ - 03/06/2006
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 99999, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 99999, "Neto del Mes", Neto)
End If


'Limpiar items que suman al articulo 23
For I = 1 To 100
    Items_ART_23(I) = False
Next I
val_impdebitos = 0


' Recorro todos los items de Ganancias
StrSql = "SELECT * FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_Item

Do While Not rs_Item.EOF
  
  ' Impuestos y debitos Bancarios va como Promocion
  If (rs_Item!Itenro = 29 Or rs_Item!Itenro = 55) And Ret_Mes = 12 Then
        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!ternro & _
                 " AND desano=" & Ret_Ano & _
                 " AND itenro = " & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
        If Not rs_Desmen.EOF Then
           If rs_Item!Itenro = 29 Then
             val_impdebitos = rs_Desmen!desmondec * 0.34
           Else
            val_impdebitos = rs_Desmen!desmondec * 0.17
           End If
        End If
        rs_Desmen.Close
  Else
    
    Select Case rs_Item!itetipotope
    Case 1: ' el valor a tomar es lo que dice la escala
    
        StrSql = "SELECT * FROM valitem WHERE viano =" & Ret_Ano & _
                 " AND itenro=" & rs_Item!Itenro & _
                 " AND vimes =" & Ret_Mes
        OpenRecordset StrSql, rs_valitem
        
        Do While Not rs_valitem.EOF
            Items_DDJJ(rs_valitem!Itenro) = rs_valitem!vimonto
            Items_TOPE(rs_valitem!Itenro) = rs_valitem!vimonto
            
            rs_valitem.MoveNext
        Loop
    ' End case 1
    ' ------------------------------------------------------------------------
    
    Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
        ' Busco la declaracion jurada
        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!ternro & _
                 " AND desano=" & Ret_Ano & _
                 " AND itenro = " & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
        
        Do While Not rs_Desmen.EOF
            If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                If rs_Item!Itenro = 3 Then
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                    
                    'FGZ - 19/04/2004
                    If rs_Item!Itenro <= 4 Then
                        If Not EsNulo(rs_Desmen!descuit) Then
                            I = 11
                            Distinto = rs_Traza_gan!Cuit_entidad11 <> rs_Desmen!descuit
                            Do While (I <= Tope) And Distinto
                                I = I + 1
                                Select Case I
                                Case 11:
                                    Distinto = rs_Traza_gan!Cuit_entidad11 <> rs_Desmen!descuit
                                Case 12:
                                    Distinto = rs_Traza_gan!Cuit_entidad12 <> rs_Desmen!descuit
                                Case 13:
                                    Distinto = rs_Traza_gan!Cuit_entidad13 <> rs_Desmen!descuit
                                Case 14:
                                    Distinto = rs_Traza_gan!Cuit_entidad14 <> rs_Desmen!descuit
                                End Select
                            Loop
                          
                            If I > Tope And I <= 14 Then
                                StrSql = "UPDATE traza_gan SET "
                                StrSql = StrSql & " cuit_entidad" & I & "='" & rs_Desmen!descuit & "',"
                                StrSql = StrSql & " entidad" & I & "='" & rs_Desmen!DesRazsoc & "',"
                                StrSql = StrSql & " monto_entidad" & I & "=" & rs_Desmen!desmondec
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
                                'StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                OpenRecordset StrSql, rs_Traza_gan
                                
                                
                                Tope = Tope + 1
                            Else
                                If I = 15 Then
                                    Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!Empleg ' empleado.empleg
                                Else
                                    StrSql = "UPDATE traza_gan SET "
                                    StrSql = StrSql & " monto_entidad" & I & "= monto_entidad" & I & " + " & rs_Desmen!desmondec
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
                                    'StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                    OpenRecordset StrSql, rs_Traza_gan
                                End If
                            End If
                        Else
                            Total_Empresa = Total_Empresa + rs_Desmen!desmondec
                        End If
                    End If
                    'FGZ - 19/04/2004
                    
                Else
                    If rs_Desmen!desmenprorra = 0 Then 'no es parejito
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + rs_Desmen!desmondec
                    Else
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                    End If
                End If
            End If
            
            
            rs_Desmen.MoveNext
        Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_Desliq
        If rs_Desliq.EOF Then
            If CBool(USA_DEBUG) Then
                Flog.writeline "No hay datos de retenciones anteriores"
            End If
        End If
        Do While Not rs_Desliq.EOF
            Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
            'Si el desliq prorratea debo proporcionarlo
            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)

            rs_Desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_Item!Itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acuNro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itaprorratea) Then
                    If CBool(rs_itemacum!itasigno) Then
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                    Else
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                    End If
                Else
                    If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                    Else
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - IIf((rs_Item!Itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
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
                 " AND itemconc.itenro =" & rs_Item!Itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) Then
                If CBool(rs_itemconc!itcsigno) Then
                    Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                Else
                    Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                End If
            Else
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + rs_itemconc!dlimonto
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                Else
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - rs_itemconc!dlimonto
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - IIf((rs_Item!Itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
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
     
        StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_Ano & _
                 " AND vimes = " & Ret_Mes & _
                 " AND itenro =" & rs_Item!Itenro
        OpenRecordset StrSql, rs_valitem
         Do While Not rs_valitem.EOF
            Items_TOPE(rs_Item!Itenro) = rs_valitem!vimonto
         
            rs_valitem.MoveNext
         Loop
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_Ano & _
                 " AND itenro =" & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
         Do While Not rs_Desmen.EOF
            If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                If rs_Desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                Else
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                End If
            End If
         
            rs_Desmen.MoveNext
         Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_Desliq

        Do While Not rs_Desliq.EOF
            Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto

            rs_Desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
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
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
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
        
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            Items_TOPE(rs_Item!Itenro) = Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_Item!itesigno) Then
            Items_TOPE(rs_Item!Itenro) = -Items_TOPE(rs_Item!Itenro)
        End If
        
    ' End case 3
    ' ------------------------------------------------------------------------
    Case 4:
        ' Tomo los valores de la DDJJ y el valor de la escala (cargas de familia)
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_Ano & _
                 " AND itenro =" & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
         Do While Not rs_Desmen.EOF
            If Month(rs_Desmen!desfechas) <= Ret_Mes Then
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec * (Month(rs_Desmen!desfechas) - Month(rs_Desmen!desfecdes) + 1)
            Else
                If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1)
                End If
            End If
        
            rs_Desmen.MoveNext
         Loop
        
        If Items_DDJJ(rs_Item!Itenro) > 0 Then
            StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_Ano & _
                     " AND vimes = " & Ret_Mes & _
                     " AND itenro =" & rs_Item!Itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_Item!Itenro) = rs_valitem!vimonto / Ret_Mes * Items_DDJJ(rs_Item!Itenro)
             
                rs_valitem.MoveNext
             Loop
        End If
    ' End case 4
    ' ------------------------------------------------------------------------
        
    Case 5:
        I = 1
        j = 1
        'Hasta = IIf(50 > Len(rs_item!iteitemstope), 50, rs_item!iteitemstope)
        Hasta = 100
        Terminar = False
        Do While j <= Hasta And Not Terminar
            pos1 = I
            pos2 = InStr(I, rs_Item!iteitemstope, ",") - 1
            If pos2 > 0 Then
                Texto = Mid(rs_Item!iteitemstope, pos1, pos2 - pos1 + 1)
            Else
                pos2 = Len(rs_Item!iteitemstope)
                Texto = Mid(rs_Item!iteitemstope, pos1, pos2 - pos1 + 1)
                Terminar = True
            End If
            
            If Texto <> "" Then
                If Mid(Texto, 1, 1) = "-" Then
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                Else
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                End If
            End If
            I = pos2 + 2
            j = j + 1
        Loop
        
        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) * rs_item!iteporctope / 100
        'FGZ - 14/10/2005
        If Items_TOPE(rs_Item!Itenro) < 0 Then
            Items_TOPE(rs_Item!Itenro) = 0
        Else
            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) * rs_Item!iteporctope / 100
        End If
    
    
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_Ano & _
                 " AND itenro =" & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
         Do While Not rs_Desmen.EOF
            If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                If rs_Desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                Else
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                End If
            End If
            ' Tocado por Maxi 26/05/2004 faltaba el parejito
            'If Month(rs_desmen!desfecdes) <= Ret_mes Then
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
            'Else
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
            'End If
         
            ' FGZ - 19/04/2004
            If rs_Item!Itenro = 20 Then 'Honorarios medicos
                If Not EsNulo(rs_Desmen!descuit) Then
                    StrSql = "UPDATE traza_gan SET "
                    StrSql = StrSql & " cuit_entidad9 ='" & rs_Desmen!descuit & "',"
                    StrSql = StrSql & " entidad9='" & rs_Desmen!DesRazsoc & "',"
                    StrSql = StrSql & " monto_entidad9=" & rs_Desmen!desmondec
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
                    'StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                    OpenRecordset StrSql, rs_Traza_gan
                    
                    Tope = Tope + 1
                End If
            End If
            ' Se saca el 23/05/2006
            If (rs_Item!Itenro = 29) Or (rs_Item!Itenro = 55) Then 'Impuesto al debito bancario
                StrSql = "UPDATE traza_gan SET "
                StrSql = StrSql & " promo =" & val_impdebitos
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                'StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            
                'FGZ - 22/12/2004
                'Leo la tabla
                StrSql = "SELECT * FROM traza_gan WHERE "
                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                'StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                OpenRecordset StrSql, rs_Traza_gan
            End If
            ' FGZ - 19/04/2004
            
            rs_Desmen.MoveNext
         Loop
    
    
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_Desliq

        Do While Not rs_Desliq.EOF
            Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto

            rs_Desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro=" & rs_Item!Itenro & _
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
            rs_itemacum.MoveNext
        Loop
        ' FGZ - 05/03/2004 Nuevo Hasta acá -------------------------
        
        ' FGZ - como prevliq y conliq se unieron en detliq queda uno solo
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

        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            Items_TOPE(rs_Item!Itenro) = Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
        Else
            'FGZ - 24/08/2005
            If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) = 0 Then
                Items_TOPE(rs_Item!Itenro) = 0
            End If
            'FGZ - 24/08/2005
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_Item!itesigno) Then
            Items_TOPE(rs_Item!Itenro) = -Items_TOPE(rs_Item!Itenro)
        End If

    ' End case 5
    ' ------------------------------------------------------------------------
    Case Else:
    End Select
   End If
    
    ' SI SE TOMA PARA LA GANANCIA NETA, DA VUELTA EL SIGNO DEL TOPE SOLO PARA ITEMS
    ' QUE SE TOPEAN DE ALGUNA FORMA Y NO SALEN DEL RECIBO DE SUELDO.
    ' "Como saber que no sale del Recibo" ?
    
    If rs_Item!Itenro > 7 Then
        Items_TOPE(rs_Item!Itenro) = IIf(CBool(rs_Item!itesigno), Items_TOPE(rs_Item!Itenro), Abs(Items_TOPE(rs_Item!Itenro)))
    End If
    
    
    'Armo la traza del item
    If CBool(USA_DEBUG) Then
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-DDJJ" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_DDJJ(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Liq" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_LIQ(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-LiqAnt" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_OLD_LIQ(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Prorr" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_PRORR(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Tope" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_TOPE(rs_Item!Itenro)
    End If
    If HACE_TRAZA Then
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_OLD_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_PRORR(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_TOPE(rs_Item!Itenro))
    End If
        
    
    'Calcula la Ganancia Imponible
    If CBool(rs_Item!itesigno) Then
        'FGZ - 13/09/2005
        'los items que suman en descuentos
        If rs_Item!Itenro >= 5 Then
            Descuentos = Descuentos + Items_TOPE(rs_Item!Itenro)
        End If
    
        Gan_Imponible = Gan_Imponible + Items_TOPE(rs_Item!Itenro)
    Else
        If (rs_Item!itetipotope = 1) Or (rs_Item!itetipotope = 4) Then
            Ded_a23 = Ded_a23 - Items_TOPE(rs_Item!Itenro)
            Items_ART_23(rs_Item!Itenro) = True
        Else
            Deducciones = Deducciones - Items_TOPE(rs_Item!Itenro)
        End If
    End If
            
    rs_Item.MoveNext
Loop
            
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "----------------------------------------------"
        'Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Neta: " & Gan_Imponible
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Bruta: " & (Gan_Imponible - Descuentos + Items_TOPE(50))
        Flog.writeline Espacios(Tabulador * 3) & "9- Gan. Bruta - CMA y DONA.: " & Gan_Imponible
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Neta: " & (Gan_Imponible + Deducciones)
        Flog.writeline Espacios(Tabulador * 3) & "9- Total Deducciones: " & Deducciones
        Flog.writeline Espacios(Tabulador * 3) & "9- Total art. 23: " & Ded_a23
    End If
    If HACE_TRAZA Then
        'Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Neta ", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Bruta ", Gan_Imponible - Descuentos + Items_TOPE(100))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Gan. Bruta - CMA y DONA.", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Neta ", (Gan_Imponible + Deducciones))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Para Machinea ", (Gan_Imponible + Deducciones - Items_TOPE(100)))
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
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    OpenRecordset StrSql, rs_Traza_gan
    
    ' Para el SAC de diciembre 2008 (item 50) se resta el monto para entrar a deducciones
    If Ret_Ano >= 2000 And Gan_Imponible > 0 Then
        StrSql = "SELECT * FROM escala_ded " & _
                 " WHERE esd_topeinf <= " & ((Gan_Imponible + Deducciones - Items_TOPE(50)) / Ret_Mes * 12) & _
                 " AND esd_topesup >=" & ((Gan_Imponible + Deducciones - Items_TOPE(50)) / Ret_Mes * 12)
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
        
'        'Guardo el porcentaje de deduccion
    StrSql = "UPDATE traza_gan SET "
    StrSql = StrSql & " porcdeduc =" & Por_Deduccion
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
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
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    OpenRecordset StrSql, rs_Traza_gan
    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible" & Gan_Imponible
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Imponible", Gan_Imponible)
    End If
            
                
    If Gan_Imponible > 0 Then
        'Entrar en la escala con las ganancias acumuladas
        StrSql = "SELECT * FROM escala " & _
                 " WHERE escmes =" & Ret_Mes & _
                 " AND escano =" & Ret_Ano & _
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
    I = 18
    
    Do While I <= 100
        'FGZ - 22/07/2005
        'el item 30 no debe sumar en otros
        If I <> 30 Then
            Otros = Otros + Abs(Items_TOPE(I))
        End If
        I = I + 1
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
    'FGZ - 23/07/2005
    'StrSql = StrSql & ", viaticos = 0"
    StrSql = StrSql & ", viaticos = " & (Items_TOPE(30))
    'FGZ - 23/07/2005
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
'    StrSql = "SELECT * FROM ficharet " & _
'             " WHERE empleado =" & buliq_empleado!ternro And Ficharet.Fecha >= fecha_ini_año
'    OpenRecordset StrSql, rs_Ficharet
'
'    Do While Not rs_Ficharet.EOF
'        If (Month(rs_Ficharet!Fecha) <= Ret_Mes) And (Year(rs_Ficharet!Fecha) = Ret_Ano) Then
'            Ret_Ant = Ret_Ant + rs_Ficharet!importe
'        End If
'        rs_Ficharet.MoveNext
'    Loop
    
    'Armo Fecha hasta como el ultimo dia del mes
    If (Ret_Mes = 12) Then
        fechaFichaH = CDate("31/12/" & Ret_Ano)
    Else
        fechaFichaH = CDate("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1
    End If
    
    fechaFichaD = CDate("01/01/" & Ret_Ano)
    
    StrSql = "SELECT SUM(importe) monto FROM ficharet " & _
             " WHERE empleado =" & buliq_empleado!ternro & _
             " AND fecha <= " & ConvFecha(fechaFichaH) & _
             " AND fecha >= " & ConvFecha(fechaFichaD)
    OpenRecordset StrSql, rs_Ficharet
    If Not rs_Ficharet.EOF Then
        If Not IsNull(rs_Ficharet!Monto) Then
            Ret_Ant = Ret_Ant + rs_Ficharet!Monto
        End If
    End If
    rs_Ficharet.Close
    
    'Calcular la retencion
    Retencion = Impuesto_Escala - Ret_Ant
    
    'Calculo de Impuesto y Debitos Bancarios, solo aplica si el impuesto retiene, si devuelve para el otro año lo declarado para este item
    If Retencion > 0 Then
        If val_impdebitos > Retencion Then
            val_impdebitos = Retencion
            Retencion = 0
        Else
            Retencion = Retencion - val_impdebitos
        End If
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Imp Debitos Banc", val_impdebitos)
    End If
    'Si hay devolucion suma los impdebitos pedido por Ruben Vacarezza 22/02/2008
    If Retencion < 0 Then
            Retencion = Retencion - val_impdebitos
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "Imp Debitos Banc", val_impdebitos)
    End If
    
    
    
    ' Para el F649 va en el 9b
    StrSql = "UPDATE traza_gan SET "
    StrSql = StrSql & " promo =" & val_impdebitos
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
            
    
    'FGZ - 30/12/2004
    'Determinar el saldo
    StrSql = "SELECT * FROM traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Not rs_Traza_gan.EOF Then
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
        'StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
        If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
        OpenRecordset StrSql, rs_Traza_gan
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
    I = 1
    Hasta = 100
    Do While I <= Hasta
        'FGZ 23/04/2004
        ' guardo los item_tope para poder usarlo en el reporte de Ganancias
        If Items_TOPE(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope
            
            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET monto =" & Items_TOPE(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
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
        
        If Items_PRORR(I) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!ternro & "," & _
                     ConvFecha(buliq_proceso!profecpago) & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_PRORR(I) & "," & _
                     "-1," & _
                     I & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        
        'FGZ 31/03/2005
        ' guardo los item_ddjj para poder usarlo en el reporte de Ganancias
        If Items_DDJJ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,ddjj,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_DDJJ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET ddjj =" & Items_DDJJ(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET prorr =" & Items_PRORR(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_old_liq para poder usarlo en el reporte de Ganancias
        If Items_OLD_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET old_liq =" & Items_OLD_LIQ(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_liq para poder usarlo en el reporte de Ganancias
        If Items_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET liq =" & Items_LIQ(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 31/03/2005
        
        I = I + 1
    Loop

    exito = Bien
    for_Ganancias = Monto
    
' Cierro todo y libero
'  If rs_WF_EscalaUTM.State = adStateOpen Then rs_WF_EscalaUTM.Close
'    Set rs_WF_EscalaUTM = Nothing
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
If rs_Traza_gan_items_tope.State = adStateOpen Then rs_Traza_gan_items_tope.Close
    Set rs_Traza_gan_items_tope = Nothing
If rs_escala_ded.State = adStateOpen Then rs_escala_ded.Close
    Set rs_escala_ded = Nothing
If rs_acumulador.State = adStateOpen Then rs_acumulador.Close
    Set rs_acumulador = Nothing

End Function


Public Function for_Ganancias_Schering(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
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
Dim Devuelve As Double
Dim Tope_Gral As Double
Dim Neto As Double
Dim prorratea As Double
Dim Retencion As Double
Dim Gan_Imponible As Double
Dim Deducciones As Double
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
Dim rs_escala_ded As New ADODB.Recordset
Dim rs_escala As New ADODB.Recordset
Dim rs_Ficharet As New ADODB.Recordset
Dim rs_Traza_gan_items_tope As New ADODB.Recordset

Dim Hasta As Integer

' FGZ - 12/02/2004
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Double
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
Dim Total_Empresa As Double
Dim Tope As Integer
'Dim rs_Rep19 As New ADODB.Recordset
Dim rs_Traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Double
Total_Empresa = 0
Tope = 10

' Primero limpio la traza
StrSql = "DELETE FROM traza_gan WHERE "
StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
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
         buliq_empleado!Empleg & _
         ")"
objConn.Execute StrSql, , adExecuteNoRecords

'FGZ - 22/12/2004
'Leo la tabla
StrSql = "SELECT * FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
OpenRecordset StrSql, rs_Traza_gan
    

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
ini_anyo_ret = C_Date("01/01/" & Ret_Ano)
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
    Ret_Mes = 12
    'FGZ - 27/09/2004
    fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
End If

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 3) & "---------Formula-----------------------------"
    Flog.writeline Espacios(Tabulador * 3) & "Fecha del Proceso " & buliq_proceso!profecpago
    Flog.writeline Espacios(Tabulador * 3) & "Mes de Retencion " & Ret_Mes
    Flog.writeline Espacios(Tabulador * 3) & "Año de Retencion " & Ret_Ano
    Flog.writeline Espacios(Tabulador * 3) & "Fin mes de Retencion " & fin_mes_ret
    
    Flog.writeline Espacios(Tabulador * 3) & "Máxima Ret. en % " & Tope_Gral
    Flog.writeline Espacios(Tabulador * 3) & "Neto del Mes " & Neto
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Tope_Gral, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, p_Neto, "Neto del Mes", Neto)
End If


'Limpiar items que suman al articulo 23
For I = 1 To 50
    Items_ART_23(I) = False
Next I



' Recorro todos los items de Ganancias
StrSql = "SELECT * FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_Item

Do While Not rs_Item.EOF
    
    Select Case rs_Item!itetipotope
    Case 1: ' el valor a tomar es lo que dice la escala
    
        StrSql = "SELECT * FROM valitem WHERE viano =" & Ret_Ano & _
                 " AND itenro=" & rs_Item!Itenro & _
                 " AND vimes =" & Ret_Mes
        OpenRecordset StrSql, rs_valitem
        
        Do While Not rs_valitem.EOF
            Items_DDJJ(rs_valitem!Itenro) = rs_valitem!vimonto
            Items_TOPE(rs_valitem!Itenro) = rs_valitem!vimonto
            
            rs_valitem.MoveNext
        Loop
    ' End case 1
    ' ------------------------------------------------------------------------
    
    Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
        ' Busco la declaracion jurada
        StrSql = "SELECT * FROM desmen WHERE empleado =" & buliq_empleado!ternro & _
                 " AND desano=" & Ret_Ano & _
                 " AND itenro = " & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
        
        Do While Not rs_Desmen.EOF
            If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                If rs_Item!Itenro = 3 Then
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / 12 * Ret_Mes, rs_Desmen!desmondec)
                    
                    'FGZ - 19/04/2004
                    If rs_Item!Itenro <= 4 Then
                        If Not EsNulo(rs_Desmen!descuit) Then
                            I = 11
                            Distinto = rs_Traza_gan!Cuit_entidad11 <> rs_Desmen!descuit
                            Do While (I <= Tope) And Distinto
                                I = I + 1
                                Select Case I
                                Case 11:
                                    Distinto = rs_Traza_gan!Cuit_entidad11 <> rs_Desmen!descuit
                                Case 12:
                                    Distinto = rs_Traza_gan!Cuit_entidad12 <> rs_Desmen!descuit
                                Case 13:
                                    Distinto = rs_Traza_gan!Cuit_entidad13 <> rs_Desmen!descuit
                                Case 14:
                                    Distinto = rs_Traza_gan!Cuit_entidad14 <> rs_Desmen!descuit
                                End Select
                            Loop
                          
                            If I > Tope And I <= 14 Then
                                StrSql = "UPDATE traza_gan SET "
                                StrSql = StrSql & " cuit_entidad" & I & "='" & rs_Desmen!descuit & "',"
                                StrSql = StrSql & " entidad" & I & "='" & rs_Desmen!DesRazsoc & "',"
                                StrSql = StrSql & " monto_entidad" & I & "=" & rs_Desmen!desmondec
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
                                'StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                OpenRecordset StrSql, rs_Traza_gan
                                
                                
                                Tope = Tope + 1
                            Else
                                If I = 15 Then
                                    Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!Empleg ' empleado.empleg
                                Else
                                    StrSql = "UPDATE traza_gan SET "
                                    StrSql = StrSql & " monto_entidad" & I & "= monto_entidad" & I & " + " & rs_Desmen!desmondec
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
                                    'StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                                    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                    OpenRecordset StrSql, rs_Traza_gan
                                End If
                            End If
                        Else
                            Total_Empresa = Total_Empresa + rs_Desmen!desmondec
                        End If
                    End If
                    'FGZ - 19/04/2004
                    
                Else
                    If rs_Desmen!desmenprorra = 0 Then 'no es parejito
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + rs_Desmen!desmondec
                    Else
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / 12 * Ret_Mes, rs_Desmen!desmondec)
                    End If
                End If
            End If
            
            
            rs_Desmen.MoveNext
        Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_Desliq

        Do While Not rs_Desliq.EOF
            Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
            'Si el desliq prorratea debo proporcionarlo
            'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 Or CBool(rs_desliq!dlprorratea)) And (prorratea = 1), rs_desliq!dlmonto / (13 - Month(rs_desliq!dlfecha)) * (Ret_mes - Month(rs_desliq!dlfecha) + 1), rs_desliq!dlmonto)
            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), rs_Desliq!Dlmonto / 12 * Ret_Mes, rs_Desliq!Dlmonto)
            rs_Desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_Item!Itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acuNro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itaprorratea) Then
                    If CBool(rs_itemacum!itasigno) Then
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + Aux_Acu_Monto
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / 12 * Ret_Mes, Aux_Acu_Monto)
                    Else
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - Aux_Acu_Monto
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - IIf(prorratea = 1, Aux_Acu_Monto / 12 * Ret_Mes, Aux_Acu_Monto)
                    End If
                Else
                    If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 And prorratea = 1), Aux_Acu_Monto / 12 * Ret_Mes, Aux_Acu_Monto)
                    Else
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
                        'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), Aux_Acu_Monto / (13 - Ret_mes), Aux_Acu_Monto)
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - IIf((rs_Item!Itenro = 3 And prorratea = 1), Aux_Acu_Monto / 12 * Ret_Mes, Aux_Acu_Monto)
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
                 " AND itemconc.itenro =" & rs_Item!Itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) Then
                If CBool(rs_itemconc!itcsigno) Then
                    Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + rs_itemconc!dlimonto
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / 12 * Ret_Mes, rs_itemconc!dlimonto)
                Else
                    Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - rs_itemconc!dlimonto
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / 12 * Ret_Mes, rs_itemconc!dlimonto)
                End If
            Else
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + rs_itemconc!dlimonto
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / 12 * Ret_Mes, rs_itemconc!dlimonto)
                Else
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - rs_itemconc!dlimonto
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - IIf((rs_item!itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / (13 - Ret_mes), rs_itemconc!dlimonto)
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - IIf((rs_Item!Itenro = 3 And prorratea = 1), rs_itemconc!dlimonto / 12 * Ret_Mes, rs_itemconc!dlimonto)
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
     
        StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_Ano & _
                 " AND vimes = " & Ret_Mes & _
                 " AND itenro =" & rs_Item!Itenro
        OpenRecordset StrSql, rs_valitem
         Do While Not rs_valitem.EOF
            Items_TOPE(rs_Item!Itenro) = rs_valitem!vimonto
         
            rs_valitem.MoveNext
         Loop
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_Ano & _
                 " AND itenro =" & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
         Do While Not rs_Desmen.EOF
            If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                If rs_Desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                Else
                    'Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / 12 * Ret_Mes, rs_Desmen!desmondec)
                End If
            End If
         
            rs_Desmen.MoveNext
         Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_Desliq

        Do While Not rs_Desliq.EOF
                        If CBool(rs_Desliq!Dlprorratea) And (prorratea = 1) Then
                                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto / 12 * Ret_Mes
                        Else
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
                        End If
            rs_Desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro =" & rs_Item!Itenro & _
                 " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemacum
        Do While Not rs_itemacum.EOF
            Acum = CStr(rs_itemacum!acuNro)
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
        
                If CBool(rs_itemacum!itaprorratea) Then
                    If CBool(rs_itemacum!itasigno) Then
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), (Aux_Acu_Monto / 12 * Ret_Mes), Aux_Acu_Monto)
                    Else
                        Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), (Aux_Acu_Monto / 12 * Ret_Mes), Aux_Acu_Monto)
                    End If
                Else
                        If CBool(rs_itemacum!itasigno) Then
                            Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Aux_Acu_Monto
                        Else
                            Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - Aux_Acu_Monto
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
                 " AND itemconc.itenro =" & rs_Item!Itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) Then
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / 12 * Ret_Mes, rs_itemconc!dlimonto)
                Else
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / 12 * Ret_Mes, rs_itemconc!dlimonto)
                End If
            Else
                If CBool(rs_itemconc!itcsigno) Then
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + rs_itemconc!dlimonto
                Else
                    Items_LIQ(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) - rs_itemconc!dlimonto
                End If
            End If
        
            rs_itemconc.MoveNext
        Loop
        
        'Topeo los valores
        'Tomo los valores con signo negativo, ya que salen de la liquidacion y forman parte del neto
        ' Mauricio 15-03-2000
        
        
        'LLEVO TODO A ABSOLUTO PARA PODER COMPARAR CONTRA LA ESCALA
        If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            Items_TOPE(rs_Item!Itenro) = Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_Item!itesigno) Then
            Items_TOPE(rs_Item!Itenro) = -Items_TOPE(rs_Item!Itenro)
        End If
        
    ' End case 3
    ' ------------------------------------------------------------------------
    Case 4:
        ' Tomo los valores de la DDJJ y el valor de la escala (cargas de familia)
        
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_Ano & _
                 " AND itenro =" & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
         Do While Not rs_Desmen.EOF
            If Month(rs_Desmen!desfechas) <= Ret_Mes Then
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec * (Month(rs_Desmen!desfechas) - Month(rs_Desmen!desfecdes) + 1)
            Else
                If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1)
                End If
            End If
        
            rs_Desmen.MoveNext
         Loop
        
        If Items_DDJJ(rs_Item!Itenro) > 0 Then
            StrSql = "SELECT * FROM valitem WHERE viano = " & Ret_Ano & _
                     " AND vimes = " & Ret_Mes & _
                     " AND itenro =" & rs_Item!Itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_Item!Itenro) = rs_valitem!vimonto / Ret_Mes * Items_DDJJ(rs_Item!Itenro)
             
                rs_valitem.MoveNext
             Loop
        End If
    ' End case 4
    ' ------------------------------------------------------------------------
        
    Case 5:
        I = 1
        j = 1
        'Hasta = IIf(50 > Len(rs_item!iteitemstope), 50, rs_item!iteitemstope)
        Hasta = 50
        Terminar = False
        Do While j <= Hasta And Not Terminar
            pos1 = I
            pos2 = InStr(I, rs_Item!iteitemstope, ",") - 1
            If pos2 > 0 Then
                Texto = Mid(rs_Item!iteitemstope, pos1, pos2 - pos1 + 1)
            Else
                pos2 = Len(rs_Item!iteitemstope)
                Texto = Mid(rs_Item!iteitemstope, pos1, pos2 - pos1 + 1)
                Terminar = True
            End If
            
            If Texto <> "" Then
                If Mid(Texto, 1, 1) = "-" Then
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) - Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                Else
                    'Items_TOPE(rs_item!itenro) = Items_TOPE(rs_item!itenro) + Items_TOPE(Mid(rs_item!iteitemstope, 2, InStr(1, rs_item!iteitemstope, ",") - 2))
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + Items_TOPE(Mid(Texto, 2, Len(Texto) - 1))
                End If
            End If
            I = pos2 + 2
            j = j + 1
        Loop
        
        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) * rs_Item!iteporctope / 100
    
    
        'Busco la declaracion Jurada
        StrSql = "SELECT * FROM desmen WHERE empleado = " & buliq_empleado!ternro & _
                 " AND desano = " & Ret_Ano & _
                 " AND itenro =" & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
         Do While Not rs_Desmen.EOF
            If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                If rs_Desmen!desmenprorra = 0 Then ' No es parejito
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                Else
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                End If
            End If
            ' Tocado por Maxi 26/05/2004 faltaba el parejito
            'If Month(rs_desmen!desfecdes) <= Ret_mes Then
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + rs_desmen!desmondec
            'Else
            '    Items_DDJJ(rs_item!itenro) = Items_DDJJ(rs_item!itenro) + IIf((prorratea = 1) And (Ret_mes <= Month(rs_desmen!desfechas)), rs_desmen!desmondec / (Month(rs_desmen!desfechas) + 1 - Month(rs_desmen!desfecdes)) * (Ret_mes - Month(rs_desmen!desfecdes) + 1), rs_desmen!desmondec)
            'End If
         
            ' FGZ - 19/04/2004
            If rs_Item!Itenro = 20 Then 'Honorarios medicos
                If Not EsNulo(rs_Desmen!descuit) Then
                    StrSql = "UPDATE traza_gan SET "
                    StrSql = StrSql & " cuit_entidad9 ='" & rs_Desmen!descuit & "',"
                    StrSql = StrSql & " entidad9='" & rs_Desmen!DesRazsoc & "',"
                    StrSql = StrSql & " monto_entidad9=" & rs_Desmen!desmondec
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
                    'StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                    OpenRecordset StrSql, rs_Traza_gan
                    
                    Tope = Tope + 1
                End If
            End If
            If rs_Item!Itenro = 22 Then 'Impuesto al debito bancario
                StrSql = "UPDATE traza_gan SET "
                StrSql = StrSql & " promo =" & rs_Desmen!desmondec
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
                'StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                OpenRecordset StrSql, rs_Traza_gan
            End If
            ' FGZ - 19/04/2004
            
            rs_Desmen.MoveNext
         Loop
    
    
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_Desliq

        Do While Not rs_Desliq.EOF
            Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto

            rs_Desliq.MoveNext
        Loop
        
        'Busco los acumuladores de la liquidacion
        ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
        StrSql = "SELECT * FROM itemacum " & _
                 " WHERE itenro=" & rs_Item!Itenro & _
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
        If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            Items_TOPE(rs_Item!Itenro) = Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
        End If
        
        'SI ES GANANCIA NETA, ENTONCES LA VUELVO A NEGATIVO PARA QUE LA DISMINUYA, YA QUE ESTE TOPE TRATA SOLO
        ' "ACHIQUE" DE GANANCIA IMPONIBLE
        If CBool(rs_Item!itesigno) Then
            Items_TOPE(rs_Item!Itenro) = -Items_TOPE(rs_Item!Itenro)
        End If

    ' End case 5
    ' ------------------------------------------------------------------------
    Case Else:
    End Select
    
    
    ' SI SE TOMA PARA LA GANANCIA NETA, DA VUELTA EL SIGNO DEL TOPE SOLO PARA ITEMS
    ' QUE SE TOPEAN DE ALGUNA FORMA Y NO SALEN DEL RECIBO DE SUELDO.
    ' "Como saber que no sale del Recibo" ?
    
    If rs_Item!Itenro > 7 Then
        Items_TOPE(rs_Item!Itenro) = IIf(CBool(rs_Item!itesigno), Items_TOPE(rs_Item!Itenro), Abs(Items_TOPE(rs_Item!Itenro)))
    End If
    
    
    'Armo la traza del item
    If CBool(USA_DEBUG) Then
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-DDJJ" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_DDJJ(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Liq" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_LIQ(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-LiqAnt" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_OLD_LIQ(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Prorr" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_PRORR(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Tope" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_TOPE(rs_Item!Itenro)
    End If
    If HACE_TRAZA Then
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_OLD_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_PRORR(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, Texto, Items_TOPE(rs_Item!Itenro))
    End If
        
    
    'Calcula la Ganancia Imponible
    If CBool(rs_Item!itesigno) Then
        Gan_Imponible = Gan_Imponible + Items_TOPE(rs_Item!Itenro)
    Else
        If (rs_Item!itetipotope = 1) Or (rs_Item!itetipotope = 4) Then
            Ded_a23 = Ded_a23 - Items_TOPE(rs_Item!Itenro)
            Items_ART_23(rs_Item!Itenro) = True
        Else
            Deducciones = Deducciones - Items_TOPE(rs_Item!Itenro)
        End If
    End If
            
    rs_Item.MoveNext
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
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    OpenRecordset StrSql, rs_Traza_gan
    
    If Ret_Ano >= 2000 And Gan_Imponible > 0 Then
        StrSql = "SELECT * FROM escala_ded " & _
                 " WHERE esd_topeinf <= " & ((Gan_Imponible + Deducciones) / Ret_Mes * 12) & _
                 " AND esd_topesup >=" & ((Gan_Imponible + Deducciones) / Ret_Mes * 12)
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
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    OpenRecordset StrSql, rs_Traza_gan
    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible" & Gan_Imponible
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 0, "9- Ganancia Imponible", Gan_Imponible)
    End If
            
                
    If Gan_Imponible > 0 Then
        'Entrar en la escala con las ganancias acumuladas
        StrSql = "SELECT * FROM escala " & _
                 " WHERE escmes =" & Ret_Mes & _
                 " AND escano =" & Ret_Ano & _
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
    I = 18
    
    Do While I <= 50
        Otros = Otros + Abs(Items_TOPE(I))
        I = I + 1
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
    OpenRecordset StrSql, rs_Ficharet
    
    Do While Not rs_Ficharet.EOF
        If (Month(rs_Ficharet!Fecha) <= Ret_Mes) And (Year(rs_Ficharet!Fecha) = Ret_Ano) Then
            Ret_Ant = Ret_Ant + rs_Ficharet!importe
        End If
        rs_Ficharet.MoveNext
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
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Not rs_Traza_gan.EOF Then
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
        'StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
        If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
        OpenRecordset StrSql, rs_Traza_gan
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
    I = 1
    Hasta = 50
    Do While I <= Hasta
        'FGZ 23/04/2004
        ' guardo los item_tope para poder usarlo en el reporte de Ganancias
        If Items_TOPE(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope
            
            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET monto =" & Items_TOPE(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
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
        
        If Items_PRORR(I) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!ternro & "," & _
                     ConvFecha(buliq_proceso!profecpago) & "," & _
                     buliq_proceso!pronro & "," & _
                     Items_PRORR(I) & "," & _
                     "0," & _
                     I & _
                     ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        
        
        'FGZ 31/03/2005
        ' guardo los item_ddjj para poder usarlo en el reporte de Ganancias
        If Items_DDJJ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,ddjj,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_DDJJ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET ddjj =" & Items_DDJJ(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET prorr =" & Items_PRORR(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_old_liq para poder usarlo en el reporte de Ganancias
        If Items_OLD_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET old_liq =" & Items_OLD_LIQ(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_liq para poder usarlo en el reporte de Ganancias
        If Items_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE traza_gan_item_top SET liq =" & Items_LIQ(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 31/03/2005
        
        I = I + 1
    Loop

    exito = Bien
    for_Ganancias_Schering = Monto
End Function


Public Sub InsertarFichaRet(ByVal ternro As Long, ByVal Fecha As Date, ByVal importe As Double, ByVal pronro As Long)
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
             importe & "," & _
             "-1" & _
             ")"
    objConn.Execute StrSql, , adExecuteNoRecords


End Sub



Public Function For_Baseacci() As Double
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
Dim monto_manual     As Double
Dim divisor          As Double
Dim multiplicador    As Double
Dim cant_meses       As Integer
Dim fecha_desde      As Date
Dim fecha_hasta      As Date
Dim Meses            As Integer
Dim Meses_Validos    As Integer
Dim Monto_Acum       As Double
Dim Monto_Mes        As Double
Dim Texto            As String
Dim mestope As Integer
Dim fectope As Date
Dim Anio As Integer
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
    fecha_desde = IIf(Month(buliq_periodo!pliqdesde) = 1, C_Date("01/12/" & Year(buliq_periodo!pliqdesde) - 1), C_Date("01/" & Month(buliq_periodo!pliqdesde) - 1 & "/" & Year(buliq_periodo!pliqdesde)))
                            
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
                Flog.writeline Espacios(Tabulador * 1) & "No se ha encontrado el complemento de la licencia " & rs_Emp_Lic!emp_licnro & " del empleado " & buliq_empleado!Empleg
            End If
        Else
            StrSql = "SELECT * FROM so_accidente "
            StrSql = StrSql & " WHERE rs_accidente.accnro = " & rs_Lic_Accid!accnro
            OpenRecordset StrSql, rs_Accidente
                    
            If rs_Accidente.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 1) & "No se ha encontrado el accidente de la licencia " & rs_Emp_Lic!emp_licnro & " del empleado " & buliq_empleado!Empleg
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
                        Flog.writeline Espacios(Tabulador * 1) & "No se ha encontrado un periodo de liquidacion para el accidente de la licencia " & rs_Emp_Lic!emp_licnro & " del empleado " & buliq_empleado!Empleg
                    End If
                Else
                    'CALCULO DE ACIDENTES PARA JORNALES
                    If buliq_empleado!folinro = 2 Then
                        mestope = IIf(Month(rs_Accidente!accfecha) - 12 < 0, 12 + (Month(rs_Accidente!accfecha) - 12), Month(rs_Accidente!accfecha) - 12)
                        fectope = C_Date(Day(rs_Accidente!accfecha) & "/" & mestope & "/" & Year(rs_Accidente!accfecha) - 1)
                            
                        Call bus_Antfases(rs_Accidente!accfecha, fectope, dia, mes, Anio)
                        'RUN antfases.p (recid(buliq-empleado), accidente.accfecha, fectope, output dia, output mes, output anio)
                        divisor = (Anio * 360) + (mes * 30) + dia
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
                        fectope = C_Date(Day(rs_Accidente!accfecha) & "/" & mestope & "/" & Year(rs_Accidente!accfecha) - 1)
                        
                        Call bus_Antfases(rs_Accidente!accfecha, fectope, dia, mes, Anio)
                        'RUN antfases.p (recid(buliq-empleado), accidente.accfecha, fectope, output dia, output mes, output anio).
                        divisor = (Anio * 360) + (mes * 30) + dia
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


Public Function for_Nivelar(ByVal AFecha As Date) As Double
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
Dim Valor_Nivel   As Double
Dim Valor_Neto    As Double

Dim entero        As Long
Dim Neto_Truncado As Double
Dim Decimales     As Double
Dim v_monto As Double

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


Public Function for_Ganancias_Petroleros(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias sin horas extras.
' Autor      : FGZ
' Fecha      : 24/01/2007
' Ultima Mod.:
' Descripcion: se agregó el item 30 (movilidad) y se computa su valor en traza_gan.viaticos.
' ---------------------------------------------------------------------------------------------
Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim p_prorratea As Integer  'Si prorratea o no para liq. finales
Dim p_Ajusta As Integer  'Si ajusta ficharet o no

Dim I As Long
Dim Ajusta_Ficharet As Boolean
Dim Ret_Real As Double
Dim Ret_Sin_Ext As Double
Dim Ret_Mes As Integer
Dim Ret_Ano As Integer
Dim Ret_Actual As Double

'Vectores para manejar temporales
Dim Traza_Gan(1 To 10) As TTraza_Gan
Dim Traza_Gan_Item_Top(1 To 100) As TTraza_Gan_Item_Top
Dim Desliq(1 To 100) As TDesliq
Dim Ficharet(1 To 1) As TFicharet

'Recorsets Auxiliares
Dim rs_Desliq As New ADODB.Recordset
Dim rs_Traza_gan As New ADODB.Recordset
Dim rs_Traza_gan_items_tope As New ADODB.Recordset
Dim rs_Ficharet As New ADODB.Recordset
Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Periodos As New ADODB.Recordset

'Comienzo
    p_Ajusta = 1024
    
    Bien = False

    'Obtencion de los parametros de WorkFile
    Ajusta_Ficharet = False
    StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha =" & ConvFecha(AFecha)
    OpenRecordset StrSql, rs_wf_tpa
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case p_Ajusta:
            Ajusta_Ficharet = CBool(rs_wf_tpa!Valor)
        End Select
        
        rs_wf_tpa.MoveNext
    Loop


    '-----------------------------------------------------------------------------
    '   Depuracion y almacenamiento de temporales
    
    
    'TRAZA_GAN --->
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "--- Depuracion de traza_gan ---"
    End If
    'guardo todos los traza_gan liquidados
    StrSql = "SELECT * FROM traza_gan WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    'StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).Concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Traza_gan
    I = 0
    Do While Not rs_Traza_gan.EOF
        I = I + 1
        Traza_Gan(I).PliqNro = IIf(Not EsNulo(rs_Traza_gan!PliqNro), rs_Traza_gan!PliqNro, 0)
        Traza_Gan(I).concnro = IIf(Not EsNulo(rs_Traza_gan!concnro), rs_Traza_gan!concnro, 0)
        Traza_Gan(I).Empresa = IIf(Not EsNulo(rs_Traza_gan!Empresa), rs_Traza_gan!Empresa, 0)
        Traza_Gan(I).Fecha_pago = IIf(Not EsNulo(rs_Traza_gan!Fecha_pago), rs_Traza_gan!Fecha_pago, 0)
        Traza_Gan(I).ternro = IIf(Not EsNulo(rs_Traza_gan!ternro), rs_Traza_gan!ternro, 0)
        Traza_Gan(I).Msr = IIf(Not EsNulo(rs_Traza_gan!Msr), rs_Traza_gan!Msr, 0)
        Traza_Gan(I).Nomsr = IIf(Not EsNulo(rs_Traza_gan!Nomsr), rs_Traza_gan!Nomsr, 0)
        Traza_Gan(I).Nogan = IIf(Not EsNulo(rs_Traza_gan!Nogan), rs_Traza_gan!Nogan, 0)
        Traza_Gan(I).Jubilacion = IIf(Not EsNulo(rs_Traza_gan!Jubilacion), rs_Traza_gan!Jubilacion, 0)
        Traza_Gan(I).Osocial = IIf(Not EsNulo(rs_Traza_gan!Osocial), rs_Traza_gan!Osocial, 0)
        Traza_Gan(I).Cuota_medico = IIf(Not EsNulo(rs_Traza_gan!Cuota_medico), rs_Traza_gan!Cuota_medico, 0)
        Traza_Gan(I).Prima_seguro = IIf(Not EsNulo(rs_Traza_gan!Prima_seguro), rs_Traza_gan!Prima_seguro, 0)
        Traza_Gan(I).Sepelio = IIf(Not EsNulo(rs_Traza_gan!Sepelio), rs_Traza_gan!Sepelio, 0)
        Traza_Gan(I).Estimados = IIf(Not EsNulo(rs_Traza_gan!Estimados), rs_Traza_gan!Estimados, 0)
        Traza_Gan(I).Otras = IIf(Not EsNulo(rs_Traza_gan!Otras), rs_Traza_gan!Otras, 0)
        Traza_Gan(I).Donacion = IIf(Not EsNulo(rs_Traza_gan!Donacion), rs_Traza_gan!Donacion, 0)
        Traza_Gan(I).Dedesp = IIf(Not EsNulo(rs_Traza_gan!Dedesp), rs_Traza_gan!Dedesp, 0)
        Traza_Gan(I).Noimpo = IIf(Not EsNulo(rs_Traza_gan!Noimpo), rs_Traza_gan!Noimpo, 0)
        Traza_Gan(I).Car_flia = IIf(Not EsNulo(rs_Traza_gan!Car_flia), rs_Traza_gan!Car_flia, 0)
        Traza_Gan(I).conyuge = IIf(Not EsNulo(rs_Traza_gan!conyuge), rs_Traza_gan!conyuge, 0)
        Traza_Gan(I).Hijo = IIf(Not EsNulo(rs_Traza_gan!Hijo), rs_Traza_gan!Hijo, 0)
        Traza_Gan(I).Otras_cargas = IIf(Not EsNulo(rs_Traza_gan!Otras_cargas), rs_Traza_gan!Otras_cargas, 0)
        Traza_Gan(I).Retenciones = IIf(Not EsNulo(rs_Traza_gan!Retenciones), rs_Traza_gan!Retenciones, 0)
        Traza_Gan(I).Promo = IIf(Not EsNulo(rs_Traza_gan!Promo), rs_Traza_gan!Promo, 0)
        Traza_Gan(I).Saldo = IIf(Not EsNulo(rs_Traza_gan!Saldo), rs_Traza_gan!Saldo, 0)
        Traza_Gan(I).Sindicato = IIf(Not EsNulo(rs_Traza_gan!Sindicato), rs_Traza_gan!Sindicato, 0)
        Traza_Gan(I).Ret_Mes = IIf(Not EsNulo(rs_Traza_gan!Ret_Mes), rs_Traza_gan!Ret_Mes, 0)
        Traza_Gan(I).Mon_conyuge = IIf(Not EsNulo(rs_Traza_gan!Mon_conyuge), rs_Traza_gan!Mon_conyuge, 0)
        Traza_Gan(I).Mon_hijo = IIf(Not EsNulo(rs_Traza_gan!Mon_hijo), rs_Traza_gan!Mon_hijo, 0)
        Traza_Gan(I).Mon_otras = IIf(Not EsNulo(rs_Traza_gan!Mon_otras), rs_Traza_gan!Mon_otras, 0)
        Traza_Gan(I).Viaticos = IIf(Not EsNulo(rs_Traza_gan!Viaticos), rs_Traza_gan!Viaticos, 0)
        Traza_Gan(I).Amortizacion = IIf(Not EsNulo(rs_Traza_gan!Amortizacion), rs_Traza_gan!Amortizacion, 0)
        Traza_Gan(I).Entidad1 = IIf(Not EsNulo(rs_Traza_gan!Entidad1), rs_Traza_gan!Entidad1, "")
        Traza_Gan(I).Entidad2 = IIf(Not EsNulo(rs_Traza_gan!Entidad2), rs_Traza_gan!Entidad2, "")
        Traza_Gan(I).Entidad3 = IIf(Not EsNulo(rs_Traza_gan!Entidad3), rs_Traza_gan!Entidad3, "")
        Traza_Gan(I).Entidad4 = IIf(Not EsNulo(rs_Traza_gan!Entidad4), rs_Traza_gan!Entidad4, "")
        Traza_Gan(I).Entidad5 = IIf(Not EsNulo(rs_Traza_gan!Entidad5), rs_Traza_gan!Entidad5, "")
        Traza_Gan(I).Entidad6 = IIf(Not EsNulo(rs_Traza_gan!Entidad6), rs_Traza_gan!Entidad6, "")
        Traza_Gan(I).Entidad7 = IIf(Not EsNulo(rs_Traza_gan!Entidad7), rs_Traza_gan!Entidad7, "")
        Traza_Gan(I).Entidad8 = IIf(Not EsNulo(rs_Traza_gan!Entidad8), rs_Traza_gan!Entidad8, "")
        Traza_Gan(I).Entidad9 = IIf(Not EsNulo(rs_Traza_gan!Entidad9), rs_Traza_gan!Entidad9, "")
        Traza_Gan(I).Entidad10 = IIf(Not EsNulo(rs_Traza_gan!Entidad10), rs_Traza_gan!Entidad10, "")
        Traza_Gan(I).Entidad11 = IIf(Not EsNulo(rs_Traza_gan!Entidad11), rs_Traza_gan!Entidad11, "")
        Traza_Gan(I).Entidad12 = IIf(Not EsNulo(rs_Traza_gan!Entidad12), rs_Traza_gan!Entidad12, "")
        Traza_Gan(I).Entidad13 = IIf(Not EsNulo(rs_Traza_gan!Entidad13), rs_Traza_gan!Entidad13, "")
        Traza_Gan(I).Entidad14 = IIf(Not EsNulo(rs_Traza_gan!Entidad14), rs_Traza_gan!Entidad14, "")
        Traza_Gan(I).Cuit_entidad1 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad1), rs_Traza_gan!Cuit_entidad1, "")
        Traza_Gan(I).Cuit_entidad2 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad2), rs_Traza_gan!Cuit_entidad2, "")
        Traza_Gan(I).Cuit_entidad3 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad3), rs_Traza_gan!Cuit_entidad3, "")
        Traza_Gan(I).Cuit_entidad4 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad4), rs_Traza_gan!Cuit_entidad4, "")
        Traza_Gan(I).Cuit_entidad5 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad5), rs_Traza_gan!Cuit_entidad5, "")
        Traza_Gan(I).Cuit_entidad6 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad6), rs_Traza_gan!Cuit_entidad6, "")
        Traza_Gan(I).Cuit_entidad7 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad7), rs_Traza_gan!Cuit_entidad7, "")
        Traza_Gan(I).Cuit_entidad8 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad8), rs_Traza_gan!Cuit_entidad8, "")
        Traza_Gan(I).Cuit_entidad9 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad9), rs_Traza_gan!Cuit_entidad9, "")
        Traza_Gan(I).Cuit_entidad10 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad10), rs_Traza_gan!Cuit_entidad10, "")
        Traza_Gan(I).Cuit_entidad11 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad11), rs_Traza_gan!Cuit_entidad11, "")
        Traza_Gan(I).Cuit_entidad12 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad12), rs_Traza_gan!Cuit_entidad12, "")
        Traza_Gan(I).Cuit_entidad13 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad13), rs_Traza_gan!Cuit_entidad13, "")
        Traza_Gan(I).Cuit_entidad14 = IIf(Not EsNulo(rs_Traza_gan!Cuit_entidad14), rs_Traza_gan!Cuit_entidad14, "")
        Traza_Gan(I).Monto_entidad1 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad1), rs_Traza_gan!Monto_entidad1, 0)
        Traza_Gan(I).Monto_entidad2 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad2), rs_Traza_gan!Monto_entidad2, 0)
        Traza_Gan(I).Monto_entidad3 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad3), rs_Traza_gan!Monto_entidad3, 0)
        Traza_Gan(I).Monto_entidad4 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad4), rs_Traza_gan!Monto_entidad4, 0)
        Traza_Gan(I).Monto_entidad5 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad5), rs_Traza_gan!Monto_entidad5, 0)
        Traza_Gan(I).Monto_entidad6 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad6), rs_Traza_gan!Monto_entidad6, 0)
        Traza_Gan(I).Monto_entidad7 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad7), rs_Traza_gan!Monto_entidad7, 0)
        Traza_Gan(I).Monto_entidad8 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad8), rs_Traza_gan!Monto_entidad8, 0)
        Traza_Gan(I).Monto_entidad9 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad9), rs_Traza_gan!Monto_entidad9, 0)
        Traza_Gan(I).Monto_entidad10 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad10), rs_Traza_gan!Monto_entidad10, 0)
        Traza_Gan(I).Monto_entidad11 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad11), rs_Traza_gan!Monto_entidad11, 0)
        Traza_Gan(I).Monto_entidad12 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad12), rs_Traza_gan!Monto_entidad12, 0)
        Traza_Gan(I).Monto_entidad13 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad13), rs_Traza_gan!Monto_entidad13, 0)
        Traza_Gan(I).Monto_entidad14 = IIf(Not EsNulo(rs_Traza_gan!Monto_entidad14), rs_Traza_gan!Monto_entidad14, 0)
        Traza_Gan(I).Ganimpo = IIf(Not EsNulo(rs_Traza_gan!Ganimpo), rs_Traza_gan!Ganimpo, 0)
        Traza_Gan(I).Ganneta = IIf(Not EsNulo(rs_Traza_gan!Ganneta), rs_Traza_gan!Ganneta, 0)
        Traza_Gan(I).Total_entidad1 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad1), rs_Traza_gan!Total_entidad1, 0)
        Traza_Gan(I).Total_entidad2 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad2), rs_Traza_gan!Total_entidad2, 0)
        Traza_Gan(I).Total_entidad3 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad3), rs_Traza_gan!Total_entidad3, 0)
        Traza_Gan(I).Total_entidad4 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad4), rs_Traza_gan!Total_entidad4, 0)
        Traza_Gan(I).Total_entidad5 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad5), rs_Traza_gan!Total_entidad5, 0)
        Traza_Gan(I).Total_entidad6 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad6), rs_Traza_gan!Total_entidad6, 0)
        Traza_Gan(I).Total_entidad7 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad7), rs_Traza_gan!Total_entidad7, 0)
        Traza_Gan(I).Total_entidad8 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad8), rs_Traza_gan!Total_entidad8, 0)
        Traza_Gan(I).Total_entidad9 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad9), rs_Traza_gan!Total_entidad9, 0)
        Traza_Gan(I).Total_entidad10 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad10), rs_Traza_gan!Total_entidad10, 0)
        Traza_Gan(I).Total_entidad11 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad11), rs_Traza_gan!Total_entidad11, 0)
        Traza_Gan(I).Total_entidad12 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad12), rs_Traza_gan!Total_entidad12, 0)
        Traza_Gan(I).Total_entidad13 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad13), rs_Traza_gan!Total_entidad13, 0)
        Traza_Gan(I).Total_entidad14 = IIf(Not EsNulo(rs_Traza_gan!Total_entidad14), rs_Traza_gan!Total_entidad14, 0)
        Traza_Gan(I).pronro = IIf(Not EsNulo(rs_Traza_gan!pronro), rs_Traza_gan!pronro, 0)
        Traza_Gan(I).Imp_deter = IIf(Not EsNulo(rs_Traza_gan!Imp_deter), rs_Traza_gan!Imp_deter, 0)
        Traza_Gan(I).Eme_medicas = IIf(Not EsNulo(rs_Traza_gan!Eme_medicas), rs_Traza_gan!Eme_medicas, 0)
        Traza_Gan(I).Seguro_optativo = IIf(Not EsNulo(rs_Traza_gan!Seguro_optativo), rs_Traza_gan!Seguro_optativo, 0)
        Traza_Gan(I).Seguro_retiro = IIf(Not EsNulo(rs_Traza_gan!Seguro_retiro), rs_Traza_gan!Seguro_retiro, 0)
        Traza_Gan(I).Tope_os_priv = IIf(Not EsNulo(rs_Traza_gan!Tope_os_priv), rs_Traza_gan!Tope_os_priv, 0)
        Traza_Gan(I).Empleg = IIf(Not EsNulo(rs_Traza_gan!Empleg), rs_Traza_gan!Empleg, 0)
        
        rs_Traza_gan.MoveNext
    Loop
    
    'Traza_gan - Borro los que voy a recalcular
    StrSql = "DELETE FROM traza_gan WHERE"
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    'StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).Concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    'TRAZA_GAN --->
            
            
    'TRAZA_GAN_ITEM_TOP --->
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "--- Depuracion de traza_gan_item_top ---"
    End If
    StrSql = "SELECT * FROM traza_gan_item_top "
    StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    OpenRecordset StrSql, rs_Traza_gan_items_tope
    I = 0
    Do While Not rs_Traza_gan_items_tope.EOF
        I = I + 1
        Traza_Gan_Item_Top(I).Itenro = IIf(Not EsNulo(rs_Traza_gan_items_tope!Itenro), rs_Traza_gan_items_tope!Itenro, 0)
        Traza_Gan_Item_Top(I).ternro = IIf(Not EsNulo(rs_Traza_gan_items_tope!ternro), rs_Traza_gan_items_tope!ternro, 0)
        Traza_Gan_Item_Top(I).pronro = IIf(Not EsNulo(rs_Traza_gan_items_tope!pronro), rs_Traza_gan_items_tope!pronro, 0)
        Traza_Gan_Item_Top(I).Empresa = IIf(Not EsNulo(rs_Traza_gan_items_tope!Empresa), rs_Traza_gan_items_tope!Empresa, 0)
        Traza_Gan_Item_Top(I).Monto = IIf(Not EsNulo(rs_Traza_gan_items_tope!Monto), rs_Traza_gan_items_tope!Monto, 0)
        Traza_Gan_Item_Top(I).Ddjj = IIf(Not EsNulo(rs_Traza_gan_items_tope!Ddjj), rs_Traza_gan_items_tope!Ddjj, 0)
        Traza_Gan_Item_Top(I).Old_liq = IIf(Not EsNulo(rs_Traza_gan_items_tope!Old_liq), rs_Traza_gan_items_tope!Old_liq, 0)
        Traza_Gan_Item_Top(I).Liq = IIf(Not EsNulo(rs_Traza_gan_items_tope!Liq), rs_Traza_gan_items_tope!Liq, 0)
        Traza_Gan_Item_Top(I).Prorr = IIf(Not EsNulo(rs_Traza_gan_items_tope!Prorr), rs_Traza_gan_items_tope!Prorr, 0)
   
        rs_Traza_gan_items_tope.MoveNext
    Loop
            
    'Traza_gan_item_top - Borro los que voy a recalcular
    StrSql = "DELETE FROM traza_gan_item_top "
    StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    objConn.Execute StrSql, , adExecuteNoRecords
    'TRAZA_GAN_ITEM_TOP --->
        
    
    'DESLIQ --->
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "--- Depuracion de desliq ---"
    End If
    StrSql = "SELECT * FROM desliq WHERE empleado = " & buliq_empleado!ternro
    StrSql = StrSql & " AND dlfecha = " & ConvFecha(buliq_proceso!profecpago)
    OpenRecordset StrSql, rs_Desliq
    I = 0
    Do While Not rs_Desliq.EOF
        I = I + 1
        Desliq(I).Itenro = IIf(Not EsNulo(rs_Desliq!Itenro), rs_Desliq!Itenro, 0)
        Desliq(I).Empleado = IIf(Not EsNulo(rs_Desliq!Empleado), rs_Desliq!Empleado, 0)
        Desliq(I).Dlfecha = IIf(Not EsNulo(rs_Desliq!Dlfecha), rs_Desliq!Dlfecha, 0)
        Desliq(I).pronro = IIf(Not EsNulo(rs_Desliq!pronro), rs_Desliq!pronro, 0)
        Desliq(I).Dlmonto = IIf(Not EsNulo(rs_Desliq!Dlmonto), rs_Desliq!Dlmonto, 0)
        Desliq(I).Dlprorratea = IIf(Not EsNulo(rs_Desliq!Dlprorratea), rs_Desliq!Dlprorratea, 0)
        
        rs_Desliq.MoveNext
    Loop
    
    'Traza_desliq - Borro los que voy a recalcular
    StrSql = "DELETE FROM desliq WHERE empleado = " & buliq_empleado!ternro
    StrSql = StrSql & " AND dlfecha = " & ConvFecha(buliq_proceso!profecpago)
    objConn.Execute StrSql, , adExecuteNoRecords
    'DESLIQ --->
        
        
        
    'FICHARET --->
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "--- Depuracion de ficharet ---"
    End If
    StrSql = "SELECT * FROM ficharet WHERE empleado = " & buliq_empleado!ternro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    OpenRecordset StrSql, rs_Ficharet
    I = 0
    Do While Not rs_Ficharet.EOF
        I = I + 1
        Ficharet(I).Empleado = IIf(Not EsNulo(rs_Ficharet!Empleado), rs_Ficharet!Empleado, 0)
        Ficharet(I).Fecha = IIf(Not EsNulo(rs_Ficharet!Fecha), rs_Ficharet!Fecha, 0)
        Ficharet(I).pronro = IIf(Not EsNulo(rs_Ficharet!pronro), rs_Ficharet!pronro, 0)
        Ficharet(I).importe = IIf(Not EsNulo(rs_Ficharet!importe), rs_Ficharet!importe, 0)
        Ficharet(I).liqsistema = IIf(Not EsNulo(rs_Ficharet!liqsistema), rs_Ficharet!liqsistema, 0)
        
        rs_Ficharet.MoveNext
    Loop
    
    StrSql = "DELETE FROM ficharet "
    StrSql = StrSql & " WHERE pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND empleado =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    'FICHARET --->
        
    '   Depuracion y almacenamiento de temporales
    '-----------------------------------------------------------------------------
    
    
    
    '-----------------------------------------------------------------------------
    '   Ejecucion de la formula original de ganancias
        for_Ganancias_Petroleros = for_Ganancias(NroCab, fec, Monto, Bien)
    '   Ejecucion de la formula original de ganancias
    '-----------------------------------------------------------------------------
    
    '------------------------------------------------------------------------------
    'Restauro todos los valores guardados en temporales
    
    'Traza_gan - Borro los recientemente guardados
    StrSql = "DELETE FROM traza_gan WHERE"
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND ternro =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Traza_gan - inserto los guardados anteriormente
    'FGZ - 27/09/2007 -
    'For I = 1 To 1  'UBound(Traza_Gan)
    For I = 1 To UBound(Traza_Gan)
        If Traza_Gan(I).ternro <> 0 Then
            StrSql = "INSERT INTO traza_gan (pliqnro,concnro,pronro,empresa,fecha_pago,ternro "
            StrSql = StrSql & ",msr,nomsr,nogan,jubilacion,osocial,cuota_medico, prima_seguro"
            StrSql = StrSql & ",Sepelio,Estimados,Otras,Donacion,Dedesp,Noimpo,Car_flia,Conyuge"
            StrSql = StrSql & ",Hijo,Otras_cargas,Retenciones,Promo,saldo,sindicato,ret_mes"
            StrSql = StrSql & ",Mon_conyuge,Mon_hijo,Mon_otras,Viaticos,Amortizacion"
            If Traza_Gan(I).Entidad1 <> "" Then
                StrSql = StrSql & ",Entidad1"
            End If
            If Traza_Gan(I).Entidad2 <> "" Then
                StrSql = StrSql & ",Entidad2"
            End If
            If Traza_Gan(I).Entidad3 <> "" Then
                StrSql = StrSql & ",Entidad3"
            End If
            If Traza_Gan(I).Entidad4 <> "" Then
                StrSql = StrSql & ",Entidad4"
            End If
            If Traza_Gan(I).Entidad5 <> "" Then
                StrSql = StrSql & ",Entidad5"
            End If
            If Traza_Gan(I).Entidad6 <> "" Then
                StrSql = StrSql & ",Entidad6"
            End If
            If Traza_Gan(I).Entidad7 <> "" Then
                StrSql = StrSql & ",Entidad7"
            End If
            If Traza_Gan(I).Entidad8 <> "" Then
                StrSql = StrSql & ",Entidad8"
            End If
            If Traza_Gan(I).Entidad9 <> "" Then
                StrSql = StrSql & ",Entidad9"
            End If
            If Traza_Gan(I).Entidad2 <> "" Then
                StrSql = StrSql & ",Entidad10"
            End If
            If Traza_Gan(I).Entidad11 <> "" Then
                StrSql = StrSql & ",Entidad11"
            End If
            If Traza_Gan(I).Entidad12 <> "" Then
                StrSql = StrSql & ",Entidad12"
            End If
            If Traza_Gan(I).Entidad13 <> "" Then
                StrSql = StrSql & ",Entidad13"
            End If
            If Traza_Gan(I).Entidad14 <> "" Then
                StrSql = StrSql & ",Entidad14"
            End If
            If Traza_Gan(I).Cuit_entidad1 <> "" Then
                StrSql = StrSql & ",Cuit_entidad1"
            End If
            If Traza_Gan(I).Cuit_entidad2 <> "" Then
                StrSql = StrSql & ",Cuit_entidad2"
            End If
            If Traza_Gan(I).Cuit_entidad3 <> "" Then
                StrSql = StrSql & ",Cuit_entidad3"
            End If
            If Traza_Gan(I).Cuit_entidad4 <> "" Then
                StrSql = StrSql & ",Cuit_entidad4"
            End If
            If Traza_Gan(I).Cuit_entidad5 <> "" Then
                StrSql = StrSql & ",Cuit_entidad5"
            End If
            If Traza_Gan(I).Cuit_entidad6 <> "" Then
                StrSql = StrSql & ",Cuit_entidad6"
            End If
            If Traza_Gan(I).Cuit_entidad7 <> "" Then
                StrSql = StrSql & ",Cuit_entidad7"
            End If
            If Traza_Gan(I).Cuit_entidad8 <> "" Then
                StrSql = StrSql & ",Cuit_entidad8"
            End If
            If Traza_Gan(I).Cuit_entidad9 <> "" Then
                StrSql = StrSql & ",Cuit_entidad9"
            End If
            If Traza_Gan(I).Cuit_entidad10 <> "" Then
                StrSql = StrSql & ",Cuit_entidad10"
            End If
            If Traza_Gan(I).Cuit_entidad11 <> "" Then
                StrSql = StrSql & ",Cuit_entidad11"
            End If
            If Traza_Gan(I).Cuit_entidad12 <> "" Then
                StrSql = StrSql & ",Cuit_entidad12"
            End If
            If Traza_Gan(I).Cuit_entidad13 <> "" Then
                StrSql = StrSql & ",Cuit_entidad13"
            End If
            If Traza_Gan(I).Cuit_entidad14 <> "" Then
                StrSql = StrSql & ",Cuit_entidad14"
            End If
            If Traza_Gan(I).Monto_entidad1 <> 0 Then
                StrSql = StrSql & ",monto_entidad1"
            End If
            If Traza_Gan(I).Monto_entidad2 <> 0 Then
                StrSql = StrSql & ",monto_entidad2"
            End If
            If Traza_Gan(I).Monto_entidad3 <> 0 Then
                StrSql = StrSql & ",monto_entidad3"
            End If
            If Traza_Gan(I).Monto_entidad4 <> 0 Then
                StrSql = StrSql & ",monto_entidad4"
            End If
            If Traza_Gan(I).Monto_entidad5 <> 0 Then
                StrSql = StrSql & ",monto_entidad5"
            End If
            If Traza_Gan(I).Monto_entidad6 <> 0 Then
                StrSql = StrSql & ",monto_entidad6"
            End If
            If Traza_Gan(I).Monto_entidad7 <> 0 Then
                StrSql = StrSql & ",monto_entidad7"
            End If
            If Traza_Gan(I).Monto_entidad8 <> 0 Then
                StrSql = StrSql & ",monto_entidad8"
            End If
            If Traza_Gan(I).Monto_entidad9 <> 0 Then
                StrSql = StrSql & ",monto_entidad9"
            End If
            If Traza_Gan(I).Monto_entidad10 <> 0 Then
                StrSql = StrSql & ",monto_entidad10"
            End If
            If Traza_Gan(I).Monto_entidad11 <> 0 Then
                StrSql = StrSql & ",monto_entidad11"
            End If
            If Traza_Gan(I).Monto_entidad12 <> 0 Then
                StrSql = StrSql & ",monto_entidad12"
            End If
            If Traza_Gan(I).Monto_entidad13 <> 0 Then
                StrSql = StrSql & ",monto_entidad13"
            End If
            If Traza_Gan(I).Monto_entidad14 <> 0 Then
                StrSql = StrSql & ",monto_entidad14"
            End If
            StrSql = StrSql & ",Ganimpo,Ganneta"
            If Traza_Gan(I).Total_entidad1 <> 0 Then
                StrSql = StrSql & ",total_entidad1"
            End If
            If Traza_Gan(I).Total_entidad2 <> 0 Then
                StrSql = StrSql & ",total_entidad2"
            End If
            If Traza_Gan(I).Total_entidad3 <> 0 Then
                StrSql = StrSql & ",total_entidad3"
            End If
            If Traza_Gan(I).Total_entidad4 <> 0 Then
                StrSql = StrSql & ",total_entidad4"
            End If
            If Traza_Gan(I).Total_entidad5 <> 0 Then
                StrSql = StrSql & ",total_entidad5"
            End If
            If Traza_Gan(I).Total_entidad6 <> 0 Then
                StrSql = StrSql & ",total_entidad6"
            End If
            If Traza_Gan(I).Total_entidad7 <> 0 Then
                StrSql = StrSql & ",total_entidad7"
            End If
            If Traza_Gan(I).Total_entidad8 <> 0 Then
                StrSql = StrSql & ",total_entidad8"
            End If
            If Traza_Gan(I).Total_entidad9 <> 0 Then
                StrSql = StrSql & ",total_entidad9"
            End If
            If Traza_Gan(I).Total_entidad10 <> 0 Then
                StrSql = StrSql & ",total_entidad10"
            End If
            If Traza_Gan(I).Total_entidad11 <> 0 Then
                StrSql = StrSql & ",total_entidad11"
            End If
            If Traza_Gan(I).Total_entidad12 <> 0 Then
                StrSql = StrSql & ",total_entidad12"
            End If
            If Traza_Gan(I).Total_entidad13 <> 0 Then
                StrSql = StrSql & ",total_entidad13"
            End If
            If Traza_Gan(I).Total_entidad14 <> 0 Then
                StrSql = StrSql & ",total_entidad14"
            End If
            StrSql = StrSql & ",Imp_deter,Eme_medicas,Seguro_optativo,Seguro_retiro,Tope_os_priv,empleg"
            StrSql = StrSql & ") VALUES ("
            
            StrSql = StrSql & Traza_Gan(I).PliqNro
            StrSql = StrSql & "," & Traza_Gan(I).concnro
            StrSql = StrSql & "," & Traza_Gan(I).pronro
            StrSql = StrSql & "," & Traza_Gan(I).Empresa
            StrSql = StrSql & "," & ConvFecha(Traza_Gan(I).Fecha_pago)
            StrSql = StrSql & "," & Traza_Gan(I).ternro
            StrSql = StrSql & "," & Traza_Gan(I).Msr
            StrSql = StrSql & "," & Traza_Gan(I).Nomsr
            StrSql = StrSql & "," & Traza_Gan(I).Nogan
            StrSql = StrSql & "," & Traza_Gan(I).Jubilacion
            StrSql = StrSql & "," & Traza_Gan(I).Osocial
            StrSql = StrSql & "," & Traza_Gan(I).Cuota_medico
            StrSql = StrSql & "," & Traza_Gan(I).Prima_seguro
            StrSql = StrSql & "," & Traza_Gan(I).Sepelio
            StrSql = StrSql & "," & Traza_Gan(I).Estimados
            StrSql = StrSql & "," & Traza_Gan(I).Otras
            StrSql = StrSql & "," & Traza_Gan(I).Donacion
            StrSql = StrSql & "," & Traza_Gan(I).Dedesp
            StrSql = StrSql & "," & Traza_Gan(I).Noimpo
            StrSql = StrSql & "," & Traza_Gan(I).Car_flia
            StrSql = StrSql & "," & Traza_Gan(I).conyuge
            StrSql = StrSql & "," & Traza_Gan(I).Hijo
            StrSql = StrSql & "," & Traza_Gan(I).Otras_cargas
            StrSql = StrSql & "," & Traza_Gan(I).Retenciones
            StrSql = StrSql & "," & Traza_Gan(I).Promo
            StrSql = StrSql & "," & Traza_Gan(I).Saldo
            StrSql = StrSql & "," & Traza_Gan(I).Sindicato
            StrSql = StrSql & "," & Traza_Gan(I).Ret_Mes
            StrSql = StrSql & "," & Traza_Gan(I).Mon_conyuge
            StrSql = StrSql & "," & Traza_Gan(I).Mon_hijo
            StrSql = StrSql & "," & Traza_Gan(I).Mon_otras
            StrSql = StrSql & "," & Traza_Gan(I).Viaticos
            StrSql = StrSql & "," & Traza_Gan(I).Amortizacion
            If Traza_Gan(I).Entidad1 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad1 & "'"
            End If
            If Traza_Gan(I).Entidad2 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad2 & "'"
            End If
            If Traza_Gan(I).Entidad3 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad3 & "'"
            End If
            If Traza_Gan(I).Entidad4 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad4 & "'"
            End If
            If Traza_Gan(I).Entidad5 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad5 & "'"
            End If
            If Traza_Gan(I).Entidad6 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad6 & "'"
            End If
            If Traza_Gan(I).Entidad7 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad7 & "'"
            End If
            If Traza_Gan(I).Entidad8 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad8 & "'"
            End If
            If Traza_Gan(I).Entidad9 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad9 & "'"
            End If
            If Traza_Gan(I).Entidad10 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad10 & "'"
            End If
            If Traza_Gan(I).Entidad11 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad11 & "'"
            End If
            If Traza_Gan(I).Entidad12 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad12 & "'"
            End If
            If Traza_Gan(I).Entidad13 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad13 & "'"
            End If
            If Traza_Gan(I).Entidad14 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Entidad14 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad1 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad1 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad2 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad2 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad3 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad3 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad4 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad4 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad5 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad5 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad6 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad6 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad7 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad7 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad8 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad8 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad9 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad9 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad10 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad10 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad11 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad11 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad12 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad12 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad13 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad13 & "'"
            End If
            If Traza_Gan(I).Cuit_entidad14 <> "" Then
                StrSql = StrSql & ",'" & Traza_Gan(I).Cuit_entidad14 & "'"
            End If
            If Traza_Gan(I).Monto_entidad1 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad1
            End If
            If Traza_Gan(I).Monto_entidad2 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad2
            End If
            If Traza_Gan(I).Monto_entidad3 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad3
            End If
            If Traza_Gan(I).Monto_entidad4 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad4
            End If
            If Traza_Gan(I).Monto_entidad5 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad5
            End If
            If Traza_Gan(I).Monto_entidad6 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad6
            End If
            If Traza_Gan(I).Monto_entidad7 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad7
            End If
            If Traza_Gan(I).Monto_entidad8 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad8
            End If
            If Traza_Gan(I).Monto_entidad9 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad9
            End If
            If Traza_Gan(I).Monto_entidad10 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad10
            End If
            If Traza_Gan(I).Monto_entidad11 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad11
            End If
            If Traza_Gan(I).Monto_entidad12 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad12
            End If
            If Traza_Gan(I).Monto_entidad13 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad13
            End If
            If Traza_Gan(I).Monto_entidad14 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Monto_entidad14
            End If
            StrSql = StrSql & "," & Traza_Gan(I).Ganimpo
            StrSql = StrSql & "," & Traza_Gan(I).Ganneta
            If Traza_Gan(I).Total_entidad1 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad1
            End If
            If Traza_Gan(I).Total_entidad2 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad2
            End If
            If Traza_Gan(I).Total_entidad3 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad3
            End If
            If Traza_Gan(I).Total_entidad4 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad4
            End If
            If Traza_Gan(I).Total_entidad5 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad5
            End If
            If Traza_Gan(I).Total_entidad6 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad6
            End If
            If Traza_Gan(I).Total_entidad7 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad7
            End If
            If Traza_Gan(I).Total_entidad8 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad8
            End If
            If Traza_Gan(I).Total_entidad9 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad9
            End If
            If Traza_Gan(I).Total_entidad10 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad10
            End If
            If Traza_Gan(I).Total_entidad11 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad11
            End If
            If Traza_Gan(I).Total_entidad12 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad12
            End If
            If Traza_Gan(I).Total_entidad13 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad13
            End If
            If Traza_Gan(I).Total_entidad14 <> 0 Then
                StrSql = StrSql & "," & Traza_Gan(I).Total_entidad14
            End If
            StrSql = StrSql & "," & Traza_Gan(I).Imp_deter
            StrSql = StrSql & "," & Traza_Gan(I).Eme_medicas
            StrSql = StrSql & "," & Traza_Gan(I).Seguro_optativo
            StrSql = StrSql & "," & Traza_Gan(I).Seguro_retiro
            StrSql = StrSql & "," & Traza_Gan(I).Tope_os_priv
            StrSql = StrSql & "," & Traza_Gan(I).Empleg
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Next I
    
    
    'Traza_gan_item_top - Borro los recientemente guardados
    StrSql = "DELETE FROM traza_gan_item_top "
    StrSql = StrSql & " WHERE ternro =" & buliq_empleado!ternro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Traza_gan - inserto los guardados anteriormente
    For I = 1 To 1  'UBound(Traza_Gan_item_top)
        If Traza_Gan_Item_Top(I).Itenro <> 0 Then
            StrSql = "INSERT INTO traza_gan_item_top (ternro,pronro,monto,empresa,itenro"
            If Not EsNulo(Traza_Gan_Item_Top(I).Ddjj) Then
                StrSql = StrSql & ",ddjj"
            End If
            If Not EsNulo(Traza_Gan_Item_Top(I).Old_liq) Then
                StrSql = StrSql & ",old_liq"
            End If
            If Not EsNulo(Traza_Gan_Item_Top(I).Liq) Then
                StrSql = StrSql & ",liq"
            End If
            If Not EsNulo(Traza_Gan_Item_Top(I).Prorr) Then
                StrSql = StrSql & ",prorr"
            End If
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & Traza_Gan_Item_Top(I).ternro
            StrSql = StrSql & "," & Traza_Gan_Item_Top(I).pronro
            StrSql = StrSql & "," & Traza_Gan_Item_Top(I).Monto
            StrSql = StrSql & "," & Traza_Gan_Item_Top(I).Empresa
            StrSql = StrSql & "," & Traza_Gan_Item_Top(I).Itenro
            If Not EsNulo(Traza_Gan_Item_Top(I).Ddjj) Then
                StrSql = StrSql & "," & Traza_Gan_Item_Top(I).Ddjj
            End If
            If Not EsNulo(Traza_Gan_Item_Top(I).Old_liq) Then
                StrSql = StrSql & "," & Traza_Gan_Item_Top(I).Old_liq
            End If
            If Not EsNulo(Traza_Gan_Item_Top(I).Liq) Then
                StrSql = StrSql & "," & Traza_Gan_Item_Top(I).Liq
            End If
            If Not EsNulo(Traza_Gan_Item_Top(I).Prorr) Then
                StrSql = StrSql & "," & Traza_Gan_Item_Top(I).Prorr
            End If
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Next I
    
    
    'Traza_desliq - Borro los recientemente guardados
    StrSql = "DELETE FROM desliq WHERE empleado = " & buliq_empleado!ternro
    StrSql = StrSql & " AND dlfecha = " & ConvFecha(buliq_proceso!profecpago)
    objConn.Execute StrSql, , adExecuteNoRecords
    
    For I = 1 To UBound(Desliq)
        If Desliq(I).Itenro <> 0 Then
        
            StrSql = "INSERT INTO desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro"
            StrSql = StrSql & " ) VALUES ("
            StrSql = StrSql & Desliq(I).Empleado
            StrSql = StrSql & "," & ConvFecha(Desliq(I).Dlfecha)
            StrSql = StrSql & "," & Desliq(I).pronro
            StrSql = StrSql & "," & Desliq(I).Dlmonto
            StrSql = StrSql & "," & CInt(Desliq(I).Dlprorratea)
            StrSql = StrSql & "," & Desliq(I).Itenro
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    Next I
    
    
    'ficharet - Borro los recientemente calculados
    StrSql = "DELETE FROM ficharet "
    StrSql = StrSql & " WHERE pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND empleado =" & buliq_empleado!ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    Ret_Actual = 0
    For I = 1 To 1 'UBound(ficharet)
        If Ficharet(I).Empleado <> 0 Then
            StrSql = "INSERT INTO ficharet (empleado,fecha,pronro,importe,liqsistema"
            StrSql = StrSql & ") VALUES ("
            StrSql = StrSql & Ficharet(I).Empleado
            StrSql = StrSql & "," & ConvFecha(Ficharet(I).Fecha)
            StrSql = StrSql & "," & Ficharet(I).pronro
            StrSql = StrSql & "," & Ficharet(I).importe
            StrSql = StrSql & ",-1"
            StrSql = StrSql & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
            Ret_Actual = Ficharet(I).importe
        End If
    Next I
    'Restauro todos los valores guardados en temporales
    '------------------------------------------------------------------------------
    
    
    '------------------------------------------------------------------------------
    'Ajuste de ficharet
    Ret_Real = 0
    Ret_Sin_Ext = 0
    If Ajusta_Ficharet Then
        'Calcular la diferencia entre lo retenido real (Ganancias Original) y lo que deberia haber retenido por este concepto (Ganancias sin extras)

        Ret_Mes = Month(buliq_proceso!profecpago)
        Ret_Ano = Year(buliq_proceso!profecpago)
        
        'levanto todas las ficharet del tercero y hago la pregunta dentro del loop
        StrSql = "SELECT * FROM ficharet " & _
                 " WHERE empleado =" & buliq_empleado!ternro
        OpenRecordset StrSql, rs_Ficharet
        Do While Not rs_Ficharet.EOF
            If (Month(rs_Ficharet!Fecha) <= Ret_Mes) And (Year(rs_Ficharet!Fecha) = Ret_Ano) Then
                Ret_Real = Ret_Real + rs_Ficharet!importe
            End If
            rs_Ficharet.MoveNext
        Loop
        'Le saco la retencion del proceso actual
        Ret_Real = Ret_Real - Ret_Actual

        'Busco el valor del concepto en el año
        StrSql = "SELECT sum(dlimonto) as retencion FROM periodo "
        StrSql = StrSql & " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro "
        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro "
        StrSql = StrSql & " WHERE cabliq.empleado = " & buliq_empleado!ternro
        StrSql = StrSql & " AND detliq.concnro = " & Buliq_Concepto(Concepto_Actual).concnro
        StrSql = StrSql & " AND periodo.pliqanio = " & Ret_Ano
        StrSql = StrSql & " AND periodo.pliqmes <= " & Ret_Mes
        OpenRecordset StrSql, rs_Periodos
        If Not rs_Periodos.EOF Then
            'Ret_Sin_Ext = IIf(Not EsNulo(rs_Periodos!Retencion), Abs(rs_Periodos!Retencion), 0)
            Ret_Sin_Ext = IIf(Not EsNulo(rs_Periodos!Retencion), (rs_Periodos!Retencion), 0)
        End If
    End If
    'Ajuste de ficharet
    '------------------------------------------------------------------------------
    
    If CBool(USA_DEBUG) Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 3) & "--- Monto antes del ajuste    --- " & Monto
        Flog.writeline Espacios(Tabulador * 3) & "--- Ret_Real                  --- " & Ret_Real
        Flog.writeline Espacios(Tabulador * 3) & "--- Ret_Sin_Ext               --- " & Ret_Sin_Ext
        Flog.writeline
    End If
    
    
    exito = Bien
    If Monto = 0 Then
        for_Ganancias_Petroleros = 0
    Else
        for_Ganancias_Petroleros = Monto - Ret_Real - Ret_Sin_Ext
    End If
    
    
'Cierro todo y libero
    If rs_Desliq.State = adStateOpen Then rs_Desliq.Close
    Set rs_Desliq = Nothing
    If rs_Traza_gan_items_tope.State = adStateOpen Then rs_Traza_gan_items_tope.Close
    Set rs_Traza_gan_items_tope = Nothing
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    Set rs_Traza_gan = Nothing
    If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
    Set rs_wf_tpa = Nothing
    If rs_Periodos.State = adStateOpen Then rs_Periodos.Close
    Set rs_Periodos = Nothing
End Function


Public Function For_Sac_No_Remu() As Double

Const c_AcuBaseSac = 53
Const c_AcuNoRemu = 68

Dim v_AcuBaseSac As Double
Dim v_AcuNoRemu As Double

Dim Encontro_AcuBaseSac As Boolean
Dim Encontro_AcuNoRemu As Boolean
Dim MesMax As Integer
Dim AnioMax As Integer
Dim MesDeInicioSemestre As Integer 'Mes de Inicio
Dim CantMeses As Integer
Dim MesHasta As Integer
Dim AnioHasta As Integer
Dim Cantidad As Double
Dim CantAnios As Integer
Dim MesesFuera As Integer
Dim Aux_MesDesde As Integer
Dim MontoMax As Double

Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset

    exito = False
    Encontro_AcuBaseSac = False
    Encontro_AcuNoRemu = False
    MontoMax = 0
    
    StrSql = "SELECT * FROM " & TTempWF_tpa
    OpenRecordset StrSql, rs_wf_tpa
    
    
    Do While Not rs_wf_tpa.EOF
        Select Case rs_wf_tpa!tipoparam
        Case c_AcuBaseSac:
            v_AcuBaseSac = rs_wf_tpa!Valor
            Encontro_AcuBaseSac = True
        Case c_AcuNoRemu:
            v_AcuNoRemu = rs_wf_tpa!Valor
            Encontro_AcuNoRemu = True
        Case Else
        End Select
        
        rs_wf_tpa.MoveNext
    Loop

    ' si no se obtuvieron los parametros, ==> Error.
    If Not Encontro_AcuBaseSac Or Not Encontro_AcuNoRemu Then
        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 4) & "No se encontraron todos los parametros."
        Exit Function
    End If
    
    
    If buliq_periodo!pliqmes > 6 Then
        MesDeInicioSemestre = 7 'segundo semestre
        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 4) & "Calcula Segundo Semestre."
    Else
        MesDeInicioSemestre = 1 'primer semestre
        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 4) & "Calcula Primer Semestre."
    End If
    
    'Calculo la cantidad de meses fuera del periodo
    Aux_MesDesde = MesDeInicioSemestre
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    StrSql = StrSql & " ORDER BY altfec"
    OpenRecordset StrSql, rs_Fases
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If rs_Fases!altfec > C_Date("01/" & MesDeInicioSemestre & "/" & buliq_periodo!pliqanio) Then
            Aux_MesDesde = Month(rs_Fases!altfec)
        End If
    End If
    MesesFuera = Aux_MesDesde - MesDeInicioSemestre

    
    If buliq_periodo!pliqmes > 6 Then
        If MesDeInicioSemestre <= buliq_periodo!pliqmes Then
            CantMeses = buliq_periodo!pliqmes - MesDeInicioSemestre + 1
        Else
            CantMeses = buliq_periodo!pliqmes
        End If
        MesHasta = buliq_periodo!pliqmes
        AnioHasta = buliq_periodo!pliqanio
    Else
        If MesDeInicioSemestre <= buliq_periodo!pliqmes Then
            CantMeses = buliq_periodo!pliqmes - MesDeInicioSemestre + 1
        Else
            CantMeses = buliq_periodo!pliqmes
        End If
        MesHasta = buliq_periodo!pliqmes
        AnioHasta = buliq_periodo!pliqanio
    End If
    If MesDeInicioSemestre >= 6 Then
        CantMeses = CantMeses + IIf(MesDeInicioSemestre > buliq_periodo!pliqmes, (12 - MesDeInicioSemestre) + 1, 0)
    Else
        CantMeses = CantMeses + IIf(MesDeInicioSemestre > buliq_periodo!pliqmes, (6 - MesDeInicioSemestre) + 1, 0)
    End If
    If MesDeInicioSemestre = buliq_periodo!pliqmes Then
        CantMeses = 1
    End If
    
    If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 4) & "Cantidad de Meses a tener en cuenta " & CantMeses
    
    Call AM_Max_Mes(v_AcuBaseSac, MesHasta, AnioHasta, CantMeses, CantAnios, True, Valor, Cantidad, MesMax, AnioMax, False, True, True)
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Mejor mes Base SAC " & MesMax & " " & AnioMax
        Flog.writeline Espacios(Tabulador * 4) & "Buscando Acumulador No Remunerativo para mejor mes."
    End If
    
    'Busco el monto del mejor mes
    StrSql = "SELECT ammonto FROM acu_mes " & _
             " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
             " WHERE ternro = " & buliq_empleado!ternro & _
             " AND acu_mes.acunro =" & v_AcuNoRemu & _
             " AND (" & AnioMax & " = amanio AND ammes = " & MesMax & ")"
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        MontoMax = IIf(Not EsNulo(rs_Fases!ammonto), rs_Fases!ammonto, 0)
    End If
    rs_Fases.Close


    'Si es desde el mes actual ==> busco el acu_liq de este proceso
    If ((buliq_periodo!pliqmes = MesMax) And (buliq_periodo!pliqanio = AnioMax)) Then
    
        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(v_AcuNoRemu)) Then
            MontoMax = objCache_Acu_Liq_Monto.Valor(CStr(v_AcuNoRemu)) + MontoMax
        End If
        
    End If
    
    Flog.writeline Espacios(Tabulador * 4) & "Monto encontrado" & MontoMax
    
    For_Sac_No_Remu = MontoMax
    exito = True
    
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
Set rs_wf_tpa = Nothing

End Function

