Attribute VB_Name = "Mdlformulas"
Option Explicit

Public Type TTraza_Gan
    PliqNro As Long
    ConcNro As Long
    Empresa As Long
    Fecha_pago As Date
    Ternro As Long
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
    Ternro  As Long
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

'VAriables
Dim pos1
Dim pos2


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
' Inicio Formulas de Chacomer (PY)
    Case "FOR_COMISION"
        EjecutarFormulaNoConfigurable = for_Comision
'-------------------------------------------------
' Fin Formulas de Chacomer (PY)
'-------------------------------------------------
' Inicio Formulas de Raffo
    Case "FOR_COMISIONES"
        EjecutarFormulaNoConfigurable = for_Comisiones
'-------------------------------------------------
' Fin Formulas de Raffo
        
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
    Case "FOR_GANANCIAS2013":
        EjecutarFormulaDeSistema = for_Ganancias2013(NroCab, fec, Monto, Bien)
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
    Case "FOR_PREMIO_SEMESTRE":
       EjecutarFormulaDeSistema = For_Premio_Semestre
    Case "FOR_IRRF":
        EjecutarFormulaDeSistema = for_IRRF(NroCab, fec, Monto, Bien)
    Case "FOR_GANANCIAS2017":
        EjecutarFormulaDeSistema = for_Ganancias2017(NroCab, fec, Monto, Bien)
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
    If (v_nroConce = Arr_conceptos(Concepto_Actual).ConcNro) Or (v_nroConce = 0) Then
        
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
                v_concnro = rs_Conc!ConcNro
              Else
                Exit Function
        End If
        
        If v_nroitera = 1 Then
           v_netofijo = v_netopactado
           v_nroConce = Arr_conceptos(Concepto_Actual).ConcNro
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
            StrSql = " SELECT * FROM sim_novemp " & _
                     " WHERE concnro =" & v_concnro & _
                     " AND tpanro =" & c_tpanro & _
                     " AND empleado =" & buliq_empleado!Ternro & _
                     " AND ((nevigencia = -1 " & _
                     " AND nedesde <= " & ConvFecha(Fecha_Fin) & _
                     " AND (nehasta >= " & ConvFecha(Fecha_Inicio) & _
                     " OR nehasta is null )) " & _
                     " OR nevigencia = 0)"
            OpenRecordset StrSql, rs_Conc
            
            If rs_Conc.EOF Then
               ' Inserto novedad
                 StrSql = "INSERT INTO sim_novemp (empleado, concnro, tpanro, nevalor,nevigencia,nedesde,nehasta,pronro ) VALUES (" & _
                 buliq_empleado!Ternro & "," & v_concnro & "," & c_tpanro & "," & v_diferencia & ", -1" & _
                 "," & ConvFecha(Fecha_Inicio) & _
                 "," & ConvFecha(Fecha_Fin) & _
                 "," & buliq_proceso!pronro & _
                 " )"
            objConn.Execute StrSql, , adExecuteNoRecords
            Else ' Actualizo la novedad
                 StrSql = "UPDATE sim_novemp SET nevalor = nevalor + " & v_diferencia & _
                          " WHERE concnro = " & v_concnro & _
                          " AND tpanro = " & c_tpanro & _
                          " AND empleado = " & buliq_empleado!Ternro & _
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
'Ultima Mod :   FGZ 07/01/2013
'               Impuestos y debitos Bancarios va como Promocion
'               Se agregó el ITEM 23
'Ultima Mod :   FGZ 15/01/2013
'               Impuestos y debitos Bancarios va como Promocion
'               Se cambió el ITEM 23 por el 56
' ---------------------------------------------------------------------------------------------
Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim p_prorratea As Integer  'Si prorratea o no para liq. finales
Dim p_sinprorrateo As Integer  'Indica que nunca prorratea

'Variables Locales
Dim Devuelve As Double
Dim Tope_Gral As Double
Dim Neto As Double
Dim prorratea As Double
Dim sinprorrateo As Double
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
Dim I As Long
Dim j As Integer
Dim Texto As String

'Vectores para manejar el proceso
Dim Items_DDJJ(100) As Double
Dim Items_LIQ(100) As Double
Dim Items_PRORR(100) As Double
Dim Items_PRORR_CUOTA(100) As Double
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
p_sinprorrateo = 1006


'FGZ - 19/04/2004
Dim Total_Empresa As Double
Dim Tope As Integer
'Dim rs_Rep19 As New ADODB.Recordset
Dim rs_Traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Double
Dim Cuota As Double

Total_Empresa = 0
Tope = 10

Descuentos = 0
' Primero limpio la traza
StrSql = "DELETE FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords

' Guardo la traza de Ganancia en traza_gan para utilizarla luego cuando se genere el reporte
StrSql = "INSERT INTO sim_traza_gan (pliqnro,pronro,concnro,fecha_pago,empresa,ternro,empleg) VALUES (" & _
         buliq_periodo!PliqNro & "," & _
         buliq_proceso!pronro & "," & _
         Buliq_Concepto(Concepto_Actual).ConcNro & "," & _
         ConvFecha(buliq_proceso!profecpago) & "," & _
         NroEmp & "," & _
         buliq_empleado!Ternro & "," & _
         buliq_empleado!Empleg & _
         ")"
objConn.Execute StrSql, , adExecuteNoRecords

'FGZ - 22/12/2004
'Leo la tabla
StrSql = "SELECT * FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
OpenRecordset StrSql, rs_Traza_gan
    

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).ConcNro)
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
ini_anyo_ret = C_Date("01/01/" & Ret_Ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro
sinprorrateo = 0

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
    Case p_sinprorrateo:
        sinprorrateo = rs_wf_tpa!Valor
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
      Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_Neto, "El Neto del Mes es negativo, se seteara en cero.", Neto)
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
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Neto del Mes", Neto)
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
  'FGZ - 07/01/2013 -----------------------------------------------------
  'Se agregó el ITEM 23
  'FGZ - 15/01/2013 -----------------------------------------------------
  'Se cambió el ITEM 23 por 56
  'If (rs_Item!Itenro = 29 Or rs_Item!Itenro = 55) And Ret_Mes = 12 Then
  'If (rs_Item!Itenro = 29 Or rs_Item!Itenro = 55 Or rs_Item!Itenro = 23) And Ret_Mes = 12 Then
  If (rs_Item!Itenro = 29 Or rs_Item!Itenro = 55 Or rs_Item!Itenro = 56) And Ret_Mes = 12 Then
        StrSql = "SELECT * FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                 " AND desano=" & Ret_Ano & _
                 " AND itenro = " & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
        If Not rs_Desmen.EOF Then
            If rs_Item!Itenro = 29 Then
                val_impdebitos = rs_Desmen!desmondec * 0.34
            Else
                'If rs_Item!Itenro = 23 Then
                If rs_Item!Itenro = 56 Then
                    val_impdebitos = rs_Desmen!desmondec
                Else
                    val_impdebitos = rs_Desmen!desmondec * 0.17
                End If
           End If
        End If
        rs_Desmen.Close
  'FGZ - 07/01/2013 -----------------------------------------------------
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                 " AND desano=" & Ret_Ano & _
                 " AND itenro = " & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
        
        Do While Not rs_Desmen.EOF
            If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                If rs_Item!Itenro = 3 Then
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                
                Else
                    If rs_Desmen!desmenprorra = 0 Then 'no es parejito
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + rs_Desmen!desmondec
                    Else
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                    End If
                End If
                    
                    
                    'FGZ - 19/04/2004
                    If rs_Item!Itenro <= 4 Then
                        If Not EsNulo(rs_Desmen!descuit) Then
                            I = 11
                            If Not EsNulo(rs_Traza_gan!Cuit_entidad11) Then
                                Distinto = rs_Traza_gan!Cuit_entidad11 <> rs_Desmen!descuit
                            End If
                            Do While (I <= Tope) And Distinto
                                I = I + 1
                                Select Case I
                                Case 11:
                                    Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad11), 0, rs_Traza_gan!Cuit_entidad11) <> rs_Desmen!descuit
                                Case 12:
                                    Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad12), 0, rs_Traza_gan!Cuit_entidad12) <> rs_Desmen!descuit
                                Case 13:
                                    Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad13), 0, rs_Traza_gan!Cuit_entidad13) <> rs_Desmen!descuit
                                Case 14:
                                    Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad14), 0, rs_Traza_gan!Cuit_entidad14) <> rs_Desmen!descuit
                                End Select
                            Loop
                          
                            If I > Tope And I <= 14 Then
                                StrSql = "UPDATE sim_traza_gan SET "
                                StrSql = StrSql & " cuit_entidad" & I & "='" & rs_Desmen!descuit & "',"
                                StrSql = StrSql & " entidad" & I & "='" & rs_Desmen!DesRazsoc & "',"
                                StrSql = StrSql & " monto_entidad" & I & "=" & rs_Desmen!desmondec
                                StrSql = StrSql & " WHERE "
                                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'FGZ - 22/12/2004
                                'Leo la tabla
                                StrSql = "SELECT * FROM sim_traza_gan WHERE "
                                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                'StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                OpenRecordset StrSql, rs_Traza_gan
                                
                                
                                Tope = Tope + 1
                            Else
                                If I = 15 Then
                                    Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!Empleg ' empleado.empleg
                                Else
                                    StrSql = "UPDATE sim_traza_gan SET "
                                    StrSql = StrSql & " monto_entidad" & I & "= monto_entidad" & I & " + " & rs_Desmen!desmondec
                                    StrSql = StrSql & " WHERE "
                                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    'FGZ - 22/12/2004
                                    'Leo la tabla
                                    StrSql = "SELECT * FROM sim_traza_gan WHERE "
                                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    'StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                    OpenRecordset StrSql, rs_Traza_gan
                                End If
                            End If
                        Else
                            Total_Empresa = Total_Empresa + rs_Desmen!desmondec
                        End If
                    End If
                    'FGZ - 19/04/2004
                    
                End If
            
            
            rs_Desmen.MoveNext
        Loop
        
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!Ternro & _
                 " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                 " AND dlfecha <= " & ConvFecha(fin_mes_ret)
        OpenRecordset StrSql, rs_Desliq
        If rs_Desliq.EOF Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "No hay datos de liquidaciones anteriores (desliq)"
            End If
        End If
        Do While Not rs_Desliq.EOF
            Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
            'Si el desliq prorratea debo proporcionarlo
            If CBool(rs_Desliq!Dlprorratea) Then
                Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                Cuota = IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) - (Cuota + ((rs_Desliq!Dlmonto) - (IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)), rs_Desliq!Dlmonto))))
            End If
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
        
                If CBool(rs_itemacum!itaprorratea) And (sinprorrateo = 0) Then
                    If CBool(rs_itemacum!itasigno) Then
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + Aux_Acu_Monto
                        Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                    Else
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - Aux_Acu_Monto
                        Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
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
                 " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                 " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_Item!Itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) And (sinprorrateo = 0) Then
                If CBool(rs_itemconc!itcsigno) Then
                    Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + rs_itemconc!dlimonto
                    Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                Else
                    Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - rs_itemconc!dlimonto
                    Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
        StrSql = "SELECT * FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!Ternro & _
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
                 " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                 " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                    StrSql = "UPDATE sim_traza_gan SET "
                    StrSql = StrSql & " cuit_entidad9 ='" & rs_Desmen!descuit & "',"
                    StrSql = StrSql & " entidad9='" & rs_Desmen!DesRazsoc & "',"
                    StrSql = StrSql & " monto_entidad9=" & rs_Desmen!desmondec
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'FGZ - 22/12/2004
                    'Leo la tabla
                    StrSql = "SELECT * FROM sim_traza_gan WHERE "
                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    'StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                    OpenRecordset StrSql, rs_Traza_gan
                    
                    Tope = Tope + 1
                End If
            End If
            ' Se saca el 23/05/2006
            If (rs_Item!Itenro = 29) Or (rs_Item!Itenro = 55) Then 'Impuesto al debito bancario
                StrSql = "UPDATE sim_traza_gan SET "
                StrSql = StrSql & " promo =" & val_impdebitos
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                'StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            
                'FGZ - 22/12/2004
                'Leo la tabla
                StrSql = "SELECT * FROM sim_traza_gan WHERE "
                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                'StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                OpenRecordset StrSql, rs_Traza_gan
            End If
            ' FGZ - 19/04/2004
            
            rs_Desmen.MoveNext
         Loop
    
    
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!Ternro & _
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
                 " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                 " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
        'If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
        ' Maxi 13/12/2010 Cuando hay dif de plan item13 y devuelve tiene que restar el valor liquidado por eso saco el ABS de LIQ
        
        
        'FGZ - 05/06/2013 -----------------------------------------------------------------------------------------------------
        'Restauro otra vez lo de ABS porque estaba generando problemas cuando el monto no viene por ddjj sino que viene por liq
        'If Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
        If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            'Items_TOPE(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            Items_TOPE(rs_Item!Itenro) = Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
        Else
            'FGZ - 24/08/2005
            If Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) = 0 Then
                Items_TOPE(rs_Item!Itenro) = 0
            End If
            'FGZ - 24/08/2005
        End If
        'FGZ - 05/06/2013 -----------------------------------------------------------------------------------------------------
        
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
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-ProrrCuota" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_PRORR_CUOTA(rs_Item!Itenro)
        Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Tope" & " "
        Flog.writeline Espacios(Tabulador * 3) & Texto & Items_TOPE(rs_Item!Itenro)
    End If
    If HACE_TRAZA Then
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_OLD_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-ProrrCuota"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR_CUOTA(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(rs_Item!Itenro))
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Bruta ", Gan_Imponible - Descuentos + Items_TOPE(100))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Gan. Bruta - CMA y DONA.", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Neta ", (Gan_Imponible + Deducciones))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Para Machinea ", (Gan_Imponible + Deducciones - Items_TOPE(100)))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total art. 23", Ded_a23)
    End If
            
    
    ' Calculo el porcentaje de deduccion segun la ganancia neta
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia neta para el 648
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganneta =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "No hay esc. dedu para", Gan_Imponible)
            End If
            ' No se ha encontrado la escala de deduccion para el valor gan_imponible
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- % a tomar deduc." & Por_Deduccion
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- % a tomar deduc.", Por_Deduccion)
        End If
        
        'Aplico el porcentaje a las deducciones
        Ded_a23 = Ded_a23 * Por_Deduccion / 100
        
'        'Guardo el porcentaje de deduccion
        StrSql = "UPDATE sim_traza_gan SET "
        StrSql = StrSql & " porcdeduc =" & Por_Deduccion
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
            
    
    ' Calculo la Ganancia imponible
    Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia imponible para el 648
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganimpo =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    OpenRecordset StrSql, rs_Traza_gan
    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible" & Gan_Imponible
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Gan_Imponible)
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
    
    StrSql = "UPDATE sim_traza_gan SET "
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
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
'    StrSql = "SELECT * FROM sim_ficharet " & _
'             " WHERE empleado =" & buliq_empleado!ternro
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
    
    StrSql = "SELECT SUM(importe) monto FROM sim_ficharet " & _
             " WHERE empleado =" & buliq_empleado!Ternro & _
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc", val_impdebitos)
    End If
    'Si hay devolucion suma los impdebitos pedido por Ruben Vacarezza 22/02/2008
    If Retencion < 0 Then
            Retencion = Retencion - val_impdebitos
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc", val_impdebitos)
    End If
    
    
    
    ' Para el F649 va en el 9b
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " promo =" & val_impdebitos
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
            
    
    'FGZ - 30/12/2004
    'Determinar el saldo
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Not rs_Traza_gan.EOF Then
        StrSql = "UPDATE sim_traza_gan SET "
        StrSql = StrSql & "  saldo =" & Retencion
        StrSql = StrSql & "  ,retenciones =" & Ret_Ant
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 22/12/2004
        'Leo la tabla
        StrSql = "SELECT * FROM sim_traza_gan WHERE "
        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        'StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Retenciones anteriores", Ret_Ant)
        If Gan_Imponible > 0 Then
            If Not rs_escala.EOF Then
                rs_escala.MoveFirst
                If Not rs_escala.EOF Then
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", rs_escala!escporexe)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
                Else
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", 0)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver, x Tope General", Retencion)
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
        Call InsertarFichaRet(buliq_empleado!Ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope
            
            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET monto =" & Items_TOPE(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
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
        
        If Items_PRORR(I) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
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

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR_CUOTA(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR_CUOTA(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET prorr =" & Items_PRORR_CUOTA(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET old_liq =" & Items_OLD_LIQ(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
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


Public Function for_Ganancias_old(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
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
Dim p_sinprorrateo As Integer  'Indica que nunca prorratea

'Variables Locales
Dim Devuelve As Double
Dim Tope_Gral As Double
Dim Neto As Double
Dim prorratea As Double
Dim sinprorrateo As Double
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
Dim I As Long
Dim j As Integer
Dim Texto As String

'Vectores para manejar el proceso
Dim Items_DDJJ(100) As Double
Dim Items_LIQ(100) As Double
Dim Items_PRORR(100) As Double
Dim Items_PRORR_CUOTA(100) As Double
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
p_sinprorrateo = 1006


'FGZ - 19/04/2004
Dim Total_Empresa As Double
Dim Tope As Integer
'Dim rs_Rep19 As New ADODB.Recordset
Dim rs_Traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Double
Dim Cuota As Double

Total_Empresa = 0
Tope = 10

Descuentos = 0
' Primero limpio la traza
StrSql = "DELETE FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords

' Guardo la traza de Ganancia en traza_gan para utilizarla luego cuando se genere el reporte
StrSql = "INSERT INTO sim_traza_gan (pliqnro,pronro,concnro,fecha_pago,empresa,ternro,empleg) VALUES (" & _
         buliq_periodo!PliqNro & "," & _
         buliq_proceso!pronro & "," & _
         Buliq_Concepto(Concepto_Actual).ConcNro & "," & _
         ConvFecha(buliq_proceso!profecpago) & "," & _
         NroEmp & "," & _
         buliq_empleado!Ternro & "," & _
         buliq_empleado!Empleg & _
         ")"
objConn.Execute StrSql, , adExecuteNoRecords

'FGZ - 22/12/2004
'Leo la tabla
StrSql = "SELECT * FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
OpenRecordset StrSql, rs_Traza_gan
    

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).ConcNro)
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
ini_anyo_ret = C_Date("01/01/" & Ret_Ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro
sinprorrateo = 0

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
    Case p_sinprorrateo:
        sinprorrateo = rs_wf_tpa!Valor
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
      Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_Neto, "El Neto del Mes es negativo, se seteara en cero.", Neto)
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
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Neto del Mes", Neto)
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
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
                                StrSql = "UPDATE sim_traza_gan SET "
                                StrSql = StrSql & " cuit_entidad" & I & "='" & rs_Desmen!descuit & "',"
                                StrSql = StrSql & " entidad" & I & "='" & rs_Desmen!DesRazsoc & "',"
                                StrSql = StrSql & " monto_entidad" & I & "=" & rs_Desmen!desmondec
                                StrSql = StrSql & " WHERE "
                                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'FGZ - 22/12/2004
                                'Leo la tabla
                                StrSql = "SELECT * FROM sim_traza_gan WHERE "
                                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                'StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                OpenRecordset StrSql, rs_Traza_gan
                                
                                
                                Tope = Tope + 1
                            Else
                                If I = 15 Then
                                    Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!Empleg ' empleado.empleg
                                Else
                                    StrSql = "UPDATE sim_traza_gan SET "
                                    StrSql = StrSql & " monto_entidad" & I & "= monto_entidad" & I & " + " & rs_Desmen!desmondec
                                    StrSql = StrSql & " WHERE "
                                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    'FGZ - 22/12/2004
                                    'Leo la tabla
                                    StrSql = "SELECT * FROM sim_traza_gan WHERE "
                                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    'StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
        StrSql = "SELECT * FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!Ternro & _
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
        
             If CBool(rs_itemacum!itaprorratea) And (sinprorrateo = 0) Then
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
                 " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                 " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND itemconc.itenro =" & rs_Item!Itenro & _
                 " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
        OpenRecordset StrSql, rs_itemconc
        
        Do While Not rs_itemconc.EOF
            If CBool(rs_itemconc!itcprorratea) And (sinprorrateo = 0) Then
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
        StrSql = "SELECT * FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!Ternro & _
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
                 " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                 " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                    StrSql = "UPDATE sim_traza_gan SET "
                    StrSql = StrSql & " cuit_entidad9 ='" & rs_Desmen!descuit & "',"
                    StrSql = StrSql & " entidad9='" & rs_Desmen!DesRazsoc & "',"
                    StrSql = StrSql & " monto_entidad9=" & rs_Desmen!desmondec
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'FGZ - 22/12/2004
                    'Leo la tabla
                    StrSql = "SELECT * FROM sim_traza_gan WHERE "
                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    'StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                    OpenRecordset StrSql, rs_Traza_gan
                    
                    Tope = Tope + 1
                End If
            End If
            ' Se saca el 23/05/2006
            If (rs_Item!Itenro = 29) Or (rs_Item!Itenro = 55) Then 'Impuesto al debito bancario
                StrSql = "UPDATE sim_traza_gan SET "
                StrSql = StrSql & " promo =" & val_impdebitos
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                'StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            
                'FGZ - 22/12/2004
                'Leo la tabla
                StrSql = "SELECT * FROM sim_traza_gan WHERE "
                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                'StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                OpenRecordset StrSql, rs_Traza_gan
            End If
            ' FGZ - 19/04/2004
            
            rs_Desmen.MoveNext
         Loop
    
    
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!Ternro & _
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
                 " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                 " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_OLD_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(rs_Item!Itenro))
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Bruta ", Gan_Imponible - Descuentos + Items_TOPE(100))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Gan. Bruta - CMA y DONA.", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Neta ", (Gan_Imponible + Deducciones))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Para Machinea ", (Gan_Imponible + Deducciones - Items_TOPE(100)))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total art. 23", Ded_a23)
    End If
            
    
    ' Calculo el porcentaje de deduccion segun la ganancia neta
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia neta para el 648
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganneta =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "No hay esc. dedu para", Gan_Imponible)
            End If
            ' No se ha encontrado la escala de deduccion para el valor gan_imponible
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- % a tomar deduc." & Por_Deduccion
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- % a tomar deduc.", Por_Deduccion)
        End If
        
        'Aplico el porcentaje a las deducciones
        Ded_a23 = Ded_a23 * Por_Deduccion / 100
        
'        'Guardo el porcentaje de deduccion
        StrSql = "UPDATE sim_traza_gan SET "
        StrSql = StrSql & " porcdeduc =" & Por_Deduccion
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
    End If
            
    
    ' Calculo la Ganancia imponible
    Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia imponible para el 648
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganimpo =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    OpenRecordset StrSql, rs_Traza_gan
    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible" & Gan_Imponible
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Gan_Imponible)
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
    
    StrSql = "UPDATE sim_traza_gan SET "
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
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
'    StrSql = "SELECT * FROM sim_ficharet " & _
'             " WHERE empleado =" & buliq_empleado!ternro
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
    
    StrSql = "SELECT SUM(importe) monto FROM sim_ficharet " & _
             " WHERE empleado =" & buliq_empleado!Ternro & _
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc", val_impdebitos)
    End If
    'Si hay devolucion suma los impdebitos pedido por Ruben Vacarezza 22/02/2008
    If Retencion < 0 Then
            Retencion = Retencion - val_impdebitos
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc", val_impdebitos)
    End If
    
    
    
    ' Para el F649 va en el 9b
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " promo =" & val_impdebitos
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
            
    
    'FGZ - 30/12/2004
    'Determinar el saldo
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Not rs_Traza_gan.EOF Then
        StrSql = "UPDATE sim_traza_gan SET "
        StrSql = StrSql & "  saldo =" & Retencion
        StrSql = StrSql & "  ,retenciones =" & Ret_Ant
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 22/12/2004
        'Leo la tabla
        StrSql = "SELECT * FROM sim_traza_gan WHERE "
        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        'StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Retenciones anteriores", Ret_Ant)
        If Gan_Imponible > 0 Then
            If Not rs_escala.EOF Then
                rs_escala.MoveFirst
                If Not rs_escala.EOF Then
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", rs_escala!escporexe)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
                Else
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", 0)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver, x Tope General", Retencion)
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
        Call InsertarFichaRet(buliq_empleado!Ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope
            
            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET monto =" & Items_TOPE(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
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
        
        If Items_PRORR(I) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
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

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET prorr =" & Items_PRORR(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET old_liq =" & Items_OLD_LIQ(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
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
        'FGZ 31/03/2005
        
        I = I + 1
    Loop

    exito = Bien
    for_Ganancias_old = Monto
    
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
StrSql = "DELETE FROM sim_traza_gan WHERE "
StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords

' Guardo la traza de Ganancia en traza_gan para utilizarla luego cuando se genere el reporte
StrSql = "INSERT INTO sim_traza_gan (pliqnro,pronro,concnro,fecha_pago,empresa,ternro,empleg) VALUES (" & _
         buliq_periodo!PliqNro & "," & _
         buliq_proceso!pronro & "," & _
         Buliq_Concepto(Concepto_Actual).ConcNro & "," & _
         ConvFecha(buliq_proceso!profecpago) & "," & _
         NroEmp & "," & _
         buliq_empleado!Ternro & "," & _
         buliq_empleado!Empleg & _
         ")"
objConn.Execute StrSql, , adExecuteNoRecords

'FGZ - 22/12/2004
'Leo la tabla
StrSql = "SELECT * FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
OpenRecordset StrSql, rs_Traza_gan
    

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).ConcNro)
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
ini_anyo_ret = C_Date("01/01/" & Ret_Ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro

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
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_Tope_Gral, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_Neto, "Neto del Mes", Neto)
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
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
                                StrSql = "UPDATE sim_traza_gan SET "
                                StrSql = StrSql & " cuit_entidad" & I & "='" & rs_Desmen!descuit & "',"
                                StrSql = StrSql & " entidad" & I & "='" & rs_Desmen!DesRazsoc & "',"
                                StrSql = StrSql & " monto_entidad" & I & "=" & rs_Desmen!desmondec
                                StrSql = StrSql & " WHERE "
                                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                objConn.Execute StrSql, , adExecuteNoRecords
                                'FGZ - 22/12/2004
                                'Leo la tabla
                                StrSql = "SELECT * FROM sim_traza_gan WHERE "
                                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                'StrSql = StrSql & " AND empresa =" & NroEmp
                                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                OpenRecordset StrSql, rs_Traza_gan
                                
                                
                                Tope = Tope + 1
                            Else
                                If I = 15 Then
                                    Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!Empleg ' empleado.empleg
                                Else
                                    StrSql = "UPDATE sim_traza_gan SET "
                                    StrSql = StrSql & " monto_entidad" & I & "= monto_entidad" & I & " + " & rs_Desmen!desmondec
                                    StrSql = StrSql & " WHERE "
                                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    
                                    'FGZ - 22/12/2004
                                    'Leo la tabla
                                    StrSql = "SELECT * FROM sim_traza_gan WHERE "
                                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    'StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
                 " AND empleado = " & buliq_empleado!Ternro & _
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
                 " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                 " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                 " AND empleado = " & buliq_empleado!Ternro & _
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
                 " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                 " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
        StrSql = "SELECT * FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                    StrSql = "UPDATE sim_traza_gan SET "
                    StrSql = StrSql & " cuit_entidad9 ='" & rs_Desmen!descuit & "',"
                    StrSql = StrSql & " entidad9='" & rs_Desmen!DesRazsoc & "',"
                    StrSql = StrSql & " monto_entidad9=" & rs_Desmen!desmondec
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'FGZ - 22/12/2004
                    'Leo la tabla
                    StrSql = "SELECT * FROM sim_traza_gan WHERE "
                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    'StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                    OpenRecordset StrSql, rs_Traza_gan
                    
                    Tope = Tope + 1
                End If
            End If
            If rs_Item!Itenro = 22 Then 'Impuesto al debito bancario
                StrSql = "UPDATE sim_traza_gan SET "
                StrSql = StrSql & " promo =" & rs_Desmen!desmondec
                StrSql = StrSql & " WHERE "
                StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                objConn.Execute StrSql, , adExecuteNoRecords
            
                'FGZ - 22/12/2004
                'Leo la tabla
                StrSql = "SELECT * FROM sim_traza_gan WHERE "
                StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                'StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                OpenRecordset StrSql, rs_Traza_gan
            End If
            ' FGZ - 19/04/2004
            
            rs_Desmen.MoveNext
         Loop
    
    
        'Busco las liquidaciones anteriores
        StrSql = "SELECT * FROM desliq WHERE itenro =" & rs_Item!Itenro & _
                 " AND empleado = " & buliq_empleado!Ternro & _
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
                 " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                 " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_OLD_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(rs_Item!Itenro))
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Neta ", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total art. 23", Ded_a23)
    End If
            
    
    ' Calculo el porcentaje de deduccion segun la ganancia neta
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia neta para el 648
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganneta =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "No hay esc. dedu para", Gan_Imponible)
            End If
            ' No se ha encontrado la escala de deduccion para el valor gan_imponible
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- % a tomar deduc." & Por_Deduccion
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- % a tomar deduc.", Por_Deduccion)
        End If
        
        'Aplico el porcentaje a las deducciones
        Ded_a23 = Ded_a23 * Por_Deduccion / 100
    End If
            
    
    ' Calculo la Ganancia imponible
    Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23
    
    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia imponible para el 648
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganimpo =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    OpenRecordset StrSql, rs_Traza_gan
    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible" & Gan_Imponible
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Gan_Imponible)
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
    
    StrSql = "UPDATE sim_traza_gan SET "
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
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
             " WHERE empleado =" & buliq_empleado!Ternro
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
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Not rs_Traza_gan.EOF Then
        StrSql = "UPDATE sim_traza_gan SET "
        StrSql = StrSql & "  saldo =" & Retencion
        StrSql = StrSql & "  ,retenciones =" & Ret_Ant
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 22/12/2004
        'Leo la tabla
        StrSql = "SELECT * FROM sim_traza_gan WHERE "
        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        'StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Retenciones anteriores", Ret_Ant)
        If Gan_Imponible > 0 Then
            If Not rs_escala.EOF Then
                rs_escala.MoveFirst
                If Not rs_escala.EOF Then
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", rs_escala!escporexe)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
                Else
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", 0)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver, x Tope General", Retencion)
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
        Call InsertarFichaRet(buliq_empleado!Ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope
            
            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET monto =" & Items_TOPE(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
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
        
        If Items_PRORR(I) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
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

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET prorr =" & Items_PRORR(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET old_liq =" & Items_OLD_LIQ(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
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
            StrSql = "SELECT * FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
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
        'FGZ 31/03/2005
        
        I = I + 1
    Loop

    exito = Bien
    for_Ganancias_Schering = Monto
End Function


Public Sub InsertarFichaRet(ByVal Ternro As Long, ByVal Fecha As Date, ByVal importe As Double, ByVal pronro As Long)
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
    
    StrSql = "INSERT INTO sim_ficharet (empleado,fecha,pronro,importe,liqsistema) VALUES (" & _
             Ternro & "," & _
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
Dim Mes As Integer
Dim Dia As Integer
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_acumulador, "1- Acumulador", n_acumulador)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_Monto_Manual, "1- Monto manual", monto_manual)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_divisor, "1- Divisor", divisor)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_Multiplicador, "1- Multiplicador", multiplicador)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_can_meses, "1- Cant meses", cant_meses)
    End If


If monto_manual = 0 Then
    'Busco la licencia del empleado para ver el calculo
    StrSql = "SELECT * FROM sim_emp_lic WHERE (empleado = " & buliq_empleado!Ternro & " )" & _
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
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 9, "2- Nro accidente ", rs_Accidente!accnro)
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
                            
                        Call bus_Antfases(rs_Accidente!accfecha, fectope, Dia, Mes, Anio)
                        'RUN antfases.p (recid(buliq-empleado), accidente.accfecha, fectope, output dia, output mes, output anio)
                        divisor = (Anio * 360) + (Mes * 30) + Dia
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
                            StrSql = "SELECT * FROM sim_proceso "
                            StrSql = StrSql & " INNER JOIN sim_cabliq ON sim_proceso.pronro = sim_cabliq.pronro "
                            StrSql = StrSql & " INNER JOIN sim_acu_liq ON sim_cabliq.cliqnro = sim_acu_liq.cliqnro "
                            StrSql = StrSql & " WHERE sim_proceso.pliqnro = " & rs_Periodo!PliqNro
                            StrSql = StrSql & " AND sim_cabliq.empleado = " & buliq_empleado!Ternro
                            StrSql = StrSql & " AND sim_acu_liq.acunro = " & n_acumulador
                            OpenRecordset StrSql, rs_Acu_Liq
                            Do While Not rs_Acu_Liq.EOF
                                  Monto_Mes = Monto_Mes + rs_Acu_Liq!almonto
                                rs_Acu_Liq.MoveNext
                            Loop
                            If HACE_TRAZA Then
                                Texto = Meses & "- Monto mes " & rs_Periodo!pliqmes & " del Anio " & rs_Periodo!pliqanio
                                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 9, Texto, Monto_Mes)
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
                        
                        Call bus_Antfases(rs_Accidente!accfecha, fectope, Dia, Mes, Anio)
                        'RUN antfases.p (recid(buliq-empleado), accidente.accfecha, fectope, output dia, output mes, output anio).
                        divisor = (Anio * 360) + (Mes * 30) + Dia
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
    StrSql = "SELECT * FROM sim_traza_gan WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    'StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).Concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_Traza_gan
    I = 0
    Do While Not rs_Traza_gan.EOF
        I = I + 1
        Traza_Gan(I).PliqNro = IIf(Not EsNulo(rs_Traza_gan!PliqNro), rs_Traza_gan!PliqNro, 0)
        Traza_Gan(I).ConcNro = IIf(Not EsNulo(rs_Traza_gan!ConcNro), rs_Traza_gan!ConcNro, 0)
        Traza_Gan(I).Empresa = IIf(Not EsNulo(rs_Traza_gan!Empresa), rs_Traza_gan!Empresa, 0)
        Traza_Gan(I).Fecha_pago = IIf(Not EsNulo(rs_Traza_gan!Fecha_pago), rs_Traza_gan!Fecha_pago, 0)
        Traza_Gan(I).Ternro = IIf(Not EsNulo(rs_Traza_gan!Ternro), rs_Traza_gan!Ternro, 0)
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
    StrSql = "DELETE FROM sim_traza_gan WHERE"
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    'StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).Concnro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    'TRAZA_GAN --->
            
            
    'TRAZA_GAN_ITEM_TOP --->
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "--- Depuracion de traza_gan_item_top ---"
    End If
    StrSql = "SELECT * FROM sim_traza_gan_item_top "
    StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    OpenRecordset StrSql, rs_Traza_gan_items_tope
    I = 0
    Do While Not rs_Traza_gan_items_tope.EOF
        I = I + 1
        Traza_Gan_Item_Top(I).Itenro = IIf(Not EsNulo(rs_Traza_gan_items_tope!Itenro), rs_Traza_gan_items_tope!Itenro, 0)
        Traza_Gan_Item_Top(I).Ternro = IIf(Not EsNulo(rs_Traza_gan_items_tope!Ternro), rs_Traza_gan_items_tope!Ternro, 0)
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
    StrSql = "DELETE FROM sim_traza_gan_item_top "
    StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    objConn.Execute StrSql, , adExecuteNoRecords
    'TRAZA_GAN_ITEM_TOP --->
        
    
    'DESLIQ --->
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "--- Depuracion de desliq ---"
    End If
    StrSql = "SELECT * FROM desliq WHERE empleado = " & buliq_empleado!Ternro
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
    StrSql = "DELETE FROM sim_desliq WHERE empleado = " & buliq_empleado!Ternro
    StrSql = StrSql & " AND dlfecha = " & ConvFecha(buliq_proceso!profecpago)
    objConn.Execute StrSql, , adExecuteNoRecords
    'DESLIQ --->
        
        
        
    'FICHARET --->
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "--- Depuracion de ficharet ---"
    End If
    StrSql = "SELECT * FROM ficharet WHERE empleado = " & buliq_empleado!Ternro
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
    
    StrSql = "DELETE FROM sim_ficharet "
    StrSql = StrSql & " WHERE pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND empleado =" & buliq_empleado!Ternro
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
    StrSql = "DELETE FROM sim_traza_gan WHERE"
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Traza_gan - inserto los guardados anteriormente
    'FGZ - 27/09/2007 -
    'For I = 1 To 1  'UBound(Traza_Gan)
    For I = 1 To UBound(Traza_Gan)
        If Traza_Gan(I).Ternro <> 0 Then
            StrSql = "INSERT INTO sim_traza_gan (pliqnro,concnro,pronro,empresa,fecha_pago,ternro "
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
            StrSql = StrSql & "," & Traza_Gan(I).ConcNro
            StrSql = StrSql & "," & Traza_Gan(I).pronro
            StrSql = StrSql & "," & Traza_Gan(I).Empresa
            StrSql = StrSql & "," & ConvFecha(Traza_Gan(I).Fecha_pago)
            StrSql = StrSql & "," & Traza_Gan(I).Ternro
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
    StrSql = "DELETE FROM sim_traza_gan_item_top "
    StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    objConn.Execute StrSql, , adExecuteNoRecords
    
    'Traza_gan - inserto los guardados anteriormente
    For I = 1 To 1  'UBound(Traza_Gan_item_top)
        If Traza_Gan_Item_Top(I).Itenro <> 0 Then
            StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,monto,empresa,itenro"
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
            StrSql = StrSql & Traza_Gan_Item_Top(I).Ternro
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
    StrSql = "DELETE FROM sim_desliq WHERE empleado = " & buliq_empleado!Ternro
    StrSql = StrSql & " AND dlfecha = " & ConvFecha(buliq_proceso!profecpago)
    objConn.Execute StrSql, , adExecuteNoRecords
    
    For I = 1 To UBound(Desliq)
        If Desliq(I).Itenro <> 0 Then
        
            StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro"
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
    StrSql = "DELETE FROM sim_ficharet "
    StrSql = StrSql & " WHERE pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND empleado =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    Ret_Actual = 0
    For I = 1 To 1 'UBound(ficharet)
        If Ficharet(I).Empleado <> 0 Then
            StrSql = "INSERT INTO sim_ficharet (empleado,fecha,pronro,importe,liqsistema"
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
                 " WHERE empleado =" & buliq_empleado!Ternro
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
        StrSql = StrSql & " INNER JOIN sim_proceso ON periodo.pliqnro = sim_proceso.pliqnro "
        StrSql = StrSql & " INNER JOIN sim_cabliq ON sim_proceso.pronro = sim_cabliq.pronro "
        StrSql = StrSql & " INNER JOIN sim_detliq ON sim_cabliq.cliqnro = sim_detliq.cliqnro "
        StrSql = StrSql & " WHERE sim_cabliq.empleado = " & buliq_empleado!Ternro
        StrSql = StrSql & " AND sim_detliq.concnro = " & Buliq_Concepto(Concepto_Actual).ConcNro
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
    StrSql = "SELECT * FROM sim_fases WHERE real = -1 AND empleado = " & buliq_empleado!Ternro
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
    StrSql = "SELECT ammonto FROM sim_acu_mes " & _
             " INNER JOIN acumulador ON sim_acu_mes.acunro = acumulador.acunro " & _
             " WHERE ternro = " & buliq_empleado!Ternro & _
             " AND sim_acu_mes.acunro =" & v_AcuNoRemu & _
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
Public Function For_Premio_Semestre() As Double
    
Dim Porcent As Double
Dim ConcAjusteSacNro As Long
Dim Monto As Double

Dim CantPremio As Long
Dim AcumPremioMensual As Long
Dim ConcAjusteSac As String
Dim p_CantPremio As Integer
Dim p_AcumPremioMensual As Integer
Dim p_ConcAjusteSac As Integer


'Inicializacion de parametros
p_CantPremio = 77
p_AcumPremioMensual = 68
p_ConcAjusteSac = 7

Dim rs_Datos As New ADODB.Recordset
Dim rs_Per As New ADODB.Recordset

    exito = False
    
    'Obtencion de los parametros de WorkFile
    StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha = " & ConvFecha(AFecha)
    OpenRecordset StrSql, rs_Datos
    
    Do While Not rs_Datos.EOF
        Select Case rs_Datos!tipoparam
        Case p_CantPremio:
            CantPremio = rs_Datos!Valor
        Case p_AcumPremioMensual:
            AcumPremioMensual = rs_Datos!Valor
        Case p_ConcAjusteSac:
            ConcAjusteSac = rs_Datos!Valor
        End Select
        
        rs_Datos.MoveNext
    Loop
    rs_Datos.Close
    
    
    
    If ((CantPremio < 4) And ((buliq_periodo!pliqmes = 4) Or (buliq_periodo!pliqmes = 10))) Then
        
        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 3) & "Formula de premio mensual retorna 0. Cant Premios < 4 o el mes no es ni 4 ni 10"
        
        exito = True
        For_Premio_Semestre = 0
        
    Else
        
        Select Case CantPremio
            Case 4:
                Porcent = 35 / 100
            Case 5:
                Porcent = 20 / 100
            Case 6:
                Porcent = 10 / 100
            Case Else:
                exito = False
                If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 3) & "La cantidad de premios debe ser entre 4 y 6"
                For_Premio_Semestre = 0
                Exit Function
        End Select
        
        StrSql = "SELECT concnro FROM concepto WHERE conccod = '" & ConcAjusteSac & "'"
        OpenRecordset StrSql, rs_Datos
        If rs_Datos.EOF Then
            
            exito = False
            If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 3) & "Concepto de ajuste " & ConcAjusteSac & " No encontrado en base"
            For_Premio_Semestre = 0
            Exit Function
        
        Else
        
            ConcAjusteSacNro = rs_Datos!ConcNro
            
        End If
        
        
        'Busco los acumuladores de los premios
        StrSql = "SELECT ammes, amanio, ammonto"
        StrSql = StrSql & " FROM sim_acu_mes"
        StrSql = StrSql & " WHERE ternro = " & buliq_empleado!Ternro
        StrSql = StrSql & " AND acunro = " & AcumPremioMensual
        If buliq_periodo!pliqmes = 10 Then
            StrSql = StrSql & " (amanio = " & buliq_periodo!pliqanio & " AND ammes <= 9)"
        Else
            StrSql = StrSql & " (amanio = " & buliq_periodo!pliqanio & " AND ammes <= 3) OR (amanio = " & CInt(buliq_periodo!pliqanio) - 1 & " AND ammes >= 10)"
        End If
        StrSql = StrSql & " ORDER BY amanio, ammes"
        OpenRecordset StrSql, rs_Datos
        
        'Por cada premio creo una novedad retroactiva
        Do While Not rs_Datos.EOF
        
            If Not EsNulo(rs_Datos!ammonto) Then
                If CDbl(rs_Datos!ammonto) <> 0 Then
                
                    StrSql = "SELECT pliqdesde, pliqhasta FROM periodo"
                    StrSql = StrSql & " WHERE pliqanio = " & rs_Datos!amanio
                    StrSql = StrSql & " AND pliqmes = " & rs_Datos!ammes
                    OpenRecordset StrSql, rs_Per
                    If rs_Per.EOF Then
                        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 3) & "No se encontro el periodo para el mes " & rs_Datos!ammes & " y año " & rs_Datos!amanio & " No se crea novedad"
                        exito = False
                        For_Premio_Semestre = 0
                        Exit Function
                    Else
                    
                        StrSql = "INSERT INTO sim_novemp"
                        StrSql = StrSql & " (concnro,tpanro,empleado,nevalor"
                        StrSql = StrSql & " ,nepliqdesde,nepliqhasta)"
                        StrSql = StrSql & " VALUES"
                        StrSql = StrSql & " (" & ConcAjusteSacNro
                        StrSql = StrSql & " ,35"
                        StrSql = StrSql & " ," & buliq_empleado!Ternro
                        StrSql = StrSql & " ," & Porcent * CDbl(rs_Datos!ammonto)
                        StrSql = StrSql & " ," & ConvFecha(rs_Per!pliqdesde)
                        StrSql = StrSql & " ," & ConvFecha(rs_Per!pliqhasta)
                        StrSql = StrSql & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        Monto = Monto + Porcent * CDbl(rs_Datos!ammonto)
                        
                        If CBool(USA_DEBUG) Then Flog.writeline Espacios(Tabulador * 3) & "Se creo novedad"
                    
                    End If
                    
                End If
            End If
            
            rs_Datos.MoveNext
        Loop
        
        exito = True
        For_Premio_Semestre = Monto
    
    End If
    
    
If rs_Datos.State = adStateOpen Then rs_Datos.Close
Set rs_Datos = Nothing

End Function

Public Function for_Ganancias2017(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias 2013.
' Autor      :
' Fecha      :
' Ultima Mod.: 26/01/2017
' Descripcion: nueva formula 2017.
' Ultima Mod.:
' ---------------------------------------------------------------------------------------------
Dim p_Devuelve As Integer           'si devuelve ganancia o no
Dim p_Tope_Gral As Integer          'Tope Gral de retencion
Dim p_Neto As Integer               'Base para el tope
Dim p_prorratea As Integer          'Si prorratea o no para liq. finales
Dim p_sinprorrateo As Integer       'Indica que nunca prorratea
Dim p_Deduccion_Zona                'EAM(6.73)- Parámetro utilizado para Deduccion de zona desfavorable
Dim p_AC_HorasExtras As Long        'EAM(6.73) - Busco en el acumulador las horas extras que tributan a Ganancia Imponible

'Variables Locales
Dim Devuelve As Double
Dim Tope_Gral As Double
Dim Neto As Double
Dim prorratea As Double
Dim sinprorrateo As Double
Dim Retencion As Double
Dim Gan_Imponible As Double
Dim Deducciones As Double
Dim Descuentos As Double
Dim Ded_a23 As Double
Dim Impuesto_Escala As Double
Dim Ret_Ant As Double

Dim Por_Deduccion_zona As Double
Dim valor_ant As Double
Dim valor_act As Double 'sebastian stremel - 05/09/2013
Dim AC_HorasExtras As Integer

Dim Ret_Mes As Integer
Dim Ret_Ano As Integer
Dim fin_mes_ret As Date
Dim ini_anyo_ret As Date
Dim Con_liquid As Integer
Dim I As Long
Dim j As Integer
Dim Texto As String

'Vectores para manejar el proceso
Dim Items_DDJJ(100) As Double
Dim Items_LIQ(100) As Double
Dim Items_PRORR(100) As Double
Dim Items_PRORR_CUOTA(100) As Double
Dim Items_OLD_LIQ(100) As Double
Dim Items_TOPE(100) As Double
Dim Items_ART_23(100) As Boolean

'Recorsets Auxiliares
'Dim rs_wf_tpa As New ADODB.Recordset
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
Dim rs_Aux As New ADODB.Recordset
Dim Hasta As Integer
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Double
Dim val_impdebitos As Double
Dim fechaFichaH As Date
Dim fechaFichaD As Date

Dim Terminar As Boolean

Dim Total_Empresa As Double
Dim Tope As Integer
Dim rs_Traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Double
Dim Cuota As Double
Dim Beneficio As Boolean
Dim p_Beneficio As Long
Dim ctrlItem20y56 As Boolean        'EAM (6.43) Controla los item 20 y 56 cuando es una liquidacion final o la fecha de pago es 31/12
Dim Extranjero As Boolean           'EAM (6.44) Controla si es expatriado
Dim p_Extranjero As Long            'EAM (6.44) Controla si es expatriado

Bien = False
Por_Deduccion_zona = 0
ctrlItem20y56 = False


'Comienzo
p_Devuelve = 1001
p_Tope_Gral = 1002
p_Neto = 1003
p_prorratea = 1005
p_sinprorrateo = 1006
p_Deduccion_Zona = 1008 'EAM(6.73)- Parámetro utilizado para Deduccion de zona desfavorable
p_AC_HorasExtras = 143 'Concepto de Neto de SAC 2013
p_Beneficio = 1140      'Beneficio Item56
p_Extranjero = 1141     'Controla si es expatriado

Total_Empresa = 0
Tope = 10
Descuentos = 0
Beneficio = False


' Primero limpio la traza
Call LimpiaTraza_Gan

StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
OpenRecordset StrSql, rs_Traza_gan
    
If HACE_TRAZA Then
    Call LimpiarTrazaConcepto(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro)
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
ini_anyo_ret = C_Date("01/01/" & Ret_Ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro
sinprorrateo = 0

'Obtencion de los parametros de WorkFile
For I = LI_WF_Tpa To LS_WF_Tpa
    Select Case Arr_WF_TPA(I).tipoparam
    Case p_Devuelve:
        Devuelve = Arr_WF_TPA(I).Valor
    Case p_Tope_Gral:
        Tope_Gral = Arr_WF_TPA(I).Valor
    Case p_Neto:
        Neto = Arr_WF_TPA(I).Valor
    Case p_prorratea:
        prorratea = Arr_WF_TPA(I).Valor
    Case p_sinprorrateo:
        sinprorrateo = Arr_WF_TPA(I).Valor
    Case p_Deduccion_Zona:
        Por_Deduccion_zona = Arr_WF_TPA(I).Valor
    Case p_AC_HorasExtras:
        AC_HorasExtras = Arr_WF_TPA(I).Valor
    Case p_Beneficio:
        Beneficio = CBool(Arr_WF_TPA(I).Valor)
    Case p_Extranjero:
        Extranjero = CBool(Arr_WF_TPA(I).Valor)
    End Select
Next I

'Si es una liq. final no prorratea y tomo la escala de diciembre
If prorratea = 0 Then
    Ret_Mes = 12
    fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
End If

'EAM (v6.43) - Si la fecha de pago es 31/12 tiene que tener en cuenta el item 20 y 56, sino NO.
'EAM (6.57) - se agrego la condición (prorratea = 0) para que se tenga en cuenta tambien los item cuando es una liq. final
If CDate(buliq_proceso!profecpago) = CDate("31/12/" & Ret_Ano) Or (prorratea = 0) Then
    ctrlItem20y56 = True
End If


If Neto < 0 Then
   If CBool(USA_DEBUG) Then
      Flog.writeline Espacios(Tabulador * 3) & "El Neto del mes es negativo, se setea en cero."
   End If
   If HACE_TRAZA Then
      Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_Neto, "El Neto del Mes es negativo, se seteara en cero.", Neto)
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
    
    Flog.writeline Espacios(Tabulador * 3) & "Beneficio devolucion Anticipada " & Beneficio
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Neto del Mes", Neto)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Beneficio devolucion Anticipada", Beneficio)
End If

    'Limpiar items que suman al articulo 23
    For I = 1 To 100
        Items_ART_23(I) = False
    Next I
    val_impdebitos = 0
    'val_ComprasExt = 0
    

' Recorro todos los items de Ganancias
'FGZ - 08/06/2012 ----------------
StrSql = "SELECT itenro,itetipotope,itesigno,iteitemstope,iteporctope,itenom FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_Item
Do While Not rs_Item.EOF
  'FGZ - 08/10/2013 -----------------------------------------------------------------------------------
  ' Impuestos y debitos Bancarios va como Promocion
  ' Ahora Compras en exterior tb
  If (rs_Item!Itenro = 29 Or rs_Item!Itenro = 55 Or rs_Item!Itenro = 56) Then
    If (Ret_Mes = 12) And (ctrlItem20y56 = True) Then
        StrSql = "SELECT desmondec FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                 " AND desano=" & Ret_Ano & _
                 " AND itenro = " & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
        'If Not rs_Desmen.EOF Then
        Do While Not rs_Desmen.EOF
           If rs_Item!Itenro = 29 Then
             'val_impdebitos = rs_Desmen!desmondec * 0.34
             val_impdebitos = val_impdebitos + (rs_Desmen!desmondec * 0.34)
            
            'FGZ - 16/12/2015 ------------------------
            Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.34)
            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.34)
             'FGZ - 16/12/2015 ------------------------
           Else
                'If rs_Item!Itenro = 23 Then
                If rs_Item!Itenro = 56 Then
                    'val_impdebitos = rs_Desmen!desmondec
                    val_impdebitos = val_impdebitos + rs_Desmen!desmondec
                    
                    'EAM (6.56) - Se comenta la linea para que tenga en cuenta el valor mas de una vez
                    'FGZ - 16/12/2015 ------------------------
                    'Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (rs_Desmen!desmondec)
                    'Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (rs_Desmen!desmondec)
                    'FGZ - 16/12/2015 ------------------------
                Else
                    'val_impdebitos = rs_Desmen!desmondec * 0.17
                    val_impdebitos = val_impdebitos + (rs_Desmen!desmondec * 0.17)

                    'FGZ - 16/12/2015 ------------------------
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.17)
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.17)
                    'FGZ - 16/12/2015 ------------------------

                End If
           End If
        'End If
            
            rs_Desmen.MoveNext
        Loop
        
        rs_Desmen.Close
    Else
        If rs_Item!Itenro = 56 Then
            If Beneficio Then
                StrSql = "SELECT sum(desmondec) total FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                         " AND desano=" & Ret_Ano & _
                         " AND Month(desfecdes) <= " & Ret_Mes & _
                         " AND itenro = " & rs_Item!Itenro
                OpenRecordset StrSql, rs_Desmen
                If Not rs_Desmen.EOF Then
                    'FGZ - 06/04/2015 --------------------------------------
                    'val_impdebitos = val_impdebitos + rs_Desmen!Total
                    val_impdebitos = val_impdebitos + IIf(IsNull(rs_Desmen!total), 0, rs_Desmen!total)
                    'FGZ - 06/04/2015 --------------------------------------
                End If
            End If
        End If
    End If
    'FGZ - 19/01/2015 -----------------------------------------------
  Else
    
    'EAM (v6.40) - Solo se considera item 20 si es fin de año o final
    'EAM (v6.57) - Se agrego condicion para que se tenga en cuenta si es fin de año o es liquidacion final
                '(((Ret_Mes <> 12) And (ctrlItem20y56 = False)) Or ((Ret_Mes = 12) And (ctrlItem20y56 = False)))
    If (rs_Item!Itenro = 20) And (((Ret_Mes <> 12) And (ctrlItem20y56 = False)) Or ((Ret_Mes = 12) And (ctrlItem20y56 = False))) Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Item 20 No considerado"
        End If
    Else
        'EAM (6.44) - Si el Item es 10,11,12 o17 y Extranjero percibe su sueldo en Argentina. No se tiene en cuenta el item
        If ((rs_Item!Itenro = 17) Or (rs_Item!Itenro = 10) Or (rs_Item!Itenro = 11) Or (rs_Item!Itenro = 12)) And (Extranjero) Then
            Flog.writeline Espacios(Tabulador * 3) & "No se tiene en cuenta el Item " & rs_Item!Itenro & ". Extranjero que percibe su sueldo en Argentina"
            GoTo SiguienteItem
        End If
        
                
        Select Case rs_Item!itetipotope
        Case 1: ' el valor a tomar es lo que dice la escala
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT itenro,vimonto FROM valitem WHERE viano =" & Ret_Ano & _
                     " AND itenro=" & rs_Item!Itenro & _
                     " AND vimes =" & Ret_Mes
            OpenRecordset StrSql, rs_valitem
            
            Do While Not rs_valitem.EOF
                Items_DDJJ(rs_valitem!Itenro) = rs_valitem!vimonto
                Items_TOPE(rs_valitem!Itenro) = rs_valitem!vimonto
                
                rs_valitem.MoveNext
            Loop
    
'            'Agregado Maxi 29/08/2013 -------------------------------------------------------------------------------------
'            If rs_Item!Itenro = 16 Then
'
            'EAM(6.73) - Busco los acumuladores de la liquidacion
            StrSql = "SELECT acunro,itasigno FROM itemacum " & _
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

            'EAM(6.73) - Busco los conceptos de la liquidacion
            StrSql = "SELECT itcsigno,dlimonto FROM itemconc " & _
                    " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                    " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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


            'EAM(6.73) - Busco las liquidaciones anteriores
            StrSql = "SELECT dlmonto,dlprorratea,dlfecha FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                    " AND empleado = " & buliq_empleado!Ternro & _
                    " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                    " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
            If rs_Desliq.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "No hay datos de liquidaciones anteriores (desliq)"
                End If
            End If
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
                'Si el desliq prorratea debo proporcionarlo
                If CBool(rs_Desliq!Dlprorratea) Then
                    Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    'Cuota = IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) - (Cuota + ((rs_Desliq!Dlmonto) - (IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)), rs_Desliq!Dlmonto))))
                End If
                rs_Desliq.MoveNext
            Loop

            ' ------------------------------------------------------------------------
        
        Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
            ' Busco la declaracion jurada
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas,descuit,desrazsoc FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                     " AND desano=" & Ret_Ano & _
                     " AND itenro = " & rs_Item!Itenro
            OpenRecordset StrSql, rs_Desmen
            
            Do While Not rs_Desmen.EOF
                If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                    If rs_Item!Itenro = 3 Then
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                    
                    Else
                        If rs_Desmen!desmenprorra = 0 Then 'no es parejito
                            Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + rs_Desmen!desmondec
                        Else
                            Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                        End If
                    End If
                        
                        'FGZ - 19/04/2004
                        If rs_Item!Itenro <= 4 Then
                            If Not EsNulo(rs_Desmen!descuit) Then
                                I = 11
                                If Not EsNulo(rs_Traza_gan!Cuit_entidad11) Then
                                    Distinto = rs_Traza_gan!Cuit_entidad11 <> rs_Desmen!descuit
                                End If
                                Do While (I <= Tope) And Distinto
                                    I = I + 1
                                    Select Case I
                                    Case 11:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad11), 0, rs_Traza_gan!Cuit_entidad11) <> rs_Desmen!descuit
                                    Case 12:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad12), 0, rs_Traza_gan!Cuit_entidad12) <> rs_Desmen!descuit
                                    Case 13:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad13), 0, rs_Traza_gan!Cuit_entidad13) <> rs_Desmen!descuit
                                    Case 14:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad14), 0, rs_Traza_gan!Cuit_entidad14) <> rs_Desmen!descuit
                                    End Select
                                Loop
                              
                                If I > Tope And I <= 14 Then
                                    StrSql = "UPDATE sim_traza_gan SET "
                                    StrSql = StrSql & " cuit_entidad" & I & "='" & rs_Desmen!descuit & "',"
                                    StrSql = StrSql & " entidad" & I & "='" & rs_Desmen!DesRazsoc & "',"
                                    StrSql = StrSql & " monto_entidad" & I & "=" & rs_Desmen!desmondec
                                    StrSql = StrSql & " WHERE "
                                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    'FGZ - 22/12/2004
                                    'Leo la tabla
                                    'FGZ - 08/06/2012 ---------
                                    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
                                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    'StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                    OpenRecordset StrSql, rs_Traza_gan
                                    
                                    
                                    Tope = Tope + 1
                                Else
                                    If I = 15 Then
                                        Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!Empleg ' empleado.empleg
                                    Else
                                        StrSql = "UPDATE sim_traza_gan SET "
                                        StrSql = StrSql & " monto_entidad" & I & "= monto_entidad" & I & " + " & rs_Desmen!desmondec
                                        StrSql = StrSql & " WHERE "
                                        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                        StrSql = StrSql & " AND empresa =" & NroEmp
                                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                        
                                        'FGZ - 22/12/2004
                                        'Leo la tabla
                                        'FGZ - 08/06/2012 ---------------
                                        StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
                                        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                        'StrSql = StrSql & " AND empresa =" & NroEmp
                                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                        'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                        OpenRecordset StrSql, rs_Traza_gan
                                    End If
                                End If
                            Else
                                Total_Empresa = Total_Empresa + rs_Desmen!desmondec
                            End If
                        End If
                        'FGZ - 19/04/2004
                    End If
                
                
                rs_Desmen.MoveNext
            Loop
            
            'Busco las liquidaciones anteriores
            'FGZ - 08/06/2012 ------------
            StrSql = "SELECT dlmonto,dlprorratea,dlfecha FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
            If rs_Desliq.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "No hay datos de liquidaciones anteriores (desliq)"
                End If
            End If
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
                'Si el desliq prorratea debo proporcionarlo
                If CBool(rs_Desliq!Dlprorratea) Then
                    Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    Cuota = IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) - (Cuota + ((rs_Desliq!Dlmonto) - (IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)), rs_Desliq!Dlmonto))))
                End If
                Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
    
                rs_Desliq.MoveNext
            Loop
            
            'Busco los acumuladores de la liquidacion
            ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
            'FGZ - 08/06/2012 ----------
            StrSql = "SELECT acunro,itaprorratea,itasigno FROM itemacum " & _
                     " WHERE itenro =" & rs_Item!Itenro & _
                     " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
            OpenRecordset StrSql, rs_itemacum
            Do While Not rs_itemacum.EOF
                Acum = CStr(rs_itemacum!acuNro)
                If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                    Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
            
                    If CBool(rs_itemacum!itaprorratea) And (sinprorrateo = 0) Then
                        If CBool(rs_itemacum!itasigno) Then
                            Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + Aux_Acu_Monto
                            Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                        Else
                            Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - Aux_Acu_Monto
                            Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
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
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT itcprorratea,itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                     " AND itemconc.itenro =" & rs_Item!Itenro & _
                     " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
            OpenRecordset StrSql, rs_itemconc
            Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcprorratea) And (sinprorrateo = 0) Then
                    If CBool(rs_itemconc!itcsigno) Then
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + rs_itemconc!dlimonto
                        Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                    Else
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - rs_itemconc!dlimonto
                        Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
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
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT vimonto FROM valitem WHERE viano = " & Ret_Ano & _
                     " AND vimes = " & Ret_Mes & _
                     " AND itenro =" & rs_Item!Itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_Item!Itenro) = rs_valitem!vimonto
             
                rs_valitem.MoveNext
             Loop
            
            'Busco la declaracion Jurada
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
            'FGZ - 08/06/2012 ------------------------------
            StrSql = "SELECT dlmonto FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
    
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
    
                rs_Desliq.MoveNext
            Loop
            
            'Busco los acumuladores de la liquidacion
            ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT acunro,itasigno FROM itemacum " & _
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
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                'FGZ - 08/06/2012 -------------
                StrSql = "SELECT vimonto FROM valitem WHERE viano = " & Ret_Ano & _
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
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas,descuit,desrazsoc FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                        StrSql = "UPDATE sim_traza_gan SET "
                        StrSql = StrSql & " cuit_entidad9 ='" & rs_Desmen!descuit & "',"
                        StrSql = StrSql & " entidad9='" & rs_Desmen!DesRazsoc & "',"
                        StrSql = StrSql & " monto_entidad9=" & rs_Desmen!desmondec
                        StrSql = StrSql & " WHERE "
                        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                        StrSql = StrSql & " AND empresa =" & NroEmp
                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        'FGZ - 22/12/2004
                        'Leo la tabla
                        'FGZ - 08/06/2012 ------------------
                        StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
                        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                        'StrSql = StrSql & " AND empresa =" & NroEmp
                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                        'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                        OpenRecordset StrSql, rs_Traza_gan
                        
                        Tope = Tope + 1
                    End If
                End If
                'FGZ - 08/10/2013 -----------------------------------------------------------------
                ' Se saca el 23/05/2006
                'If (rs_Item!Itenro = 29) Or (rs_Item!Itenro = 55) Then 'Impuesto al debito bancario
                'le agrego item 56  'Compras en exterior
                If (rs_Item!Itenro = 29) Or (rs_Item!Itenro = 55) Or (rs_Item!Itenro = 56) Then
                    StrSql = "UPDATE sim_traza_gan SET "
                    StrSql = StrSql & " promo =" & val_impdebitos
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    'StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                
                    'FGZ - 22/12/2004
                    'Leo la tabla
                    'FGZ - 08/06/2012 ------------------
                    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE " & _
                    "pliqnro =" & buliq_periodo!PliqNro & _
                    " AND pronro =" & buliq_proceso!pronro & _
                    " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro & _
                    " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago) & _
                    " AND ternro =" & buliq_empleado!Ternro
                    OpenRecordset StrSql, rs_Traza_gan
                End If
                ' FGZ - 19/04/2004
                
                rs_Desmen.MoveNext
             Loop
        
            ''FGZ - 08/06/2012 ------------------
            'Busco las liquidaciones anteriores
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT dlmonto FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
    
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
    
                rs_Desliq.MoveNext
            Loop
            
            'Busco los acumuladores de la liquidacion
            ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT acunro,itasigno FROM itemacum " & _
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
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
            'If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            ' Maxi 13/12/2010 Cuando hay dif de plan item13 y devuelve tiene que restar el valor liquidado por eso saco el ABS de LIQ
            
            'FGZ - 05/06/2013 -----------------------------------------------------------------------------------------------------
            'Restauro otra vez lo de ABS porque estaba generando problemas cuando el monto no viene por ddjj sino que viene por liq
            'If Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
                'Items_TOPE(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
                Items_TOPE(rs_Item!Itenro) = Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            Else
                'FGZ - 24/08/2005
                If Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) = 0 Then
                    Items_TOPE(rs_Item!Itenro) = 0
                End If
                'FGZ - 24/08/2005
            End If
            'FGZ - 05/06/2013 -----------------------------------------------------------------------------------------------------
            
            
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
   End If


    ' SI SE TOMA PARA LA GANANCIA NETA, DA VUELTA EL SIGNO DEL TOPE SOLO PARA ITEMS
    ' QUE SE TOPEAN DE ALGUNA FORMA Y NO SALEN DEL RECIBO DE SUELDO.
    ' "Como saber que no sale del Recibo" ?
    
    If rs_Item!Itenro > 7 Then
        Items_TOPE(rs_Item!Itenro) = IIf(CBool(rs_Item!itesigno), Items_TOPE(rs_Item!Itenro), Abs(Items_TOPE(rs_Item!Itenro)))
    End If


    ''EAM (6.73) - Controlo los item que si son zona desfavorable son incrementado
    Select Case rs_Item!Itenro
        Case 10, 11:
                'EAM (6.73) - Aumento la escala para zona desfavorable en el procentaje del parámetro 1008
                valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                If (Por_Deduccion_zona > 0) Then
                    Items_TOPE(rs_Item!Itenro) = ((((valor_ant) * (1 + (Por_Deduccion_zona / 100))))) * (Items_DDJJ(rs_Item!Itenro) / Ret_Mes)
                Else
                    Items_TOPE(rs_Item!Itenro) = (valor_ant) * (Items_DDJJ(rs_Item!Itenro) / Ret_Mes)
                End If

        Case 16, 17, 31:
                'EAM (6.73) - Aumento la escala para zona desfavorable en el procentaje del parámetro 1008
                If (Por_Deduccion_zona > 0) Then
                    valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                    Items_TOPE(rs_Item!Itenro) = valor_ant + ((valor_ant) * ((Por_Deduccion_zona / 100)))
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + Abs(Items_LIQ(rs_Item!Itenro))
                End If
                
                If (rs_Item!Itenro = 31) Then Items_TOPE(rs_Item!Itenro) = IIf(Items_DDJJ(31) > Items_TOPE(rs_Item!Itenro), Items_TOPE(rs_Item!Itenro), Items_DDJJ(31))
                
    End Select
    'FGZ - 05/02/2014 -------------------------------------------------------------------------------------------------------------------


    
    'FGZ - 12/05/2015 -------------------------------------------------------------------
    If rs_Item!Itenro <> 29 And rs_Item!Itenro <> 55 And rs_Item!Itenro <> 56 Then
        If CBool(USA_DEBUG) Then
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-DDJJ" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_DDJJ(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Liq" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_LIQ(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-LiqAnt" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_OLD_LIQ(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Prorr" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_PRORR(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-ProrrCuota" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_PRORR_CUOTA(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Tope" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_TOPE(rs_Item!Itenro)
        End If
        If HACE_TRAZA Then
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_LIQ(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-LiqAnt"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_OLD_LIQ(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Prorr"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-ProrrCuota"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR_CUOTA(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Tope"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(rs_Item!Itenro))
        End If
    End If
    'FGZ - 12/05/2015 -------------------------------------------------------------------
    
    
    'EAM (6.57) - Si el item es 19 no lo tengo en cuenta en el cálculo de ganancia imponible
    'EAM (6.75) - Si el item es 34,35,36,37 no lo tengo en cuenta en el cálculo de ganancia imponible
    If (rs_Item!Itenro <> 19) And (rs_Item!Itenro <> 34) And (rs_Item!Itenro <> 35) And (rs_Item!Itenro <> 36) And (rs_Item!Itenro <> 37) Then
        'Calcula la Ganancia Imponible
        If CBool(rs_Item!itesigno) Then
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
    End If
      
SiguienteItem:
    rs_Item.MoveNext
Loop
  
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Bruta: " & (Gan_Imponible - Descuentos + Items_TOPE(50))
        Flog.writeline Espacios(Tabulador * 3) & "9- Gan. Bruta - CMA y DONA.: " & Gan_Imponible
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Neta: " & (Gan_Imponible + Deducciones)
        Flog.writeline Espacios(Tabulador * 3) & "9- Total Deducciones: " & Deducciones
        Flog.writeline Espacios(Tabulador * 3) & "9- Total art. 23: " & Ded_a23
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Bruta ", Gan_Imponible - Descuentos + Items_TOPE(100))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Gan. Bruta - CMA y DONA.", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Neta ", (Gan_Imponible + Deducciones))
        'Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Para Machinea ", (Gan_Imponible + Deducciones - Items_TOPE(100)))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total art. 23", Ded_a23)
    End If
            
    
    ' Calculo el porcentaje de deduccion segun la ganancia neta
    'Uso el campo para guardar la ganancia neta para el 648
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganneta =" & Gan_Imponible
    StrSql = StrSql & ", ganimpo =" & (Gan_Imponible + Deducciones + Ded_a23)
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
  
    
    'Calculo la Ganancia imponible
    Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23


 
    If Gan_Imponible > 0 Then
        StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
                 " WHERE escmes =" & Ret_Mes & _
                 " AND escano =" & Ret_Ano & _
                 " AND escinf <= " & Gan_Imponible & _
                 " AND escsup >= " & Gan_Imponible
        OpenRecordset StrSql, rs_escala
        
        'EAM(6.73) - Reinicio la variable porque se utiliza arriba
        Aux_Acu_Monto = 0
        
        'EAM(6.75) - Las horas extras se calcula con los item 3,35,36,37
        Aux_Acu_Monto = Items_TOPE(34) + Items_TOPE(35) + Items_TOPE(36) + Items_TOPE(37)

    
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible " & Gan_Imponible
            Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible Total " & (Gan_Imponible + Aux_Acu_Monto)
            Flog.writeline Espacios(Tabulador * 3) & "9- Horas Extras: " & Gan_Imponible & " Valor del Horas Extras: " & Aux_Acu_Monto
        End If
        
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Gan_Imponible)
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible Total", (Gan_Imponible + Aux_Acu_Monto))
        End If
    
        'EAM(6.73) - Una vez que resolvio la escala sin Horas Extras le sumo a ganancia imponible las horas extras
        Gan_Imponible = (Gan_Imponible + Aux_Acu_Monto)
        
        If Not rs_escala.EOF Then
            Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
        Else
            Impuesto_Escala = 0
        End If
    Else
        Impuesto_Escala = 0
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Gan_Imponible)
    End If
        
    
  
    

    
   'EAM(6.73) - Lo saque porque en principio no se utiliza.
'    ' FGZ - 19/04/2004
'    Otros = 0
'    I = 18
'
'    Do While I <= 100
'        'FGZ - 22/07/2005
'        'el item 30 no debe sumar en otros
'        If I <> 30 Then
'            Otros = Otros + Abs(Items_TOPE(I))
'        End If
'        I = I + 1
'    Loop
    
    StrSql = "UPDATE sim_traza_gan SET "
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
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
                
    ' Calculo las retenciones ya realizadas
    Ret_Ant = 0
        
    'Armo Fecha hasta como el ultimo dia del mes
    If (Ret_Mes = 12) Then
        fechaFichaH = CDate("31/12/" & Ret_Ano)
    Else
        fechaFichaH = CDate("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1
    End If
    
    fechaFichaD = CDate("01/01/" & Ret_Ano)
    
    StrSql = "SELECT SUM(importe) monto FROM sim_ficharet " & _
             " WHERE empleado =" & buliq_empleado!Ternro & _
             " AND fecha <= " & ConvFecha(fechaFichaH) & _
             " AND fecha >= " & ConvFecha(fechaFichaD)
    OpenRecordset StrSql, rs_Ficharet
    If Not rs_Ficharet.EOF Then
        If Not IsNull(rs_Ficharet!Monto) Then
            Ret_Ant = Ret_Ant + rs_Ficharet!Monto
        End If
    End If
    
    'FGZ - 17/10/2013 ---------------------------------------------
    'Calcular la retencion
    Retencion = Impuesto_Escala - Ret_Ant
    
    'Calculo de Impuesto y Debitos Bancarios, solo aplica si el impuesto retiene, si devuelve para el otro año lo declarado para este item
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc", val_impdebitos)
    If val_impdebitos > Impuesto_Escala Then
        val_impdebitos = Impuesto_Escala
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc con Tope", val_impdebitos)
    End If
    
    Retencion = Retencion - val_impdebitos
    'FGZ - 17/10/2013 ---------------------------------------------
    
    
    ' Para el F649 va en el 9b
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " promo =" & val_impdebitos
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
            
    
    'FGZ - 30/12/2004
    'Determinar el saldo
    'FGZ - 08/06/2012 ------------------
    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Not rs_Traza_gan.EOF Then
        StrSql = "UPDATE sim_traza_gan SET "
        StrSql = StrSql & "  saldo =" & Retencion
        StrSql = StrSql & "  ,retenciones =" & Ret_Ant
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 22/12/2004
        'Leo la tabla
        'FGZ - 08/06/2012 ------------------
        StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        'StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Retenciones anteriores", Ret_Ant)
        If Gan_Imponible > 0 Then
                If Not rs_escala.EOF Then
                    rs_escala.MoveFirst
                    If Not rs_escala.EOF Then
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", rs_escala!escporexe)
                    Else
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", 0)
                    End If
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver, x Tope General", Retencion)
            End If
        End If
        Monto = -Retencion
    Else
        Monto = 0
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "La Retencion es " & Monto
    End If
    
    'Monto = -Retencion
    Bien = True
        
    'Retenciones / Devoluciones
    If Retencion <> 0 Then
        Call InsertarFichaRet(buliq_empleado!Ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
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
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope
            
            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET monto =" & Items_TOPE(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
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
        
        If Items_PRORR(I) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
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
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
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
                'StrSql = "UPDATE traza_gan_item_top SET ddjj =" & Items_DDJJ(I)
                'StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
                'StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                'StrSql = StrSql & " AND empresa =" & NroEmp
                'StrSql = StrSql & " AND itenro =" & I
                
                StrSql = "UPDATE sim_traza_gan_item_top SET ddjj =" & Items_DDJJ(I) & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND empresa =" & NroEmp & _
                " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR_CUOTA(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top " & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR_CUOTA(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET prorr =" & Items_PRORR_CUOTA(I) & _
                    " WHERE ternro =" & buliq_empleado!Ternro & _
                    " AND pronro =" & buliq_proceso!pronro & _
                    " AND empresa =" & NroEmp & _
                    " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_old_liq para poder usarlo en el reporte de Ganancias
        If Items_OLD_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top " & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET old_liq =" & Items_OLD_LIQ(I) & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND empresa =" & NroEmp & _
                " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_liq para poder usarlo en el reporte de Ganancias
        If Items_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top " & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND itenro =" & I
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
                StrSql = "UPDATE sim_traza_gan_item_top SET liq =" & Items_LIQ(I) & _
                    " WHERE ternro =" & buliq_empleado!Ternro & _
                    " AND pronro =" & buliq_proceso!pronro & _
                    " AND empresa =" & NroEmp & _
                    " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 31/03/2005
        
        I = I + 1
    Loop

    exito = Bien
    for_Ganancias2017 = Monto
End Function



Public Function for_Ganancias2013(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias 2013.
' Autor      :
' Fecha      :
' Ultima Mod.: 17/09/2013
' Descripcion: nueva formula 2013.
' Ultima Mod.: 17/10/2013
' Ultima Mod.: 12/05/2015
' ---------------------------------------------------------------------------------------------
Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim p_prorratea As Integer      'Si prorratea o no para liq. finales
Dim p_sinprorrateo As Integer  'Indica que nunca prorratea
Dim p_brutomensual As Integer       'Acum Bruto mensual
Dim p_Deduccion_Zona
Dim p_Leyenda_Concepto As Long   'EAM (5.44)- Se usa para mostrar la Leyenda del concepto
Dim p_NetoSac2013 As Long           'Concepto de Neto  de SAC
Dim p_IncrementoItem31 As Long

'Variables Locales
Dim Devuelve As Double
Dim Tope_Gral As Double
Dim Neto As Double
Dim prorratea As Double
Dim sinprorrateo As Double
Dim Retencion As Double
Dim Gan_Imponible As Double
Dim Gan_Imponible_Grosada As Double
Dim Aux_Gan_Imponible_Grosada As Double
Dim Deducciones As Double
Dim Descuentos As Double
Dim Ded_a23 As Double
Dim Por_Deduccion As Double
Dim Impuesto_Escala As Double
Dim Ret_Ant As Double
Dim Ret_Ant_Agosto As Double
Dim Por_Deduccion_zona As Double
'FGZ - 12/05/2015 ----------------
Dim Aux_Por_Deduccion_zona As Double
'FGZ - 12/05/2015 ----------------
Dim Leyenda_Concepto As Double  'EAM (5.44) Leyenda_Concepto
Dim valor_ant As Double
Dim valor_act As Double 'sebastian stremel - 05/09/2013
Dim NetoSAC2013 As Double
Dim IncrementoItem31 As Boolean

Dim Ret_Mes As Integer
Dim Ret_Ano As Integer
Dim fin_mes_ret As Date
Dim ini_anyo_ret As Date
Dim Con_liquid As Integer
Dim I As Long
Dim j As Integer
Dim Texto As String

'Vectores para manejar el proceso
Dim Items_DDJJ(100) As Double
Dim Items_LIQ(100) As Double
Dim Items_PRORR(100) As Double
Dim Items_PRORR_CUOTA(100) As Double
Dim Items_OLD_LIQ(100) As Double
Dim Items_TOPE(100) As Double
Dim Items_ART_23(100) As Boolean

'Recorsets Auxiliares
'Dim rs_wf_tpa As New ADODB.Recordset
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
Dim rs_Aux As New ADODB.Recordset
Dim Hasta As Integer
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Double
Dim val_impdebitos As Double
Dim fechaFichaH As Date
Dim fechaFichaD As Date
Dim fechaFichaH2, fechaFichaD2 As Date
Dim Terminar As Boolean
'Dim pos1
'Dim pos2
Dim no_tiene_old As Boolean
Dim Z1, Z2, Z3 As Double
Dim CantZ1, CantZ2, CantZ3 As Long
Dim Total_Empresa As Double
Dim Tope As Integer
Dim rs_Traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Double
Dim Cuota As Double
Dim BrutoMensual As Double
Dim Acum_Maximo As Double
Dim Acum_MaximoAux As Double
Dim Tope16Liq As Double
Dim AuxDedEspecial As Double
Dim Ret_Aux As Double
Dim AuxInicio, AuxFin As Date
Dim Gan_Imponible_Agosto As Double
Dim Aux_Hasta, Aux_Desde As Date
Dim Beneficio As Boolean
Dim p_Beneficio As Long
Dim MantenerIncremento As Boolean
Dim ctrlItem20y56 As Boolean        'EAM (6.43) Controla los item 20 y 56 cuando es una liquidacion final o la fecha de pago es 31/12
Dim Extranjero As Boolean           'EAM (6.44) Controla si es expatriado
Dim p_Extranjero As Long            'EAM (6.44) Controla si es expatriado

Bien = False
Por_Deduccion_zona = 20
IncrementoItem31 = False
ctrlItem20y56 = False


'Comienzo
p_Devuelve = 1001
p_Tope_Gral = 1002
p_Neto = 1003
p_prorratea = 1005
p_sinprorrateo = 1006
p_brutomensual = 75 ' Maxi Ver bien el codigo
p_Deduccion_Zona = 1008 ' Maxi Ver bien el codigo
p_Leyenda_Concepto = 51 'EAM(5.44)- Parámetro nuevo leyenda
p_NetoSac2013 = 143 'Concepto de Neto de SAC 2013
p_IncrementoItem31 = 58 'Incremento Item 31
p_Beneficio = 1140      'Beneficio Item56
p_Extranjero = 1141     'Controla si es expatriado

Total_Empresa = 0
Tope = 10
Descuentos = 0
AuxDedEspecial = 0
Beneficio = False


' Primero limpio la traza
Call LimpiaTraza_Gan

StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
OpenRecordset StrSql, rs_Traza_gan
    
If HACE_TRAZA Then
    Call LimpiarTrazaConcepto(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro)
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
ini_anyo_ret = C_Date("01/01/" & Ret_Ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro
sinprorrateo = 0
BrutoMensual = 0

'Obtencion de los parametros de WorkFile
For I = LI_WF_Tpa To LS_WF_Tpa
    Select Case Arr_WF_TPA(I).tipoparam
    Case p_Devuelve:
        Devuelve = Arr_WF_TPA(I).Valor
    Case p_Tope_Gral:
        Tope_Gral = Arr_WF_TPA(I).Valor
    Case p_Neto:
        Neto = Arr_WF_TPA(I).Valor
    Case p_prorratea:
        prorratea = Arr_WF_TPA(I).Valor
    Case p_sinprorrateo:
        sinprorrateo = Arr_WF_TPA(I).Valor
    Case p_brutomensual:
        BrutoMensual = Arr_WF_TPA(I).Valor
    Case p_Deduccion_Zona:
        Por_Deduccion_zona = Arr_WF_TPA(I).Valor
    Case p_Leyenda_Concepto:
        Leyenda_Concepto = Arr_WF_TPA(I).Valor
    Case p_NetoSac2013:
        NetoSAC2013 = Arr_WF_TPA(I).Valor
    Case p_IncrementoItem31:
        IncrementoItem31 = CBool(Arr_WF_TPA(I).Valor)
    Case p_Beneficio:
        Beneficio = CBool(Arr_WF_TPA(I).Valor)
    Case p_Extranjero:
        Extranjero = CBool(Arr_WF_TPA(I).Valor)
    End Select
Next I

'Si es una liq. final no prorratea y tomo la escala de diciembre
If prorratea = 0 Then
    Ret_Mes = 12
    fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
End If

'EAM (v6.43) - Si la fecha de pago es 31/12 tiene que tener en cuenta el item 20 y 56, sino NO.
'EAM (6.57) - se agrego la condición (prorratea = 0) para que se tenga en cuenta tambien los item cuando es una liq. final
If CDate(buliq_proceso!profecpago) = CDate("31/12/" & Ret_Ano) Or (prorratea = 0) Then
    ctrlItem20y56 = True
End If


If Neto < 0 Then
   If CBool(USA_DEBUG) Then
      Flog.writeline Espacios(Tabulador * 3) & "El Neto del mes es negativo, se setea en cero."
   End If
   If HACE_TRAZA Then
      Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_Neto, "El Neto del Mes es negativo, se seteara en cero.", Neto)
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
    Flog.writeline Espacios(Tabulador * 3) & "Acum Bruto " & BrutoMensual
    
    Flog.writeline Espacios(Tabulador * 3) & "Beneficio devolucion Anticipada " & Beneficio
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Neto del Mes", Neto)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Acum Bruto", BrutoMensual)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Beneficio devolucion Anticipada", Beneficio)
End If

    'Limpiar items que suman al articulo 23
    For I = 1 To 100
        Items_ART_23(I) = False
    Next I
    val_impdebitos = 0
    'val_ComprasExt = 0
    Tope16Liq = 0
    
    
    'FGZ - 21/05/2015 -------------------------------------------------
    ''FGZ - 13/05/2015 -------------------------------------------------
    ''Se agregan un par de modificaciones / controles en la funcion que busca el bruto
    ''FGZ - 05/02/2014 -------------------------------------------------
    'Acum_Maximo = BuscarBrutoAgosto2013(BrutoMensual)
    ''FGZ - 05/02/2014 -------------------------------------------------
    ''FGZ - 13/05/2015 -------------------------------------------------
    
    'Acum_Maximo = BuscarBrutoAgosto2013(BrutoMensual, MantenerIncremento)
    'FGZ - 21/05/2015 -------------------------------------------------

    'EAM (v6.54) - Se agrego monto fijo para que no controle rangos de ganancia del 2015 (Antes de Macri)
    Acum_Maximo = 50000

' Recorro todos los items de Ganancias
'FGZ - 08/06/2012 ----------------
StrSql = "SELECT itenro,itetipotope,itesigno,iteitemstope,iteporctope,itenom FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_Item
Do While Not rs_Item.EOF
  'FGZ - 08/10/2013 -----------------------------------------------------------------------------------
  ' Impuestos y debitos Bancarios va como Promocion
  ' Ahora Compras en exterior tb
  If (rs_Item!Itenro = 29 Or rs_Item!Itenro = 55 Or rs_Item!Itenro = 56) Then
    If (Ret_Mes = 12) And (ctrlItem20y56 = True) Then
        StrSql = "SELECT desmondec FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                 " AND desano=" & Ret_Ano & _
                 " AND itenro = " & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
        'If Not rs_Desmen.EOF Then
        Do While Not rs_Desmen.EOF
           If rs_Item!Itenro = 29 Then
             'val_impdebitos = rs_Desmen!desmondec * 0.34
             val_impdebitos = val_impdebitos + (rs_Desmen!desmondec * 0.34)
            
            'FGZ - 16/12/2015 ------------------------
            Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.34)
            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.34)
             'FGZ - 16/12/2015 ------------------------
           Else
                'If rs_Item!Itenro = 23 Then
                If rs_Item!Itenro = 56 Then
                    'val_impdebitos = rs_Desmen!desmondec
                    val_impdebitos = val_impdebitos + rs_Desmen!desmondec
                    
                    'EAM (6.56) - Se comenta la linea para que tenga en cuenta el valor mas de una vez
                    'FGZ - 16/12/2015 ------------------------
                    'Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (rs_Desmen!desmondec)
                    'Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (rs_Desmen!desmondec)
                    'FGZ - 16/12/2015 ------------------------
                Else
                    'val_impdebitos = rs_Desmen!desmondec * 0.17
                    val_impdebitos = val_impdebitos + (rs_Desmen!desmondec * 0.17)

                    'FGZ - 16/12/2015 ------------------------
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.17)
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.17)
                    'FGZ - 16/12/2015 ------------------------

                End If
           End If
        'End If
            
            rs_Desmen.MoveNext
        Loop
        
        rs_Desmen.Close
    Else
        If rs_Item!Itenro = 56 Then
            If Beneficio Then
                StrSql = "SELECT sum(desmondec) total FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                         " AND desano=" & Ret_Ano & _
                         " AND Month(desfecdes) <= " & Ret_Mes & _
                         " AND itenro = " & rs_Item!Itenro
                OpenRecordset StrSql, rs_Desmen
                If Not rs_Desmen.EOF Then
                    'FGZ - 06/04/2015 --------------------------------------
                    'val_impdebitos = val_impdebitos + rs_Desmen!Total
                    val_impdebitos = val_impdebitos + IIf(IsNull(rs_Desmen!total), 0, rs_Desmen!total)
                    'FGZ - 06/04/2015 --------------------------------------
                    
                    'EAM (6.56) - Se comenta la linea para que tenga en cuenta el valor mas de una vez
                    'FGZ - 16/12/2015 ------------------------
                    'Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (IIf(IsNull(rs_Desmen!Total), 0, rs_Desmen!Total))
                    'Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (IIf(IsNull(rs_Desmen!Total), 0, rs_Desmen!Total))
                    'FGZ - 16/12/2015 ------------------------
                End If
            End If
        End If
    End If
    'FGZ - 19/01/2015 -----------------------------------------------
  Else
    
    'EAM (v6.40) - Solo se considera item 20 si es fin de año o final
    'EAM (v6.57) - Se agrego condicion para que se tenga en cuenta si es fin de año o es liquidacion final
                '(((Ret_Mes <> 12) And (ctrlItem20y56 = False)) Or ((Ret_Mes = 12) And (ctrlItem20y56 = False)))
    If (rs_Item!Itenro = 20) And (((Ret_Mes <> 12) And (ctrlItem20y56 = False)) Or ((Ret_Mes = 12) And (ctrlItem20y56 = False))) Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Item 20 No considerado"
        End If
    Else
        'EAM (6.44) - Si el Item es 10,11,12 o17 y Extranjero percibe su sueldo en Argentina. No se tiene en cuenta el item
        If ((rs_Item!Itenro = 17) Or (rs_Item!Itenro = 10) Or (rs_Item!Itenro = 11) Or (rs_Item!Itenro = 12)) And (Extranjero) Then
            Flog.writeline Espacios(Tabulador * 3) & "No se tiene en cuenta el Item " & rs_Item!Itenro & ". Extranjero que percibe su sueldo en Argentina"
            GoTo SiguienteItem
        End If
        
                
        Select Case rs_Item!itetipotope
        Case 1: ' el valor a tomar es lo que dice la escala
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT itenro,vimonto FROM valitem WHERE viano =" & Ret_Ano & _
                     " AND itenro=" & rs_Item!Itenro & _
                     " AND vimes =" & Ret_Mes
            OpenRecordset StrSql, rs_valitem
            
            Do While Not rs_valitem.EOF
                Items_DDJJ(rs_valitem!Itenro) = rs_valitem!vimonto
                Items_TOPE(rs_valitem!Itenro) = rs_valitem!vimonto
                
                rs_valitem.MoveNext
            Loop
    
        'Agregado Maxi 29/08/2013 -------------------------------------------------------------------------------------
         If rs_Item!Itenro = 16 Then
    
            'Busco los acumuladores de la liquidacion
            StrSql = "SELECT acunro,itasigno FROM itemacum " & _
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
    
            ' Busco los conceptos de la liquidacion
            StrSql = "SELECT itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
    
            'Busco las liquidaciones anteriores
            StrSql = "SELECT dlmonto,dlprorratea,dlfecha FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
            If rs_Desliq.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "No hay datos de liquidaciones anteriores (desliq)"
                End If
            End If
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
                'Si el desliq prorratea debo proporcionarlo
                If CBool(rs_Desliq!Dlprorratea) Then
                    Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    'Cuota = IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) - (Cuota + ((rs_Desliq!Dlmonto) - (IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)), rs_Desliq!Dlmonto))))
                End If
                'Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
    
                rs_Desliq.MoveNext
            Loop
    
            ' En el tope guardo lo de conceptos y en DDJJ lo de escala
            no_tiene_old = False
            If Items_OLD_LIQ(rs_Item!Itenro) = 0 Then
                no_tiene_old = True
                Items_OLD_LIQ(rs_Item!Itenro) = Items_DDJJ(16) - Abs(Items_LIQ(16))
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "Item " & rs_Item!Itenro & " igual a 0. " & " Valor Actual: " & Items_OLD_LIQ(rs_Item!Itenro)
                End If
            End If
            Items_TOPE(16) = Abs(Items_LIQ(16)) + Abs(Items_OLD_LIQ(rs_Item!Itenro))
    
            'FGZ - 02/01/2014 ------------------------------------------------------------------------------------
            If Ret_Ano > 2013 Then
                If Ret_Mes >= 1 Then
                    Items_OLD_LIQ(rs_Item!Itenro) = 0
                Else
                    valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, 1)
                    
                    StrSql = "SELECT vimonto FROM valitem WHERE viano = " & Ret_Ano & _
                    " AND vimes = " & (Ret_Mes - 1) & _
                    " AND itenro =" & rs_Item!Itenro
                    OpenRecordset StrSql, rs_valitem
                    If Not rs_valitem.EOF Then
                        Items_OLD_LIQ(rs_Item!Itenro) = rs_valitem!vimonto + ((valor_ant) * ((Por_Deduccion_zona / 100)))
                    End If
                End If
            End If
            'FGZ - 02/01/2014 ------------------------------------------------------------------------------------
           
           'FGZ - 30/12/2013 ------------------------------------------------------------------------------------
           'Entre 15 y 25000 aplica el 20% aumento de la escala del item 16 o el 30% si es zona patagonia
            If (Acum_Maximo > 15000 And Acum_Maximo <= 25000) Then
                    'FGZ - 12/05/2015 -------------------------------
                    Aux_Por_Deduccion_zona = Por_Deduccion_zona
                    If Not MantenerIncremento Then
                        Por_Deduccion_zona = PorcentajeRG3770(Por_Deduccion_zona, Acum_Maximo)
                    End If
                    'FGZ - 12/05/2015 -------------------------------
                        
                    If Ret_Ano > 2013 Then
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        'Items_TOPE(16) = Items_DDJJ(16) + ((valor_ant) * ((1 + Por_Deduccion_zona / 100)))
                        'Items_TOPE(16) = Items_DDJJ(16) + ((valor_ant) * ((Por_Deduccion_zona / 100)))
                        Items_TOPE(16) = valor_ant + ((valor_ant) * ((Por_Deduccion_zona / 100)))
                        'FGZ - 21/01/2014 ---------------------------------------------------
                        
                        'FGZ - 15/12/2014 -----------------------------------------
                        Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
                        'FGZ - 15/12/2014 -----------------------------------------
                    End If
                    
                    'FGZ - 12/05/2015 -------------------------------
                    Por_Deduccion_zona = Aux_Por_Deduccion_zona
                    'FGZ - 12/05/2015 -------------------------------
            End If
            'FGZ - 30/12/2013 ------------------------------------------------------------------------------------
        
            'FGZ - 30/12/2013 -----------------------------------------
            'EAM- eso se separo por el aguinaldo. Lo controla por separado
            If (Por_Deduccion_zona = 30 And Acum_Maximo > 25000) Then
                    If Ret_Ano > 2013 Then
                        'FGZ - 28/01/2014 -----------------------------------------
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        
                        'Items_TOPE(16) = (Abs(Items_LIQ(16)) * (1 + (Por_Deduccion_zona / 100))) + Abs(Items_OLD_LIQ(rs_Item!Itenro))
                        Items_TOPE(16) = valor_ant + ((valor_ant) * ((Por_Deduccion_zona / 100)))
                        'FGZ - 28/01/2014 -----------------------------------------
                        
                        'FGZ - 15/12/2014 -----------------------------------------
                        Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
                        'FGZ - 15/12/2014 -----------------------------------------
                    End If
            Else
                If Acum_Maximo > 25000 Then
                        ''FGZ - 15/12/2014 -----------------------------------------
                        'Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
                        ''FGZ - 15/12/2014 -----------------------------------------
                        
                        'FGZ - 20/02/2015 ---------------------------------
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        Items_TOPE(16) = valor_ant
                        Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
                        'FGZ - 20/02/2015 ---------------------------------
                End If
            End If
            'FGZ - 30/12/2013 -----------------------------------------
            
            If no_tiene_old = True Then
                Items_LIQ(16) = Items_LIQ(16) - Items_OLD_LIQ(16)
            End If
            
         End If
        ' End case 1
        ' ------------------------------------------------------------------------
        
        Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
            ' Busco la declaracion jurada
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas,descuit,desrazsoc FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                     " AND desano=" & Ret_Ano & _
                     " AND itenro = " & rs_Item!Itenro
            OpenRecordset StrSql, rs_Desmen
            
            Do While Not rs_Desmen.EOF
                If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                    If rs_Item!Itenro = 3 Then
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                    
                    Else
                        If rs_Desmen!desmenprorra = 0 Then 'no es parejito
                            Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + rs_Desmen!desmondec
                        Else
                            Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                        End If
                    End If
                        
                        'FGZ - 19/04/2004
                        If rs_Item!Itenro <= 4 Then
                            If Not EsNulo(rs_Desmen!descuit) Then
                                I = 11
                                If Not EsNulo(rs_Traza_gan!Cuit_entidad11) Then
                                    Distinto = rs_Traza_gan!Cuit_entidad11 <> rs_Desmen!descuit
                                End If
                                Do While (I <= Tope) And Distinto
                                    I = I + 1
                                    Select Case I
                                    Case 11:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad11), 0, rs_Traza_gan!Cuit_entidad11) <> rs_Desmen!descuit
                                    Case 12:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad12), 0, rs_Traza_gan!Cuit_entidad12) <> rs_Desmen!descuit
                                    Case 13:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad13), 0, rs_Traza_gan!Cuit_entidad13) <> rs_Desmen!descuit
                                    Case 14:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad14), 0, rs_Traza_gan!Cuit_entidad14) <> rs_Desmen!descuit
                                    End Select
                                Loop
                              
                                If I > Tope And I <= 14 Then
                                    StrSql = "UPDATE sim_traza_gan SET "
                                    StrSql = StrSql & " cuit_entidad" & I & "='" & rs_Desmen!descuit & "',"
                                    StrSql = StrSql & " entidad" & I & "='" & rs_Desmen!DesRazsoc & "',"
                                    StrSql = StrSql & " monto_entidad" & I & "=" & rs_Desmen!desmondec
                                    StrSql = StrSql & " WHERE "
                                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    'FGZ - 22/12/2004
                                    'Leo la tabla
                                    'FGZ - 08/06/2012 ---------
                                    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
                                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    'StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                    OpenRecordset StrSql, rs_Traza_gan
                                    
                                    
                                    Tope = Tope + 1
                                Else
                                    If I = 15 Then
                                        Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!Empleg ' empleado.empleg
                                    Else
                                        StrSql = "UPDATE sim_traza_gan SET "
                                        StrSql = StrSql & " monto_entidad" & I & "= monto_entidad" & I & " + " & rs_Desmen!desmondec
                                        StrSql = StrSql & " WHERE "
                                        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                        StrSql = StrSql & " AND empresa =" & NroEmp
                                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                        
                                        'FGZ - 22/12/2004
                                        'Leo la tabla
                                        'FGZ - 08/06/2012 ---------------
                                        StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
                                        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                        'StrSql = StrSql & " AND empresa =" & NroEmp
                                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                        'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                        OpenRecordset StrSql, rs_Traza_gan
                                    End If
                                End If
                            Else
                                Total_Empresa = Total_Empresa + rs_Desmen!desmondec
                            End If
                        End If
                        'FGZ - 19/04/2004
                    End If
                
                
                rs_Desmen.MoveNext
            Loop
            
            'Busco las liquidaciones anteriores
            'FGZ - 08/06/2012 ------------
            StrSql = "SELECT dlmonto,dlprorratea,dlfecha FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
            If rs_Desliq.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "No hay datos de liquidaciones anteriores (desliq)"
                End If
            End If
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
                'Si el desliq prorratea debo proporcionarlo
                If CBool(rs_Desliq!Dlprorratea) Then
                    Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    Cuota = IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) - (Cuota + ((rs_Desliq!Dlmonto) - (IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)), rs_Desliq!Dlmonto))))
                End If
                Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
    
                rs_Desliq.MoveNext
            Loop
            
            'Busco los acumuladores de la liquidacion
            ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
            'FGZ - 08/06/2012 ----------
            StrSql = "SELECT acunro,itaprorratea,itasigno FROM itemacum " & _
                     " WHERE itenro =" & rs_Item!Itenro & _
                     " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
            OpenRecordset StrSql, rs_itemacum
            Do While Not rs_itemacum.EOF
                Acum = CStr(rs_itemacum!acuNro)
                If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                    Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
            
                    If CBool(rs_itemacum!itaprorratea) And (sinprorrateo = 0) Then
                        If CBool(rs_itemacum!itasigno) Then
                            Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + Aux_Acu_Monto
                            Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                        Else
                            Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - Aux_Acu_Monto
                            Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
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
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT itcprorratea,itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                     " AND itemconc.itenro =" & rs_Item!Itenro & _
                     " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
            OpenRecordset StrSql, rs_itemconc
            Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcprorratea) And (sinprorrateo = 0) Then
                    If CBool(rs_itemconc!itcsigno) Then
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + rs_itemconc!dlimonto
                        Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                    Else
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - rs_itemconc!dlimonto
                        Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
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
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT vimonto FROM valitem WHERE viano = " & Ret_Ano & _
                     " AND vimes = " & Ret_Mes & _
                     " AND itenro =" & rs_Item!Itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_Item!Itenro) = rs_valitem!vimonto
             
                rs_valitem.MoveNext
             Loop
            
            'Busco la declaracion Jurada
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
            'FGZ - 08/06/2012 ------------------------------
            StrSql = "SELECT dlmonto FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
    
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
    
                rs_Desliq.MoveNext
            Loop
            
            'Busco los acumuladores de la liquidacion
            ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT acunro,itasigno FROM itemacum " & _
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
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                'FGZ - 08/06/2012 -------------
                StrSql = "SELECT vimonto FROM valitem WHERE viano = " & Ret_Ano & _
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
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas,descuit,desrazsoc FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                        StrSql = "UPDATE sim_traza_gan SET "
                        StrSql = StrSql & " cuit_entidad9 ='" & rs_Desmen!descuit & "',"
                        StrSql = StrSql & " entidad9='" & rs_Desmen!DesRazsoc & "',"
                        StrSql = StrSql & " monto_entidad9=" & rs_Desmen!desmondec
                        StrSql = StrSql & " WHERE "
                        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                        StrSql = StrSql & " AND empresa =" & NroEmp
                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        'FGZ - 22/12/2004
                        'Leo la tabla
                        'FGZ - 08/06/2012 ------------------
                        StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
                        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                        'StrSql = StrSql & " AND empresa =" & NroEmp
                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                        'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                        OpenRecordset StrSql, rs_Traza_gan
                        
                        Tope = Tope + 1
                    End If
                End If
                'FGZ - 08/10/2013 -----------------------------------------------------------------
                ' Se saca el 23/05/2006
                'If (rs_Item!Itenro = 29) Or (rs_Item!Itenro = 55) Then 'Impuesto al debito bancario
                'le agrego item 56  'Compras en exterior
                If (rs_Item!Itenro = 29) Or (rs_Item!Itenro = 55) Or (rs_Item!Itenro = 56) Then
                    StrSql = "UPDATE sim_traza_gan SET "
                    StrSql = StrSql & " promo =" & val_impdebitos
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    'StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                
                    'FGZ - 22/12/2004
                    'Leo la tabla
                    'FGZ - 08/06/2012 ------------------
                    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE " & _
                    "pliqnro =" & buliq_periodo!PliqNro & _
                    " AND pronro =" & buliq_proceso!pronro & _
                    " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro & _
                    " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago) & _
                    " AND ternro =" & buliq_empleado!Ternro
                    OpenRecordset StrSql, rs_Traza_gan
                End If
                ' FGZ - 19/04/2004
                
                rs_Desmen.MoveNext
             Loop
        
            ''FGZ - 08/06/2012 ------------------
            'Busco las liquidaciones anteriores
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT dlmonto FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
    
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
    
                rs_Desliq.MoveNext
            Loop
            
            'Busco los acumuladores de la liquidacion
            ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT acunro,itasigno FROM itemacum " & _
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
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
            'If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            ' Maxi 13/12/2010 Cuando hay dif de plan item13 y devuelve tiene que restar el valor liquidado por eso saco el ABS de LIQ
            
            'FGZ - 05/06/2013 -----------------------------------------------------------------------------------------------------
            'Restauro otra vez lo de ABS porque estaba generando problemas cuando el monto no viene por ddjj sino que viene por liq
            'If Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
                'Items_TOPE(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
                Items_TOPE(rs_Item!Itenro) = Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            Else
                'FGZ - 24/08/2005
                If Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) = 0 Then
                    Items_TOPE(rs_Item!Itenro) = 0
                End If
                'FGZ - 24/08/2005
            End If
            'FGZ - 05/06/2013 -----------------------------------------------------------------------------------------------------
            
            
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
   End If


    ' SI SE TOMA PARA LA GANANCIA NETA, DA VUELTA EL SIGNO DEL TOPE SOLO PARA ITEMS
    ' QUE SE TOPEAN DE ALGUNA FORMA Y NO SALEN DEL RECIBO DE SUELDO.
    ' "Como saber que no sale del Recibo" ?
    
    If rs_Item!Itenro > 7 Then
        Items_TOPE(rs_Item!Itenro) = IIf(CBool(rs_Item!itesigno), Items_TOPE(rs_Item!Itenro), Abs(Items_TOPE(rs_Item!Itenro)))
    End If


    'FGZ - 05/02/2014 -------------------------------------------------------------------------------------------------------------------
    'Cargas de familia
    Select Case rs_Item!Itenro
        Case 10, 11, 12:
            If Items_DDJJ(rs_Item!Itenro) > 0 Then
                If Ret_Ano = 2013 Then
                    'FGZ - 05/01/2015 --------------------------
                    'Se sacó todo el control que teniamos para 2013 dado que ya no tiene vigencia
                    'FGZ - 05/01/2015 --------------------------
                Else
                    If ((Acum_Maximo > 15000 And Acum_Maximo <= 25000) Or (Por_Deduccion_zona = 30 And Acum_Maximo > 15000)) Then
                    
                        'FGZ - 12/05/2015 -------------------------------
                        Aux_Por_Deduccion_zona = Por_Deduccion_zona
                        If Not MantenerIncremento Then
                            Por_Deduccion_zona = PorcentajeRG3770(Por_Deduccion_zona, Acum_Maximo)
                        End If
                        'FGZ - 12/05/2015 -------------------------------
                    
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        Items_TOPE(rs_Item!Itenro) = ((((valor_ant) * (1 + (Por_Deduccion_zona / 100))))) * (Items_DDJJ(rs_Item!Itenro) / Ret_Mes)
                    
                        'FGZ - 12/05/2015 -------------------------------
                        Por_Deduccion_zona = Aux_Por_Deduccion_zona
                        'FGZ - 12/05/2015 -------------------------------
                    Else
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        Items_TOPE(rs_Item!Itenro) = (valor_ant) * (Items_DDJJ(rs_Item!Itenro) / Ret_Mes)
                    End If
                End If
            End If
    End Select
    'FGZ - 05/02/2014 -------------------------------------------------------------------------------------------------------------------

    'EAM(5.44)- Se cambio condicion. se controa deduccion por zona y monto
    'If Acum_Maximo > 15000 And Acum_Maximo <= 25000 Then
    If ((Acum_Maximo > 15000 And Acum_Maximo <= 25000) Or (Por_Deduccion_zona = 30 And Acum_Maximo > 15000)) Then
    
        valor_ant = 0
        valor_act = 0 'falta definir
        
        'FGZ - 12/05/2015 -------------------------------
        Aux_Por_Deduccion_zona = Por_Deduccion_zona
        If Not MantenerIncremento Then
            Por_Deduccion_zona = PorcentajeRG3770(Por_Deduccion_zona, Acum_Maximo)
        End If
        'FGZ - 12/05/2015 -------------------------------
        
        'FGZ - 12/05/2014 -----------------------------------------------------------
        If (rs_Item!Itenro = 17) Or (rs_Item!Itenro = 31) Then
            If Ret_Ano > 2013 Then
                valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                Items_TOPE(rs_Item!Itenro) = ((valor_ant) * ((1 + Por_Deduccion_zona / 100)))
                If rs_Item!Itenro = 31 Then Items_TOPE(rs_Item!Itenro) = IIf(Items_DDJJ(31) > Items_TOPE(rs_Item!Itenro), Items_TOPE(rs_Item!Itenro), Items_DDJJ(31))
            Else
                If (rs_Item!Itenro = 31 And Items_DDJJ(31) <> 0) Or (rs_Item!Itenro = 17) Then
                    StrSql = "SELECT itenro,vimonto FROM valitem WHERE viano =" & Ret_Ano & _
                            " AND itenro=" & rs_Item!Itenro & " AND vimes =" & Ret_Mes - 1
                    OpenRecordset StrSql, rs_valitem
                    If Not rs_valitem.EOF Then
                        valor_ant = rs_valitem!vimonto
                        'Si el mes es mayor a 9, busco el valor acutal.
                        If Ret_Ano = 2013 And Ret_Mes > 9 Then
                            If rs_Item!Itenro = 31 Then
                                Items_TOPE(rs_Item!Itenro) = ((((ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes) - valor_ant) * (Por_Deduccion_zona / 100)) * (Ret_Mes - 8)) + ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes))
                                Items_TOPE(rs_Item!Itenro) = IIf(Items_DDJJ(31) > Items_TOPE(rs_Item!Itenro), Items_TOPE(rs_Item!Itenro), Items_DDJJ(31))
                            Else
                                Items_TOPE(rs_Item!Itenro) = ((((Items_TOPE(rs_Item!Itenro) - valor_ant) * (Por_Deduccion_zona / 100)) * (Ret_Mes - 8)) + Items_TOPE(rs_Item!Itenro))
                            End If
                        Else
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - (Items_TOPE(rs_Item!Itenro) - valor_ant) + ((Items_TOPE(rs_Item!Itenro) - valor_ant) * (1 + (Por_Deduccion_zona / 100)))
                        End If
                    End If
                End If
            End If
        End If
        'FGZ - 12/05/2014 -----------------------------------------------------------
        
        'FGZ - 12/05/2015 -------------------------------
        Por_Deduccion_zona = Aux_Por_Deduccion_zona
        'FGZ - 12/05/2015 -------------------------------
    End If
    
    'FGZ - 12/05/2015 -------------------------------------------------------------------
    If rs_Item!Itenro <> 29 And rs_Item!Itenro <> 55 And rs_Item!Itenro <> 56 Then
        If CBool(USA_DEBUG) Then
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-DDJJ" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_DDJJ(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Liq" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_LIQ(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-LiqAnt" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_OLD_LIQ(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Prorr" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_PRORR(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-ProrrCuota" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_PRORR_CUOTA(rs_Item!Itenro)
            Texto = CStr(rs_Item!Itenro) & "-" & rs_Item!itenom & "-Tope" & " "
            Flog.writeline Espacios(Tabulador * 3) & Texto & Items_TOPE(rs_Item!Itenro)
        End If
        If HACE_TRAZA Then
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_LIQ(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-LiqAnt"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_OLD_LIQ(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Prorr"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-ProrrCuota"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR_CUOTA(rs_Item!Itenro))
            Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Tope"
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(rs_Item!Itenro))
        End If
    End If
    'FGZ - 12/05/2015 -------------------------------------------------------------------
    
    
    'EAM (6.57) - Si el item es 19 no lo tengo en cuenta en el cálculo de ganancia imponible
    If (rs_Item!Itenro <> 19) Then
        'Calcula la Ganancia Imponible
        If CBool(rs_Item!itesigno) Then
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
    End If
      
SiguienteItem:
    rs_Item.MoveNext
Loop
  
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Bruta: " & (Gan_Imponible - Descuentos + Items_TOPE(50))
        Flog.writeline Espacios(Tabulador * 3) & "9- Gan. Bruta - CMA y DONA.: " & Gan_Imponible
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Neta: " & (Gan_Imponible + Deducciones)
        Flog.writeline Espacios(Tabulador * 3) & "9- Total Deducciones: " & Deducciones
        Flog.writeline Espacios(Tabulador * 3) & "9- Total art. 23: " & Ded_a23
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Bruta ", Gan_Imponible - Descuentos + Items_TOPE(100))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Gan. Bruta - CMA y DONA.", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Neta ", (Gan_Imponible + Deducciones))
        'Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Para Machinea ", (Gan_Imponible + Deducciones - Items_TOPE(100)))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total art. 23", Ded_a23)
    End If
            
    
    ' Calculo el porcentaje de deduccion segun la ganancia neta
    'Uso el campo para guardar la ganancia neta para el 648
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganneta =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
  
    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Ret_Ano >= 2000 And Gan_Imponible > 0 Then
        StrSql = "SELECT esd_porcentaje FROM escala_ded " & _
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "No hay esc. dedu para", Gan_Imponible)
            End If
            ' No se ha encontrado la escala de deduccion para el valor gan_imponible
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- % a tomar deduc." & Por_Deduccion
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- % a tomar deduc.", Por_Deduccion)
        End If
        
        'Aplico el porcentaje a las deducciones
        Ded_a23 = Ded_a23 * Por_Deduccion / 100
        
'        'Guardo el porcentaje de deduccion
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " porcdeduc =" & Por_Deduccion
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    End If
    ' Calculo la Ganancia imponible
    Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23

    'Menos de 15000 no paga-----------------------------------------------------------------------
    Ret_Aux = 0
     If Acum_Maximo <= 15000 Then
        If Ret_Aux <> 0 Then
            'Recalculo
            Gan_Imponible_Agosto = 0
            StrSql = "SELECT ganimpo FROM sim_traza_gan WHERE "
            'StrSql = StrSql & " concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
            'StrSql = StrSql & " AND fecha_pago >=" & ConvFecha(AuxInicio)
            StrSql = StrSql & " fecha_pago >=" & ConvFecha(AuxInicio)
            StrSql = StrSql & " AND fecha_pago <=" & ConvFecha(AuxFin)
            StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " ORDER BY fecha_pago desc "
            OpenRecordset StrSql, rs_Traza_gan
            If Not rs_Traza_gan.EOF Then
                Gan_Imponible_Agosto = rs_Traza_gan!Ganimpo
            End If
           
            AuxDedEspecial = Gan_Imponible - Gan_Imponible_Agosto
            Gan_Imponible = Gan_Imponible - AuxDedEspecial
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Hay retenciones anteriores a Septiembre 2013 " & Ret_Aux
                Flog.writeline Espacios(Tabulador * 3) & "Ganancia Imponible a  Agosto 2013 = " & Gan_Imponible_Agosto
                Flog.writeline Espacios(Tabulador * 3) & "Ajuste de Deduccion especial = " & AuxDedEspecial
                Flog.writeline Espacios(Tabulador * 3) & "Nueva Ganancia Imponible = " & Gan_Imponible
            End If
            
            If AuxDedEspecial > 0 Then
                Items_TOPE(16) = Items_TOPE(16) + AuxDedEspecial
                If HACE_TRAZA Then
                    Texto = Format(CStr(16), "00") & "-Deducción Especial-Tope"
                    StrSql = "DELETE traza WHERE cliqnro = " & buliq_cabliq!cliqnro
                    StrSql = StrSql & " AND concnro = " & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND tpanro = 0"
                    StrSql = StrSql & " AND tradesc ='" & Texto & "'"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(16))
                End If
             End If
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Tope Deduccion especial = " & Items_TOPE(16)
            End If
        Else    'Igual
            Gan_Imponible = 0
            If CBool(USA_DEBUG) Then
                'Flog.writeline Espacios(Tabulador * 3) & "La Retencion es Cero porque el sueldo bruto es <= 15000"
            End If
         
            'EAM(5.44)- Inserta un concepto en detliq con valor 0. Sirve para mostrar leyenda
            'Insertar un concepto con valor 0.
            'El concepto debe ser configurable como un parametro mas de la formula. (si el concepto no existe ==> no insertar porque va a dar error)
            'El concepto se debe llamar Remuneración y/o Haber no sujeto al Impuesto a las Ganancias- Beneficio Decreto PEN 1242/2013
            StrSql = "SELECT concnro FROM concepto WHERE conccod= " & Leyenda_Concepto
            OpenRecordset StrSql, rs_Aux
            
            If rs_Aux.EOF Then
                Flog.writeline Espacios(Tabulador * 3) & "No se encuentra configurado el parametro 51 para insertar concepto 0"
            Else
                StrSql = "INSERT INTO sim_detliq (cliqnro,concnro,dlimonto,dlicant,ajustado,dlitexto,dliretro )" & _
                              " VALUES (" & buliq_cabliq!cliqnro & "," & rs_Aux!ConcNro & "," & 0 & "," & 0 & "," & -1 & ",''," & -1 & ")"
                     objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            'FGZ - 20/11/2013 ---------------------------------
            'Aquellos que perciban menos de $15.000.- de Enero a Agosto, la Deducción Especial (ITEM 16) se calcule de acuerdo a lo establecido en el Dec.1242/13 Art.1,
            'que dice lo siguiente:
            '
            'Artículo 1° - lncreméntase, respecto de las rentas mencionadas en los incisos a), b) y c) del artículo 79 de la Ley de Impuesto a las Ganancias,
            '   texto ordenado en 1997, y sus modificaciones, la deducción especial establecida en el inciso c) del artículo 23 de dicha Ley,
            '   hasta un monto equivalente al que surja de restar a la ganancia neta sujeta a impuesto las deducciones de los incisos a) y b) del mencionado artículo 23.
            
            'CALCULO
            '   ITEM01 + ITEM02 + ITEM03 - (resto de los items)
            '
            '   Resto de los items =
            '       ABS ITEM05 + ABS ITEM06 + ABS ITEM07 + ABS ITEM08 + ABS ITEM09 + ABS ITEM10 + ABS ITEM11 + ABS ITEM12 + ABS ITEM13
            '       + ABS ITEM15 + ABS ITEM16 + ABS ITEM17 + ABS ITEM20 + ABS ITEM23 + ABS ITEM24 + ABS ITEM31
    
            'Si este cálculo es menor a cero, este monto no debería sumarse al ítem 16, si es mayor a cero sí.
            AuxDedEspecial = Items_TOPE(1) + Items_TOPE(2) + Items_TOPE(3)
            AuxDedEspecial = AuxDedEspecial - Abs(Items_TOPE(5)) - Abs(Items_TOPE(6)) - Abs(Items_TOPE(7)) - Abs(Items_TOPE(8)) - Abs(Items_TOPE(9)) - Abs(Items_TOPE(10))
            AuxDedEspecial = AuxDedEspecial - Abs(Items_TOPE(11)) - Abs(Items_TOPE(12)) - Abs(Items_TOPE(13)) - Abs(Items_TOPE(15)) - Abs(Items_TOPE(16)) - Abs(Items_TOPE(17))
            AuxDedEspecial = AuxDedEspecial - Abs(Items_TOPE(20)) - Abs(Items_TOPE(23)) - Abs(Items_TOPE(24)) - Abs(Items_TOPE(31))
            
           
            If AuxDedEspecial > 0 Then
                Items_TOPE(16) = Items_TOPE(16) + AuxDedEspecial
                If HACE_TRAZA Then
                    Texto = Format(CStr(16), "00") & "-Deducción Especial-Tope"
            
                    'FGZ - 01/10/2013 - Borro antes de insertar nuevamente ---------------------
                    StrSql = "DELETE traza WHERE cliqnro = " & buliq_cabliq!cliqnro
                    StrSql = StrSql & " AND concnro = " & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND tpanro = 0"
                    StrSql = StrSql & " AND tradesc ='" & Texto & "'"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'FGZ - 01/10/2013 - Borro antes de insertar nuevamente ---------------------
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(16))
                End If
            End If
            'FGZ - 20/11/2013 ---------------------------------
        End If
     End If


    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia imponible para el 649
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganimpo =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    'FGZ - 08/06/2012 ------------------
    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    OpenRecordset StrSql, rs_Traza_gan
    
    
    'FGZ - 15/04/2014 ------------------------------------------------------------------------------------------------------------------
    'If CBool(USA_DEBUG) Then
    '    Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible" & Gan_Imponible
    'End If
    'If HACE_TRAZA Then
    '    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Gan_Imponible)
    'End If
    'FGZ - 15/04/2014 ------------------------------------------------------------------------------------------------------------------
    
    
    'FGZ - 05/01/2015 ------------------------------------------------------------------------------------------------------------------
    'FGZ - 30/12/2014 ------------------------------------------------------------------------------------------------------------------
    If Acum_Maximo <= 15000 Then
        If Gan_Imponible > 0 Then
            'FGZ - 08/06/2012 ------------------
            'Entrar en la escala con las ganancias acumuladas
            
            'FGZ - 25/02/2014 ---------------------------------------------------
            'StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
            '         " WHERE escmes =" & Ret_Mes & _
            '         " AND escano =" & Ret_Ano & _
            '         " AND escinf <= " & Gan_Imponible & _
            '         " AND escsup >= " & Gan_Imponible
            
            StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
                     " WHERE escmes =" & IIf(Ret_Aux <> 0, 8, Ret_Mes) & _
                     " AND escano =" & Ret_Ano & _
                     " AND escinf <= " & Gan_Imponible & _
                     " AND escsup >= " & Gan_Imponible
            OpenRecordset StrSql, rs_escala
            'FGZ - 25/02/2014 ---------------------------------------------------
            If Not rs_escala.EOF Then
                Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
            Else
                Impuesto_Escala = 0
            End If
        Else
            Impuesto_Escala = 0
        End If
                        
        'FGZ - 15/04/2014 ------------------------------------------------------------------------------------------------------------------
        Gan_Imponible_Grosada = 0
        Aux_Gan_Imponible_Grosada = 0
        'Entro ahora en escala de diciembre
        'FGZ - 23/05/2014 -----------------------------------------------------
        'StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
        '         " WHERE escmes =" & IIf(Ret_Aux <> 0, 12, Ret_Mes) & _
        '         " AND escano =" & Ret_Ano & _
        '         " AND escinf <= " & Gan_Imponible & _
        '         " AND escsup >= " & Gan_Imponible
        StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
                 " WHERE escmes =" & IIf(Ret_Aux <> 0, 12, Ret_Mes) & _
                 " AND escano =" & Ret_Ano & _
                 " AND esccuota <= " & Impuesto_Escala & _
                 " ORDER BY esccuota DESC"
        'FGZ - 23/05/2014 -----------------------------------------------------
        OpenRecordset StrSql, rs_escala
        'FGZ - 25/02/2014 ---------------------------------------------------
        If Not rs_escala.EOF Then
            'Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
            Aux_Gan_Imponible_Grosada = ((((Impuesto_Escala - rs_escala!esccuota) * 100 / rs_escala!escporexe)) + rs_escala!escinf)
            Gan_Imponible_Grosada = ((((Impuesto_Escala - rs_escala!esccuota) * 100 / rs_escala!escporexe)) + rs_escala!escinf) - Gan_Imponible
        Else
            Gan_Imponible_Grosada = 0
        End If
            
            
    If AuxDedEspecial > 0 Then
        'AuxDedEspecial = AuxDedEspecial - Gan_Imponible_Grosada
        Items_TOPE(16) = Items_TOPE(16) - Gan_Imponible_Grosada
        AuxDedEspecial = Items_TOPE(16)
    Else
        'FGZ - 20/02/2015 ---------------------------------
        valor_ant = ValorEscala(16, Ret_Ano, Ret_Mes)
        Items_TOPE(16) = valor_ant
        Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
        Aux_Gan_Imponible_Grosada = AuxDedEspecial - Abs(Items_LIQ(16))
        'FGZ - 20/02/2015 ---------------------------------
    End If
        
        If HACE_TRAZA Then
            Texto = Format(CStr(16), "00") & "-Deducción Especial-Tope"
            StrSql = "DELETE traza WHERE cliqnro = " & buliq_cabliq!cliqnro
            StrSql = StrSql & " AND concnro = " & Buliq_Concepto(Concepto_Actual).ConcNro
            StrSql = StrSql & " AND tpanro = 0"
            StrSql = StrSql & " AND tradesc ='" & Texto & "'"
            objConn.Execute StrSql, , adExecuteNoRecords
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(16))
        
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible " & Aux_Gan_Imponible_Grosada
            End If
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Aux_Gan_Imponible_Grosada)
            End If
        End If
        'FGZ - 15/04/2014 ------------------------------------------------------------------------------------------------------------------
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible " & Gan_Imponible
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Gan_Imponible)
        End If
        If Gan_Imponible > 0 Then
            StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
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
    End If
    'FGZ - 30/12/2014 ------------------------------------------------------------------------------------------------------------------
    'FGZ - 05/01/2015 ------------------------------------------------------------------------------------------------------------------
            
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
    
    StrSql = "UPDATE sim_traza_gan SET "
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
    If AuxDedEspecial > 0 Then
        StrSql = StrSql & ", dedesp =" & AuxDedEspecial
    Else
        StrSql = StrSql & ", dedesp =" & (Items_TOPE(16))
    End If
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
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
                
    ' Calculo las retenciones ya realizadas
    Ret_Ant = 0
        
    'Armo Fecha hasta como el ultimo dia del mes
    If (Ret_Mes = 12) Then
        fechaFichaH = CDate("31/12/" & Ret_Ano)
    Else
        fechaFichaH = CDate("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1
    End If
    
    fechaFichaD = CDate("01/01/" & Ret_Ano)
    
    StrSql = "SELECT SUM(importe) monto FROM sim_ficharet " & _
             " WHERE empleado =" & buliq_empleado!Ternro & _
             " AND fecha <= " & ConvFecha(fechaFichaH) & _
             " AND fecha >= " & ConvFecha(fechaFichaD)
    OpenRecordset StrSql, rs_Ficharet
    If Not rs_Ficharet.EOF Then
        If Not IsNull(rs_Ficharet!Monto) Then
            Ret_Ant = Ret_Ant + rs_Ficharet!Monto
        End If
    End If
    
    'FGZ - 17/10/2013 ---------------------------------------------
    'Calcular la retencion
    Retencion = Impuesto_Escala - Ret_Ant
    
    'Calculo de Impuesto y Debitos Bancarios, solo aplica si el impuesto retiene, si devuelve para el otro año lo declarado para este item
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc", val_impdebitos)
    If val_impdebitos > Impuesto_Escala Then
        val_impdebitos = Impuesto_Escala
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc con Tope", val_impdebitos)
    End If
    
    Retencion = Retencion - val_impdebitos
    'FGZ - 17/10/2013 ---------------------------------------------
    
    
    ' Para el F649 va en el 9b
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " promo =" & val_impdebitos
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
            
    
    'FGZ - 30/12/2004
    'Determinar el saldo
    'FGZ - 08/06/2012 ------------------
    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Not rs_Traza_gan.EOF Then
        StrSql = "UPDATE sim_traza_gan SET "
        StrSql = StrSql & "  saldo =" & Retencion
        StrSql = StrSql & "  ,retenciones =" & Ret_Ant
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 22/12/2004
        'Leo la tabla
        'FGZ - 08/06/2012 ------------------
        StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        'StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
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
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Retenciones anteriores", Ret_Ant)
        If Gan_Imponible > 0 Then
                If Not rs_escala.EOF Then
                    rs_escala.MoveFirst
                    If Not rs_escala.EOF Then
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", rs_escala!escporexe)
                    Else
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", 0)
                    End If
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver, x Tope General", Retencion)
            End If
        End If
        Monto = -Retencion
    Else
        Monto = 0
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "La Retencion es " & Monto
    End If
    
    'Monto = -Retencion
    Bien = True
        
    'Retenciones / Devoluciones
    If Retencion <> 0 Then
        Call InsertarFichaRet(buliq_empleado!Ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
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
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope
            
            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET monto =" & Items_TOPE(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
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
        
        If Items_PRORR(I) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
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
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
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
                'StrSql = "UPDATE traza_gan_item_top SET ddjj =" & Items_DDJJ(I)
                'StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
                'StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                'StrSql = StrSql & " AND empresa =" & NroEmp
                'StrSql = StrSql & " AND itenro =" & I
                
                StrSql = "UPDATE sim_traza_gan_item_top SET ddjj =" & Items_DDJJ(I) & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND empresa =" & NroEmp & _
                " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR_CUOTA(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top " & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR_CUOTA(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET prorr =" & Items_PRORR_CUOTA(I) & _
                    " WHERE ternro =" & buliq_empleado!Ternro & _
                    " AND pronro =" & buliq_proceso!pronro & _
                    " AND empresa =" & NroEmp & _
                    " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_old_liq para poder usarlo en el reporte de Ganancias
        If Items_OLD_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top " & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET old_liq =" & Items_OLD_LIQ(I) & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND empresa =" & NroEmp & _
                " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_liq para poder usarlo en el reporte de Ganancias
        If Items_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top " & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND itenro =" & I
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
                StrSql = "UPDATE sim_traza_gan_item_top SET liq =" & Items_LIQ(I) & _
                    " WHERE ternro =" & buliq_empleado!Ternro & _
                    " AND pronro =" & buliq_proceso!pronro & _
                    " AND empresa =" & NroEmp & _
                    " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 31/03/2005
        
        I = I + 1
    Loop

    exito = Bien
    for_Ganancias2013 = Monto
End Function



Public Function for_Ganancias2013_OLD(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de retencion de ganancias 2013.
' Autor      :
' Fecha      :
' Ultima Mod.: 17/09/2013
' Descripcion: nueva formula 2013.
' Ultima Mod.: 17/10/2013
' ---------------------------------------------------------------------------------------------
Dim p_Devuelve As Integer    'si devuelve ganancia o no
Dim p_Tope_Gral As Integer  'Tope Gral de retencion
Dim p_Neto As Integer       'Base para el tope
Dim p_prorratea As Integer      'Si prorratea o no para liq. finales
Dim p_sinprorrateo As Integer  'Indica que nunca prorratea
Dim p_brutomensual As Integer       'Acum Bruto mensual
Dim p_Deduccion_Zona
Dim p_Leyenda_Concepto As Long   'EAM (5.44)- Se usa para mostrar la Leyenda del concepto
Dim p_NetoSac2013 As Long           'Concepto de Neto  de SAC
Dim p_IncrementoItem31 As Long

'Variables Locales
Dim Devuelve As Double
Dim Tope_Gral As Double
Dim Neto As Double
Dim prorratea As Double
Dim sinprorrateo As Double
Dim Retencion As Double
Dim Gan_Imponible As Double
Dim Gan_Imponible_Grosada As Double
Dim Aux_Gan_Imponible_Grosada As Double
Dim Deducciones As Double
Dim Descuentos As Double
Dim Ded_a23 As Double
Dim Por_Deduccion As Double
Dim Impuesto_Escala As Double
Dim Ret_Ant As Double
Dim Ret_Ant_Agosto As Double
Dim Por_Deduccion_zona As Double
'FGZ - 12/05/2015 ----------------
Dim Aux_Por_Deduccion_zona As Double
'FGZ - 12/05/2015 ----------------
Dim Leyenda_Concepto As Double  'EAM (5.44) Leyenda_Concepto
Dim valor_ant As Double
Dim valor_act As Double 'sebastian stremel - 05/09/2013
Dim NetoSAC2013 As Double
Dim IncrementoItem31 As Boolean

Dim Ret_Mes As Integer
Dim Ret_Ano As Integer
Dim fin_mes_ret As Date
Dim ini_anyo_ret As Date
Dim Con_liquid As Integer
Dim I As Long
Dim j As Integer
Dim Texto As String

'Vectores para manejar el proceso
Dim Items_DDJJ(100) As Double
Dim Items_LIQ(100) As Double
Dim Items_PRORR(100) As Double
Dim Items_PRORR_CUOTA(100) As Double
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
Dim rs_Aux As New ADODB.Recordset
Dim Hasta As Integer
Dim rs_acumulador As New ADODB.Recordset
Dim Acum As Long
Dim Aux_Acu_Monto As Double
Dim val_impdebitos As Double
Dim fechaFichaH As Date
Dim fechaFichaD As Date
Dim fechaFichaH2, fechaFichaD2 As Date
Dim Terminar As Boolean
Dim no_tiene_old As Boolean
Dim Z1, Z2, Z3 As Double
Dim CantZ1, CantZ2, CantZ3 As Long
Dim Total_Empresa As Double
Dim Tope As Integer
Dim rs_Traza_gan As New ADODB.Recordset
Dim Distinto As Boolean
Dim Otros As Double
Dim Cuota As Double
Dim BrutoMensual As Double
Dim Acum_Maximo As Double
Dim Acum_MaximoAux As Double
Dim Tope16Liq As Double
Dim AuxDedEspecial As Double
Dim Ret_Aux As Double
Dim AuxInicio, AuxFin As Date
Dim Gan_Imponible_Agosto As Double
Dim Aux_Hasta, Aux_Desde As Date
Dim Beneficio As Boolean
Dim p_Beneficio As Long
Dim MantenerIncremento As Boolean
Dim ctrlItem20y56 As Boolean        'EAM (6.43) Controla los item 20 y 56 cuando es una liquidacion final o la fecha de pago es 31/12
Dim Extranjero As Boolean           'EAM (6.44) Controla si es expatriado
Dim p_Extranjero As Long            'EAM (6.44) Controla si es expatriado

Bien = False
Por_Deduccion_zona = 20
IncrementoItem31 = False

'Comienzo
p_Devuelve = 1001
p_Tope_Gral = 1002
p_Neto = 1003
p_prorratea = 1005
p_sinprorrateo = 1006
p_brutomensual = 75 ' Maxi Ver bien el codigo
p_Deduccion_Zona = 1008 ' Maxi Ver bien el codigo
p_Leyenda_Concepto = 51 'EAM(5.44)- Parámetro nuevo leyenda
p_NetoSac2013 = 143 'Concepto de Neto de SAC 2013
p_IncrementoItem31 = 58 'Incremento Item 31
p_Beneficio = 1140      'Beneficio Item56

Total_Empresa = 0
Tope = 10
Descuentos = 0
AuxDedEspecial = 0
Beneficio = False

' Primero limpio la traza
Call LimpiaTraza_Gan

StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
OpenRecordset StrSql, rs_Traza_gan
    
If HACE_TRAZA Then
    'Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).ConcNro)
    Call LimpiarTrazaConcepto(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro)
End If

Ret_Mes = Month(buliq_proceso!profecpago)
Ret_Ano = Year(buliq_proceso!profecpago)
fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
ini_anyo_ret = C_Date("01/01/" & Ret_Ano)
Con_liquid = Buliq_Concepto(Concepto_Actual).ConcNro
sinprorrateo = 0
BrutoMensual = 0

''Obtencion de los parametros de WorkFile
'For I = LI_WF_Tpa To LS_WF_Tpa
'    Select Case Arr_WF_TPA(I).tipoparam
'    Case p_Devuelve:
'        Devuelve = Arr_WF_TPA(I).Valor
'    Case p_Tope_Gral:
'        Tope_Gral = Arr_WF_TPA(I).Valor
'    Case p_Neto:
'        Neto = Arr_WF_TPA(I).Valor
'    Case p_prorratea:
'        prorratea = Arr_WF_TPA(I).Valor
'    Case p_sinprorrateo:
'        sinprorrateo = Arr_WF_TPA(I).Valor
'    Case p_brutomensual:
'        BrutoMensual = Arr_WF_TPA(I).Valor
'    Case p_Deduccion_Zona:
'        Por_Deduccion_zona = Arr_WF_TPA(I).Valor
'    Case p_Leyenda_Concepto:
'        Leyenda_Concepto = Arr_WF_TPA(I).Valor
'    Case p_NetoSac2013:
'        NetoSAC2013 = Arr_WF_TPA(I).Valor
'    Case p_IncrementoItem31:
'        IncrementoItem31 = CBool(Arr_WF_TPA(I).Valor)
'    End Select
'Next I


StrSql = "SELECT * FROM " & TTempWF_tpa
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
    Case p_sinprorrateo:
        sinprorrateo = rs_wf_tpa!Valor
    Case p_brutomensual:
        BrutoMensual = rs_wf_tpa!Valor
    Case p_Deduccion_Zona:
        Por_Deduccion_zona = rs_wf_tpa!Valor
    Case p_Leyenda_Concepto:
        Leyenda_Concepto = rs_wf_tpa!Valor
    Case p_NetoSac2013:
        NetoSAC2013 = rs_wf_tpa!Valor
    Case p_IncrementoItem31:
        IncrementoItem31 = CBool(rs_wf_tpa!Valor)
    Case p_Beneficio:
        Beneficio = CBool(Arr_WF_TPA(I).Valor)
    Case p_Extranjero:
        Extranjero = CBool(Arr_WF_TPA(I).Valor)
    End Select
    
    rs_wf_tpa.MoveNext
Loop

'Si es una liq. final no prorratea y tomo la escala de diciembre
If prorratea = 0 Then
    Ret_Mes = 12
    fin_mes_ret = IIf(Ret_Mes = 12, C_Date("01/01/" & Ret_Ano + 1) - 1, C_Date("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1)
End If

'EAM (v6.43) - Si la fecha de pago es 31/12 tiene que tener en cuenta el item 20 y 56, sino NO.
If CDate(buliq_proceso!profecpago) = CDate("31/12/" & Ret_Ano) Then
    ctrlItem20y56 = True
End If

If Neto < 0 Then
   'If CBool(USA_DEBUG) Then
   '   Flog.writeline Espacios(Tabulador * 3) & "El Neto del mes es negativo, se setea en cero."
   'End If
   If HACE_TRAZA Then
      Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, p_Neto, "El Neto del Mes es negativo, se seteara en cero.", Neto)
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
    Flog.writeline Espacios(Tabulador * 3) & "Acum Bruto " & BrutoMensual
    Flog.writeline Espacios(Tabulador * 3) & "Beneficio devolucion Anticipada " & Beneficio
End If
If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Máxima Ret. en %", Tope_Gral)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Neto del Mes", Neto)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Acum Bruto", BrutoMensual)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 99999, "Beneficio devolucion Anticipada", Beneficio)
End If

    'Limpiar items que suman al articulo 23
    For I = 1 To 100
        Items_ART_23(I) = False
    Next I
    val_impdebitos = 0
    'val_ComprasExt = 0
    Tope16Liq = 0


    'FGZ - 21/05/2015 -------------------------------------------------
    ''FGZ - 13/05/2015 -------------------------------------------------
    ''Se agregan un par de modificaciones / controles en la funcion que busca el bruto
    ''FGZ - 05/02/2014 -------------------------------------------------
    'Acum_Maximo = BuscarBrutoAgosto2013(BrutoMensual)
    ''FGZ - 05/02/2014 -------------------------------------------------
    ''FGZ - 13/05/2015 -------------------------------------------------
    Acum_Maximo = BuscarBrutoAgosto2013(BrutoMensual, MantenerIncremento)
    'FGZ - 21/05/2015 -------------------------------------------------



' Recorro todos los items de Ganancias
'FGZ - 08/06/2012 ----------------
StrSql = "SELECT itenro,itetipotope,itesigno,iteitemstope,iteporctope,itenom FROM item ORDER BY itetipotope"
OpenRecordset StrSql, rs_Item
Do While Not rs_Item.EOF
  'FGZ - 08/10/2013 -----------------------------------------------------------------------------------
  ' Impuestos y debitos Bancarios va como Promocion
  ' Ahora Compras en exterior tb
    
  If (rs_Item!Itenro = 29 Or rs_Item!Itenro = 55 Or rs_Item!Itenro = 56) Then
    If Ret_Mes = 12 Then
        StrSql = "SELECT desmondec FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                 " AND desano=" & Ret_Ano & _
                 " AND itenro = " & rs_Item!Itenro
        OpenRecordset StrSql, rs_Desmen
        'If Not rs_Desmen.EOF Then
        Do While Not rs_Desmen.EOF
           If rs_Item!Itenro = 29 Then
             'val_impdebitos = rs_Desmen!desmondec * 0.34
             val_impdebitos = val_impdebitos + (rs_Desmen!desmondec * 0.34)
            
            'FGZ - 16/12/2015 ------------------------
            Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.34)
            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.34)
            'FGZ - 16/12/2015 ------------------------

           Else
                'If rs_Item!Itenro = 23 Then
                If rs_Item!Itenro = 56 Then
                    'val_impdebitos = rs_Desmen!desmondec
                    val_impdebitos = val_impdebitos + rs_Desmen!desmondec
                    
                    'EAM (6.56) - Se comenta la linea para que tenga en cuenta el valor mas de una vez
                    'FGZ - 16/12/2015 ------------------------
                    'Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (rs_Desmen!desmondec)
                    'Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (rs_Desmen!desmondec)
                    'FGZ - 16/12/2015 ------------------------
                Else
                    'val_impdebitos = rs_Desmen!desmondec * 0.17
                    val_impdebitos = val_impdebitos + (rs_Desmen!desmondec * 0.17)
                    'FGZ - 16/12/2015 ------------------------
                    Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.34)
                    Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (rs_Desmen!desmondec * 0.34)
                    'FGZ - 16/12/2015 ------------------------
                End If
           End If
        'End If
            rs_Desmen.MoveNext
        Loop
        
        rs_Desmen.Close
    Else
        If rs_Item!Itenro = 56 Then
            If Beneficio Then
                StrSql = "SELECT sum(desmondec) total FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                         " AND desano=" & Ret_Ano & _
                         " AND Month(desfecdes) <= " & Ret_Mes & _
                         " AND itenro = " & rs_Item!Itenro
                OpenRecordset StrSql, rs_Desmen
                If Not rs_Desmen.EOF Then
                    'FGZ - 06/04/2015 --------------------------------------
                    'val_impdebitos = val_impdebitos + rs_Desmen!Total
                    val_impdebitos = val_impdebitos + IIf(IsNull(rs_Desmen!total), 0, rs_Desmen!total)
                    'FGZ - 06/04/2015 --------------------------------------
                            
                    'EAM (6.56) - Se comenta la linea para que tenga en cuenta el valor mas de una vez
                    'FGZ - 16/12/2015 ------------------------
                    'Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + (IIf(IsNull(rs_Desmen!Total), 0, rs_Desmen!Total))
                    'Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + (IIf(IsNull(rs_Desmen!Total), 0, rs_Desmen!Total))
                    'FGZ - 16/12/2015 ------------------------
                End If
            End If
        End If
    End If
    'FGZ - 19/01/2015 -----------------------------------------------
  Else
    
        'EAM (v6.40) - Solo se considera item 20 si es fin de año o final
    If rs_Item!Itenro = 20 And Ret_Mes <> 12 Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Item 20 No considerado"
        End If
    Else
        'EAM (6.44) - Si el Item es 10,11,12 o17 y Extranjero percibe su sueldo en Argentina. No se tiene en cuenta el item
        If ((rs_Item!Itenro = 17) Or (rs_Item!Itenro = 10) Or (rs_Item!Itenro = 11) Or (rs_Item!Itenro = 12)) And (Extranjero) Then
            Flog.writeline Espacios(Tabulador * 3) & "No se tiene en cuenta el Item " & rs_Item!Itenro & ". Extranjero que percibe su sueldo en Argentina"
            GoTo SiguienteItem
        End If
    
        Select Case rs_Item!itetipotope
        Case 1: ' el valor a tomar es lo que dice la escala
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT itenro,vimonto FROM valitem WHERE viano =" & Ret_Ano & _
                     " AND itenro=" & rs_Item!Itenro & _
                     " AND vimes =" & Ret_Mes
            OpenRecordset StrSql, rs_valitem
            
            Do While Not rs_valitem.EOF
                Items_DDJJ(rs_valitem!Itenro) = rs_valitem!vimonto
                Items_TOPE(rs_valitem!Itenro) = rs_valitem!vimonto
                
                rs_valitem.MoveNext
            Loop
    
        'Agregado Maxi 29/08/2013 -------------------------------------------------------------------------------------
         If rs_Item!Itenro = 16 Then
    
            'Busco los acumuladores de la liquidacion
            StrSql = "SELECT acunro,itasigno FROM itemacum " & _
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
    
            ' Busco los conceptos de la liquidacion
            StrSql = "SELECT itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
    
            'Busco las liquidaciones anteriores
            StrSql = "SELECT dlmonto,dlprorratea,dlfecha FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
            If rs_Desliq.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "No hay datos de liquidaciones anteriores (desliq)"
                End If
            End If
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
                'Si el desliq prorratea debo proporcionarlo
                If CBool(rs_Desliq!Dlprorratea) Then
                    Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    'Cuota = IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) - (Cuota + ((rs_Desliq!Dlmonto) - (IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)), rs_Desliq!Dlmonto))))
                End If
                'Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
    
                rs_Desliq.MoveNext
            Loop
    
            ' En el tope guardo lo de conceptos y en DDJJ lo de escala
            no_tiene_old = False
            If Items_OLD_LIQ(rs_Item!Itenro) = 0 Then
                no_tiene_old = True
                Items_OLD_LIQ(rs_Item!Itenro) = Items_DDJJ(16) - Abs(Items_LIQ(16))
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "Item " & rs_Item!Itenro & " igual a 0. " & " Valor Actual: " & Items_OLD_LIQ(rs_Item!Itenro)
                End If
            End If
            Items_TOPE(16) = Abs(Items_LIQ(16)) + Abs(Items_OLD_LIQ(rs_Item!Itenro))
    
            If CBool(USA_DEBUG) Then
                'FGZ - 24/02/2014 ---- le saqué los comentarios porque ya no comila por longitud -------
                'Flog.writeline Espacios(Tabulador * 3) & "Item OLD Liq Valor Actual: " & Items_OLD_LIQ(rs_Item!Itenro)
                'Flog.writeline Espacios(Tabulador * 3) & "Item Liq 16 Valor Actual: " & Abs(Items_LIQ(16))
                'Flog.writeline Espacios(Tabulador * 3) & "Item Tope 16 Valor Actual: " & Items_TOPE(16)
                'FGZ - 24/02/2014 ---- le saqué los comentarios porque ya no comila por longitud -------
            End If
           
            'FGZ - 02/01/2014 ------------------------------------------------------------------------------------
            If Ret_Ano > 2013 Then
                If Ret_Mes >= 1 Then
                    Items_OLD_LIQ(rs_Item!Itenro) = 0
                Else
                    valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, 1)
                    
                    StrSql = "SELECT vimonto FROM valitem WHERE viano = " & Ret_Ano & _
                    " AND vimes = " & (Ret_Mes - 1) & _
                    " AND itenro =" & rs_Item!Itenro
                    OpenRecordset StrSql, rs_valitem
                    If Not rs_valitem.EOF Then
                        Items_OLD_LIQ(rs_Item!Itenro) = rs_valitem!vimonto + ((valor_ant) * ((Por_Deduccion_zona / 100)))
                    End If
                End If
            End If
            'FGZ - 02/01/2014 ------------------------------------------------------------------------------------
           
           'FGZ - 30/12/2013 ------------------------------------------------------------------------------------
           'Entre 15 y 25000 aplica el 20% aumento de la escala del item 16 o el 30% si es zona patagonia
            If (Acum_Maximo > 15000 And Acum_Maximo <= 25000) Then
                    'FGZ - 12/05/2015 -------------------------------
                    Aux_Por_Deduccion_zona = Por_Deduccion_zona
                    If Not MantenerIncremento Then
                        Por_Deduccion_zona = PorcentajeRG3770(Por_Deduccion_zona, Acum_Maximo)
                    End If
                    'FGZ - 12/05/2015 -------------------------------
                    
                    If Ret_Ano > 2013 Then
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        'Items_TOPE(16) = Items_DDJJ(16) + ((valor_ant) * ((1 + Por_Deduccion_zona / 100)))
                        'Items_TOPE(16) = Items_DDJJ(16) + ((valor_ant) * ((Por_Deduccion_zona / 100)))
                        Items_TOPE(16) = valor_ant + ((valor_ant) * ((Por_Deduccion_zona / 100)))
                        'FGZ - 21/01/2014 ---------------------------------------------------
                        
                        'FGZ - 15/12/2014 -----------------------------------------
                        Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
                        'FGZ - 15/12/2014 -----------------------------------------
                    End If
                    'FGZ - 12/05/2015 -------------------------------
                    Por_Deduccion_zona = Aux_Por_Deduccion_zona
                    'FGZ - 12/05/2015 -------------------------------
            End If
            'FGZ - 30/12/2013 ------------------------------------------------------------------------------------
            
        
            'FGZ - 30/12/2013 -----------------------------------------
            'EAM- eso se separo por el aguinaldo. Lo controla por separado
            If (Por_Deduccion_zona = 30 And Acum_Maximo > 25000) Then
                If Ret_Ano = 2013 And ((Ret_Mes > 9) Or (Ret_Mes = 9 And buliq_periodo!pliqmes = 9) Or (Ret_Mes = 9 And buliq_periodo!pliqmes = 8)) Then
                    'Items_TOPE(16) = (((Abs(Items_LIQ(16)) * ((Por_Deduccion_zona / 100)))) * (Ret_Mes - 8)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_LIQ(16))
                    'Items_TOPE(16) = (((Abs(Items_LIQ(16)) * ((Por_Deduccion_zona / 100)))) * (Ret_Mes - 8)) + Abs(Items_DDJJ(16)) + Abs(Items_LIQ(16))
                    Items_TOPE(16) = (((Abs(6220.8) * ((Por_Deduccion_zona / 100)))) * (Ret_Mes - 8)) + (Items_DDJJ(16)) + ValorConcepto(NetoSAC2013, 2013, 7, False)
                Else
                    If Ret_Ano > 2013 Then
                        'FGZ - 28/01/2014 -----------------------------------------
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        
                        'Items_TOPE(16) = (Abs(Items_LIQ(16)) * (1 + (Por_Deduccion_zona / 100))) + Abs(Items_OLD_LIQ(rs_Item!Itenro))
                        Items_TOPE(16) = valor_ant + ((valor_ant) * ((Por_Deduccion_zona / 100)))
                        'FGZ - 28/01/2014 -----------------------------------------
                        
                        'FGZ - 15/12/2014 -----------------------------------------
                        Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
                        'FGZ - 15/12/2014 -----------------------------------------
                        
                    End If
                    'FGZ - 21/01/2014 ---------------------------------------------------
                End If
            Else
                If Acum_Maximo > 25000 Then
                        ''FGZ - 15/12/2014 -----------------------------------------
                        'Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
                        ''FGZ - 15/12/2014 -----------------------------------------
                        
                        'FGZ - 20/02/2015 ---------------------------------
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        Items_TOPE(16) = valor_ant
                        Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
                        'FGZ - 20/02/2015 ---------------------------------
                End If
                
            End If
            'FGZ - 30/12/2013 -----------------------------------------
            
            If no_tiene_old = True Then
                Items_LIQ(16) = Items_LIQ(16) - Items_OLD_LIQ(16)
            End If
            
         End If
        ' End case 1
        ' ------------------------------------------------------------------------
        
        Case 2: 'Tomo los valores de DDJJ y Liquidacion sin Tope
            ' Busco la declaracion jurada
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas,descuit,desrazsoc FROM sim_desmen WHERE empleado =" & buliq_empleado!Ternro & _
                     " AND desano=" & Ret_Ano & _
                     " AND itenro = " & rs_Item!Itenro
            OpenRecordset StrSql, rs_Desmen
            
            Do While Not rs_Desmen.EOF
                If Month(rs_Desmen!desfecdes) <= Ret_Mes Then
                    If rs_Item!Itenro = 3 Then
                        Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                    
                    Else
                        If rs_Desmen!desmenprorra = 0 Then 'no es parejito
                            Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + rs_Desmen!desmondec
                        Else
                            Items_DDJJ(rs_Item!Itenro) = Items_DDJJ(rs_Item!Itenro) + rs_Desmen!desmondec
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((prorratea = 1) And (Ret_Mes <= Month(rs_Desmen!desfechas)), rs_Desmen!desmondec / (Month(rs_Desmen!desfechas) + 1 - Month(rs_Desmen!desfecdes)) * (Ret_Mes - Month(rs_Desmen!desfecdes) + 1), rs_Desmen!desmondec)
                        End If
                    End If
                        
                        'FGZ - 19/04/2004
                        If rs_Item!Itenro <= 4 Then
                            If Not EsNulo(rs_Desmen!descuit) Then
                                I = 11
                                If Not EsNulo(rs_Traza_gan!Cuit_entidad11) Then
                                    Distinto = rs_Traza_gan!Cuit_entidad11 <> rs_Desmen!descuit
                                End If
                                Do While (I <= Tope) And Distinto
                                    I = I + 1
                                    Select Case I
                                    Case 11:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad11), 0, rs_Traza_gan!Cuit_entidad11) <> rs_Desmen!descuit
                                    Case 12:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad12), 0, rs_Traza_gan!Cuit_entidad12) <> rs_Desmen!descuit
                                    Case 13:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad13), 0, rs_Traza_gan!Cuit_entidad13) <> rs_Desmen!descuit
                                    Case 14:
                                        Distinto = IIf(EsNulo(rs_Traza_gan!Cuit_entidad14), 0, rs_Traza_gan!Cuit_entidad14) <> rs_Desmen!descuit
                                    End Select
                                Loop
                              
                                If I > Tope And I <= 14 Then
                                    StrSql = "UPDATE sim_traza_gan SET "
                                    StrSql = StrSql & " cuit_entidad" & I & "='" & rs_Desmen!descuit & "',"
                                    StrSql = StrSql & " entidad" & I & "='" & rs_Desmen!DesRazsoc & "',"
                                    StrSql = StrSql & " monto_entidad" & I & "=" & rs_Desmen!desmondec
                                    StrSql = StrSql & " WHERE "
                                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                    'FGZ - 22/12/2004
                                    'Leo la tabla
                                    'FGZ - 08/06/2012 ---------
                                    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
                                    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                    'StrSql = StrSql & " AND empresa =" & NroEmp
                                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                    'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                    OpenRecordset StrSql, rs_Traza_gan
                                    
                                    
                                    Tope = Tope + 1
                                Else
                                    If I = 15 Then
                                        Flog.writeline "Verifique las desgravaciones declaradas para el legajo: " & buliq_empleado!Empleg ' empleado.empleg
                                    Else
                                        StrSql = "UPDATE sim_traza_gan SET "
                                        StrSql = StrSql & " monto_entidad" & I & "= monto_entidad" & I & " + " & rs_Desmen!desmondec
                                        StrSql = StrSql & " WHERE "
                                        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                        StrSql = StrSql & " AND empresa =" & NroEmp
                                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                        
                                        'FGZ - 22/12/2004
                                        'Leo la tabla
                                        'FGZ - 08/06/2012 ---------------
                                        StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
                                        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                                        'StrSql = StrSql & " AND empresa =" & NroEmp
                                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                                        'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                                        OpenRecordset StrSql, rs_Traza_gan
                                    End If
                                End If
                            Else
                                Total_Empresa = Total_Empresa + rs_Desmen!desmondec
                            End If
                        End If
                        'FGZ - 19/04/2004
                    End If
                
                
                rs_Desmen.MoveNext
            Loop
            
            'Busco las liquidaciones anteriores
            'FGZ - 08/06/2012 ------------
            StrSql = "SELECT dlmonto,dlprorratea,dlfecha FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
            If rs_Desliq.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 3) & "No hay datos de liquidaciones anteriores (desliq)"
                End If
            End If
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
                'Si el desliq prorratea debo proporcionarlo
                If CBool(rs_Desliq!Dlprorratea) Then
                    Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    Cuota = IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)) / (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
                    Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) - (Cuota + ((rs_Desliq!Dlmonto) - (IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), (rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1)), rs_Desliq!Dlmonto))))
                End If
                Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf((rs_Item!Itenro = 3 Or CBool(rs_Desliq!Dlprorratea)) And (prorratea = 1), rs_Desliq!Dlmonto / (13 - Month(rs_Desliq!Dlfecha)) * (Ret_Mes - Month(rs_Desliq!Dlfecha) + 1), rs_Desliq!Dlmonto)
    
                rs_Desliq.MoveNext
            Loop
            
            'Busco los acumuladores de la liquidacion
            ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
            'FGZ - 08/06/2012 ----------
            StrSql = "SELECT acunro,itaprorratea,itasigno FROM itemacum " & _
                     " WHERE itenro =" & rs_Item!Itenro & _
                     " AND (itaconcnrodest is null OR itaconcnrodest = " & Con_liquid & ")"
            OpenRecordset StrSql, rs_itemacum
            Do While Not rs_itemacum.EOF
                Acum = CStr(rs_itemacum!acuNro)
                If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acum)) Then
                    Aux_Acu_Monto = objCache_Acu_Liq_Monto.Valor(CStr(Acum))
            
                    If CBool(rs_itemacum!itaprorratea) And (sinprorrateo = 0) Then
                        If CBool(rs_itemacum!itasigno) Then
                            Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + Aux_Acu_Monto
                            Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
                        Else
                            Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - Aux_Acu_Monto
                            Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) - IIf(prorratea = 1, Aux_Acu_Monto / (13 - Ret_Mes), Aux_Acu_Monto)
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
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT itcprorratea,itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
                     " AND itemconc.itenro =" & rs_Item!Itenro & _
                     " AND (itemconc.itcconcnrodest is null OR itemconc.itcconcnrodest = " & Con_liquid & ")"
            OpenRecordset StrSql, rs_itemconc
            Do While Not rs_itemconc.EOF
                If CBool(rs_itemconc!itcprorratea) And (sinprorrateo = 0) Then
                    If CBool(rs_itemconc!itcsigno) Then
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) + rs_itemconc!dlimonto
                        Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                        Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) + IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
                    Else
                        Items_PRORR(rs_Item!Itenro) = Items_PRORR(rs_Item!Itenro) - rs_itemconc!dlimonto
                        Items_PRORR_CUOTA(rs_Item!Itenro) = Items_PRORR_CUOTA(rs_Item!Itenro) - IIf(prorratea = 1, rs_itemconc!dlimonto / (13 - Ret_Mes), rs_itemconc!dlimonto)
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
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT vimonto FROM valitem WHERE viano = " & Ret_Ano & _
                     " AND vimes = " & Ret_Mes & _
                     " AND itenro =" & rs_Item!Itenro
            OpenRecordset StrSql, rs_valitem
             Do While Not rs_valitem.EOF
                Items_TOPE(rs_Item!Itenro) = rs_valitem!vimonto
             
                rs_valitem.MoveNext
             Loop
            
            'Busco la declaracion Jurada
            'FGZ - 08/06/2012 ----------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
            'FGZ - 08/06/2012 ------------------------------
            StrSql = "SELECT dlmonto FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
    
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
    
                rs_Desliq.MoveNext
            Loop
            
            'Busco los acumuladores de la liquidacion
            ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT acunro,itasigno FROM itemacum " & _
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
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                'FGZ - 08/06/2012 -------------
                StrSql = "SELECT vimonto FROM valitem WHERE viano = " & Ret_Ano & _
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
            'FGZ - 08/06/2012 -------------
            StrSql = "SELECT desmondec,desmenprorra,desfecdes,desfechas,descuit,desrazsoc FROM sim_desmen WHERE empleado = " & buliq_empleado!Ternro & _
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
                        StrSql = "UPDATE sim_traza_gan SET "
                        StrSql = StrSql & " cuit_entidad9 ='" & rs_Desmen!descuit & "',"
                        StrSql = StrSql & " entidad9='" & rs_Desmen!DesRazsoc & "',"
                        StrSql = StrSql & " monto_entidad9=" & rs_Desmen!desmondec
                        StrSql = StrSql & " WHERE "
                        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                        StrSql = StrSql & " AND empresa =" & NroEmp
                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        'FGZ - 22/12/2004
                        'Leo la tabla
                        'FGZ - 08/06/2012 ------------------
                        StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
                        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                        'StrSql = StrSql & " AND empresa =" & NroEmp
                        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                        'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
                        OpenRecordset StrSql, rs_Traza_gan
                        
                        Tope = Tope + 1
                    End If
                End If
                'FGZ - 08/10/2013 -----------------------------------------------------------------
                ' Se saca el 23/05/2006
                'If (rs_Item!Itenro = 29) Or (rs_Item!Itenro = 55) Then 'Impuesto al debito bancario
                'le agrego item 56  'Compras en exterior
                If (rs_Item!Itenro = 29) Or (rs_Item!Itenro = 55) Or (rs_Item!Itenro = 56) Then
                    StrSql = "UPDATE sim_traza_gan SET "
                    StrSql = StrSql & " promo =" & val_impdebitos
                    StrSql = StrSql & " WHERE "
                    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
                    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
                    'StrSql = StrSql & " AND empresa =" & NroEmp
                    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                
                    'FGZ - 22/12/2004
                    'Leo la tabla
                    'FGZ - 08/06/2012 ------------------
                    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE " & _
                    "pliqnro =" & buliq_periodo!PliqNro & _
                    " AND pronro =" & buliq_proceso!pronro & _
                    " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro & _
                    " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago) & _
                    " AND ternro =" & buliq_empleado!Ternro
                    OpenRecordset StrSql, rs_Traza_gan
                End If
                ' FGZ - 19/04/2004
                
                rs_Desmen.MoveNext
             Loop
        
            ''FGZ - 08/06/2012 ------------------
            'Busco las liquidaciones anteriores
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT dlmonto FROM sim_desliq WHERE itenro =" & rs_Item!Itenro & _
                     " AND empleado = " & buliq_empleado!Ternro & _
                     " AND dlfecha >= " & ConvFecha(ini_anyo_ret) & _
                     " AND dlfecha <= " & ConvFecha(fin_mes_ret)
            OpenRecordset StrSql, rs_Desliq
    
            Do While Not rs_Desliq.EOF
                Items_OLD_LIQ(rs_Item!Itenro) = Items_OLD_LIQ(rs_Item!Itenro) + rs_Desliq!Dlmonto
    
                rs_Desliq.MoveNext
            Loop
            
            'Busco los acumuladores de la liquidacion
            ' FGZ - 05/03/2004 Nuevo Desde acá -------------------------
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT acunro,itasigno FROM itemacum " & _
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
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itcsigno,dlimonto FROM itemconc " & _
                     " INNER JOIN sim_detliq ON itemconc.concnro = sim_detliq.concnro " & _
                     " WHERE sim_detliq.cliqnro = " & buliq_cabliq!cliqnro & _
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
            'If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            ' Maxi 13/12/2010 Cuando hay dif de plan item13 y devuelve tiene que restar el valor liquidado por eso saco el ABS de LIQ
            
            'FGZ - 05/06/2013 -----------------------------------------------------------------------------------------------------
            'Restauro otra vez lo de ABS porque estaba generando problemas cuando el monto no viene por ddjj sino que viene por liq
            'If Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
            If Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) < Items_TOPE(rs_Item!Itenro) Then
                'Items_TOPE(rs_Item!Itenro) = Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
                Items_TOPE(rs_Item!Itenro) = Abs(Items_LIQ(rs_Item!Itenro)) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro))
            Else
                'FGZ - 24/08/2005
                If Items_LIQ(rs_Item!Itenro) + Abs(Items_OLD_LIQ(rs_Item!Itenro)) + Abs(Items_DDJJ(rs_Item!Itenro)) = 0 Then
                    Items_TOPE(rs_Item!Itenro) = 0
                End If
                'FGZ - 24/08/2005
            End If
            'FGZ - 05/06/2013 -----------------------------------------------------------------------------------------------------
            
            
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
    End If
    

    ' SI SE TOMA PARA LA GANANCIA NETA, DA VUELTA EL SIGNO DEL TOPE SOLO PARA ITEMS
    ' QUE SE TOPEAN DE ALGUNA FORMA Y NO SALEN DEL RECIBO DE SUELDO.
    ' "Como saber que no sale del Recibo" ?
    
    If rs_Item!Itenro > 7 Then
        Items_TOPE(rs_Item!Itenro) = IIf(CBool(rs_Item!itesigno), Items_TOPE(rs_Item!Itenro), Abs(Items_TOPE(rs_Item!Itenro)))
    End If


    'FGZ - 05/02/2014 -------------------------------------------------------------------------------------------------------------------
    'Cargas de familia
    Select Case rs_Item!Itenro
        Case 10, 11, 12:
            If Items_DDJJ(rs_Item!Itenro) > 0 Then
                If Ret_Ano = 2013 Then
                    'FGZ - 05/01/2015 ------------------------------------------------------------------------------------------------------------------
                    'se sacó porque ya no es necesario
                    'FGZ - 05/01/2015 ------------------------------------------------------------------------------------------------------------------
                Else
                    'valor_act = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                    'Items_TOPE(rs_Item!Itenro) = ((((valor_act - valor_ant) * (1 + (Por_Deduccion_zona / 100)))) + valor_ant) * (Items_DDJJ(rs_Item!Itenro) / Ret_Mes)
                    If ((Acum_Maximo > 15000 And Acum_Maximo <= 25000) Or (Por_Deduccion_zona = 30 And Acum_Maximo > 15000)) Then
                        'FGZ - 12/05/2015 -------------------------------
                        Aux_Por_Deduccion_zona = Por_Deduccion_zona
                        If Not MantenerIncremento Then
                            Por_Deduccion_zona = PorcentajeRG3770(Por_Deduccion_zona, Acum_Maximo)
                        End If
                        'FGZ - 12/05/2015 -------------------------------
                    
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        Items_TOPE(rs_Item!Itenro) = ((((valor_ant) * (1 + (Por_Deduccion_zona / 100))))) * (Items_DDJJ(rs_Item!Itenro) / Ret_Mes)
                        
                        'FGZ - 12/05/2015 -------------------------------
                        Por_Deduccion_zona = Aux_Por_Deduccion_zona
                        'FGZ - 12/05/2015 -------------------------------
                    Else
                        valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                        Items_TOPE(rs_Item!Itenro) = (valor_ant) * (Items_DDJJ(rs_Item!Itenro) / Ret_Mes)
                    End If
                End If
            End If
    End Select
    'FGZ - 05/02/2014 -------------------------------------------------------------------------------------------------------------------



    'EAM(5.44)- Se cambio condicion. se controa deduccion por zona y monto
    'If Acum_Maximo > 15000 And Acum_Maximo <= 25000 Then
    If ((Acum_Maximo > 15000 And Acum_Maximo <= 25000) Or (Por_Deduccion_zona = 30 And Acum_Maximo > 15000)) Then
    
        valor_ant = 0
        valor_act = 0 'falta definir
        
        'FGZ - 12/05/2015 -------------------------------
        Aux_Por_Deduccion_zona = Por_Deduccion_zona
        If Not MantenerIncremento Then
            Por_Deduccion_zona = PorcentajeRG3770(Por_Deduccion_zona, Acum_Maximo)
        End If
        'FGZ - 12/05/2015 -------------------------------
        
        'FGZ - 12/05/2014 -----------------------------------------------------------
        If (rs_Item!Itenro = 17) Or (rs_Item!Itenro = 31) Then
            If Ret_Ano > 2013 Then
                valor_ant = ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes)
                Items_TOPE(rs_Item!Itenro) = ((valor_ant) * ((1 + Por_Deduccion_zona / 100)))
                If rs_Item!Itenro = 31 Then Items_TOPE(rs_Item!Itenro) = IIf(Items_DDJJ(31) > Items_TOPE(rs_Item!Itenro), Items_TOPE(rs_Item!Itenro), Items_DDJJ(31))
            Else
                If (rs_Item!Itenro = 31 And Items_DDJJ(31) <> 0) Or (rs_Item!Itenro = 17) Then
                    StrSql = "SELECT itenro,vimonto FROM valitem WHERE viano =" & Ret_Ano & _
                            " AND itenro=" & rs_Item!Itenro & " AND vimes =" & Ret_Mes - 1
                    OpenRecordset StrSql, rs_valitem
                    If Not rs_valitem.EOF Then
                        valor_ant = rs_valitem!vimonto
                        'Si el mes es mayor a 9, busco el valor acutal.
                        If Ret_Ano = 2013 And Ret_Mes > 9 Then
                            If rs_Item!Itenro = 31 Then
                                Items_TOPE(rs_Item!Itenro) = ((((ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes) - valor_ant) * (Por_Deduccion_zona / 100)) * (Ret_Mes - 8)) + ValorEscala(rs_Item!Itenro, Ret_Ano, Ret_Mes))
                                Items_TOPE(rs_Item!Itenro) = IIf(Items_DDJJ(31) > Items_TOPE(rs_Item!Itenro), Items_TOPE(rs_Item!Itenro), Items_DDJJ(31))
                            Else
                                Items_TOPE(rs_Item!Itenro) = ((((Items_TOPE(rs_Item!Itenro) - valor_ant) * (Por_Deduccion_zona / 100)) * (Ret_Mes - 8)) + Items_TOPE(rs_Item!Itenro))
                            End If
                        Else
                            Items_TOPE(rs_Item!Itenro) = Items_TOPE(rs_Item!Itenro) - (Items_TOPE(rs_Item!Itenro) - valor_ant) + ((Items_TOPE(rs_Item!Itenro) - valor_ant) * (1 + (Por_Deduccion_zona / 100)))
                        End If
                    End If
                End If
            End If
        End If
        'FGZ - 12/05/2014 -----------------------------------------------------------
        
        'FGZ - 12/05/2015 -------------------------------
        Por_Deduccion_zona = Aux_Por_Deduccion_zona
        'FGZ - 12/05/2015 -------------------------------
    End If

    
    'FGZ - 24/07/2014 ---------------------------------------------------------------------
    If HACE_TRAZA Then
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-DDJJ"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_DDJJ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Liq"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-LiqAnt"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_OLD_LIQ(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Prorr"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-ProrrCuota"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_PRORR_CUOTA(rs_Item!Itenro))
        Texto = Format(CStr(rs_Item!Itenro), "00") & "-" & rs_Item!itenom & "-Tope"
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(rs_Item!Itenro))
    End If
    'Calcula la Ganancia Imponible
    If CBool(rs_Item!itesigno) Then
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

SiguienteItem:
    rs_Item.MoveNext
Loop
  
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Bruta: " & (Gan_Imponible - Descuentos + Items_TOPE(50))
        Flog.writeline Espacios(Tabulador * 3) & "9- Gan. Bruta - CMA y DONA.: " & Gan_Imponible
        Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Neta: " & (Gan_Imponible + Deducciones)
        Flog.writeline Espacios(Tabulador * 3) & "9- Total Deducciones: " & Deducciones
        Flog.writeline Espacios(Tabulador * 3) & "9- Total art. 23: " & Ded_a23
    End If
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Bruta ", Gan_Imponible - Descuentos + Items_TOPE(100))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Gan. Bruta - CMA y DONA.", Gan_Imponible)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Neta ", (Gan_Imponible + Deducciones))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Para Machinea ", (Gan_Imponible + Deducciones - Items_TOPE(100)))
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total Deducciones", Deducciones)
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Total art. 23", Ded_a23)
    End If
            
    
    ' Calculo el porcentaje de deduccion segun la ganancia neta
    'Uso el campo para guardar la ganancia neta para el 648
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganneta =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
  
    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Ret_Ano >= 2000 And Gan_Imponible > 0 Then
        StrSql = "SELECT esd_porcentaje FROM escala_ded " & _
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "No hay esc. dedu para", Gan_Imponible)
            End If
            ' No se ha encontrado la escala de deduccion para el valor gan_imponible
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- % a tomar deduc." & Por_Deduccion
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- % a tomar deduc.", Por_Deduccion)
        End If
        
        'Aplico el porcentaje a las deducciones
        Ded_a23 = Ded_a23 * Por_Deduccion / 100
        
'        'Guardo el porcentaje de deduccion
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " porcdeduc =" & Por_Deduccion
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    End If
    ' Calculo la Ganancia imponible
    Gan_Imponible = Gan_Imponible + Deducciones + Ded_a23

    'Menos de 15000 no paga-----------------------------------------------------------------------
    Ret_Aux = 0
     If Acum_Maximo <= 15000 Then
        If Ret_Aux <> 0 Then
            'Recalculo
            Gan_Imponible_Agosto = 0
            StrSql = "SELECT ganimpo FROM sim_traza_gan WHERE "
            'StrSql = StrSql & " concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
            'StrSql = StrSql & " AND fecha_pago >=" & ConvFecha(AuxInicio)
            StrSql = StrSql & " fecha_pago >=" & ConvFecha(AuxInicio)
            StrSql = StrSql & " AND fecha_pago <=" & ConvFecha(AuxFin)
            StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " ORDER BY fecha_pago desc "
            OpenRecordset StrSql, rs_Traza_gan
            If Not rs_Traza_gan.EOF Then
                Gan_Imponible_Agosto = rs_Traza_gan!Ganimpo
            End If
           
            AuxDedEspecial = Gan_Imponible - Gan_Imponible_Agosto
            Gan_Imponible = Gan_Imponible - AuxDedEspecial
            
            If CBool(USA_DEBUG) Then
                'Flog.writeline Espacios(Tabulador * 3) & "Hay retenciones anteriores a Septiembre 2013 " & Ret_Aux
                Flog.writeline Espacios(Tabulador * 3) & "Ganancia Imponible a  Agosto 2013 = " & Gan_Imponible_Agosto
                Flog.writeline Espacios(Tabulador * 3) & "Ajuste de Deduccion especial = " & AuxDedEspecial
                Flog.writeline Espacios(Tabulador * 3) & "Nueva Ganancia Imponible = " & Gan_Imponible
            End If
            
            If AuxDedEspecial > 0 Then
                Items_TOPE(16) = Items_TOPE(16) + AuxDedEspecial
                If HACE_TRAZA Then
                    Texto = Format(CStr(16), "00") & "-Deducción Especial-Tope"
                    StrSql = "DELETE sim_traza WHERE cliqnro = " & buliq_cabliq!cliqnro
                    StrSql = StrSql & " AND concnro = " & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND tpanro = 0"
                    StrSql = StrSql & " AND tradesc ='" & Texto & "'"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(16))
                End If
             End If
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Tope Deduccion especial = " & Items_TOPE(16)
            End If
        Else    'Igual
            Gan_Imponible = 0
            If CBool(USA_DEBUG) Then
                'Flog.writeline Espacios(Tabulador * 3) & "La Retencion es Cero porque el sueldo bruto es <= 15000"
            End If
         
            'EAM(5.44)- Inserta un concepto en detliq con valor 0. Sirve para mostrar leyenda
            'Insertar un concepto con valor 0.
            'El concepto debe ser configurable como un parametro mas de la formula. (si el concepto no existe ==> no insertar porque va a dar error)
            'El concepto se debe llamar Remuneración y/o Haber no sujeto al Impuesto a las Ganancias- Beneficio Decreto PEN 1242/2013
            StrSql = "SELECT concnro FROM concepto WHERE conccod= " & Leyenda_Concepto
            OpenRecordset StrSql, rs_Aux
            
            If rs_Aux.EOF Then
                Flog.writeline Espacios(Tabulador * 3) & "No se encuentra configurado el parametro 51 para insertar concepto 0"
            Else
                StrSql = "INSERT INTO sim_detliq (cliqnro,concnro,dlimonto,dlicant,ajustado,dlitexto,dliretro )" & _
                              " VALUES (" & buliq_cabliq!cliqnro & "," & rs_Aux!ConcNro & "," & 0 & "," & 0 & "," & -1 & ",''," & -1 & ")"
                     objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            'FGZ - 20/11/2013 ---------------------------------
            'Aquellos que perciban menos de $15.000.- de Enero a Agosto, la Deducción Especial (ITEM 16) se calcule de acuerdo a lo establecido en el Dec.1242/13 Art.1,
            'que dice lo siguiente:
            '
            'Artículo 1° - lncreméntase, respecto de las rentas mencionadas en los incisos a), b) y c) del artículo 79 de la Ley de Impuesto a las Ganancias,
            '   texto ordenado en 1997, y sus modificaciones, la deducción especial establecida en el inciso c) del artículo 23 de dicha Ley,
            '   hasta un monto equivalente al que surja de restar a la ganancia neta sujeta a impuesto las deducciones de los incisos a) y b) del mencionado artículo 23.
            
            'CALCULO
            '   ITEM01 + ITEM02 + ITEM03 - (resto de los items)
            '
            '   Resto de los items =
            '       ABS ITEM05 + ABS ITEM06 + ABS ITEM07 + ABS ITEM08 + ABS ITEM09 + ABS ITEM10 + ABS ITEM11 + ABS ITEM12 + ABS ITEM13
            '       + ABS ITEM15 + ABS ITEM16 + ABS ITEM17 + ABS ITEM20 + ABS ITEM23 + ABS ITEM24 + ABS ITEM31
    
            'Si este cálculo es menor a cero, este monto no debería sumarse al ítem 16, si es mayor a cero sí.
            AuxDedEspecial = Items_TOPE(1) + Items_TOPE(2) + Items_TOPE(3)
            AuxDedEspecial = AuxDedEspecial - Abs(Items_TOPE(5)) - Abs(Items_TOPE(6)) - Abs(Items_TOPE(7)) - Abs(Items_TOPE(8)) - Abs(Items_TOPE(9)) - Abs(Items_TOPE(10))
            AuxDedEspecial = AuxDedEspecial - Abs(Items_TOPE(11)) - Abs(Items_TOPE(12)) - Abs(Items_TOPE(13)) - Abs(Items_TOPE(15)) - Abs(Items_TOPE(16)) - Abs(Items_TOPE(17))
            AuxDedEspecial = AuxDedEspecial - Abs(Items_TOPE(20)) - Abs(Items_TOPE(23)) - Abs(Items_TOPE(24)) - Abs(Items_TOPE(31))
            
            If AuxDedEspecial > 0 Then
                Items_TOPE(16) = Items_TOPE(16) + AuxDedEspecial
                If HACE_TRAZA Then
                    Texto = Format(CStr(16), "00") & "-Deducción Especial-Tope"
            
                    'FGZ - 01/10/2013 - Borro antes de insertar nuevamente ---------------------
                    StrSql = "DELETE sim_traza WHERE cliqnro = " & buliq_cabliq!cliqnro
                    StrSql = StrSql & " AND concnro = " & Buliq_Concepto(Concepto_Actual).ConcNro
                    StrSql = StrSql & " AND tpanro = 0"
                    StrSql = StrSql & " AND tradesc ='" & Texto & "'"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    'FGZ - 01/10/2013 - Borro antes de insertar nuevamente ---------------------
                    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(16))
                End If
             End If
            'FGZ - 20/11/2013 ---------------------------------
        End If
     End If


    ' FGZ - 19/04/2004
    'Uso el campo para guardar la ganancia imponible para el 649
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " ganimpo =" & Gan_Imponible
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    ' FGZ - 19/04/2004
    
    'FGZ - 22/12/2004
    'Leo la tabla
    'FGZ - 08/06/2012 ------------------
    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
    OpenRecordset StrSql, rs_Traza_gan
    
    
    'FGZ - 15/04/2014 ------------------------------------------------------------------------------------------------------------------
    'If CBool(USA_DEBUG) Then
    '    Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible" & Gan_Imponible
    'End If
    'If HACE_TRAZA Then
    '    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Gan_Imponible)
    'End If
    'FGZ - 15/04/2014 ------------------------------------------------------------------------------------------------------------------
            
    'FGZ - 05/01/2015 ------------------------------------------------------------------------------------------------------------------
    'FGZ - 30/12/2014 ------------------------------------------------------------------------------------------------------------------
    If Acum_Maximo <= 15000 Then
        If Gan_Imponible > 0 Then
            'FGZ - 08/06/2012 ------------------
            'Entrar en la escala con las ganancias acumuladas
            
            'FGZ - 25/02/2014 ---------------------------------------------------
            'StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
            '         " WHERE escmes =" & Ret_Mes & _
            '         " AND escano =" & Ret_Ano & _
            '         " AND escinf <= " & Gan_Imponible & _
            '         " AND escsup >= " & Gan_Imponible
            
            StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
                     " WHERE escmes =" & IIf(Ret_Aux <> 0, 8, Ret_Mes) & _
                     " AND escano =" & Ret_Ano & _
                     " AND escinf <= " & Gan_Imponible & _
                     " AND escsup >= " & Gan_Imponible
            OpenRecordset StrSql, rs_escala
            'FGZ - 25/02/2014 ---------------------------------------------------
            If Not rs_escala.EOF Then
                Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
            Else
                Impuesto_Escala = 0
            End If
        Else
            Impuesto_Escala = 0
        End If
                        
        'FGZ - 15/04/2014 ------------------------------------------------------------------------------------------------------------------
        Gan_Imponible_Grosada = 0
        Aux_Gan_Imponible_Grosada = 0
        'Entro ahora en escala de diciembre
        'FGZ - 23/05/2014 -----------------------------------------------------
        'StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
        '         " WHERE escmes =" & IIf(Ret_Aux <> 0, 12, Ret_Mes) & _
        '         " AND escano =" & Ret_Ano & _
        '         " AND escinf <= " & Gan_Imponible & _
        '         " AND escsup >= " & Gan_Imponible
        StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
                 " WHERE escmes =" & IIf(Ret_Aux <> 0, 12, Ret_Mes) & _
                 " AND escano =" & Ret_Ano & _
                 " AND esccuota <= " & Impuesto_Escala & _
                 " ORDER BY esccuota DESC"
        'FGZ - 23/05/2014 -----------------------------------------------------
        OpenRecordset StrSql, rs_escala
        'FGZ - 25/02/2014 ---------------------------------------------------
        If Not rs_escala.EOF Then
            'Impuesto_Escala = rs_escala!esccuota + ((Gan_Imponible - rs_escala!escinf) * rs_escala!escporexe / 100)
            Aux_Gan_Imponible_Grosada = ((((Impuesto_Escala - rs_escala!esccuota) * 100 / rs_escala!escporexe)) + rs_escala!escinf)
            Gan_Imponible_Grosada = ((((Impuesto_Escala - rs_escala!esccuota) * 100 / rs_escala!escporexe)) + rs_escala!escinf) - Gan_Imponible
        Else
            Gan_Imponible_Grosada = 0
        End If
            
       
        If AuxDedEspecial > 0 Then
            'AuxDedEspecial = AuxDedEspecial - Gan_Imponible_Grosada
            Items_TOPE(16) = Items_TOPE(16) - Gan_Imponible_Grosada
            AuxDedEspecial = Items_TOPE(16)
        Else
            'FGZ - 20/02/2015 ---------------------------------
            valor_ant = ValorEscala(16, Ret_Ano, Ret_Mes)
            Items_TOPE(16) = valor_ant
            Items_TOPE(16) = Items_TOPE(16) + Abs(Items_LIQ(16))
            Aux_Gan_Imponible_Grosada = AuxDedEspecial - Abs(Items_LIQ(16))
            'FGZ - 20/02/2015 ---------------------------------
        End If
    
        If HACE_TRAZA Then
            Texto = Format(CStr(16), "00") & "-Deducción Especial-Tope"
            StrSql = "DELETE sim_traza WHERE cliqnro = " & buliq_cabliq!cliqnro
            StrSql = StrSql & " AND concnro = " & Buliq_Concepto(Concepto_Actual).ConcNro
            StrSql = StrSql & " AND tpanro = 0"
            StrSql = StrSql & " AND tradesc ='" & Texto & "'"
            objConn.Execute StrSql, , adExecuteNoRecords
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, Texto, Items_TOPE(16))
        
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible " & Aux_Gan_Imponible_Grosada
            End If
            If HACE_TRAZA Then
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Aux_Gan_Imponible_Grosada)
            End If
        End If
        'FGZ - 15/04/2014 ------------------------------------------------------------------------------------------------------------------
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 3) & "9- Ganancia Imponible " & Gan_Imponible
        End If
        If HACE_TRAZA Then
            Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "9- Ganancia Imponible", Gan_Imponible)
        End If
        If Gan_Imponible > 0 Then
            StrSql = "SELECT esccuota,escinf,escporexe FROM escala " & _
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
    End If
    'FGZ - 30/12/2014 ------------------------------------------------------------------------------------------------------------------
    'FGZ - 05/01/2015 ------------------------------------------------------------------------------------------------------------------
    
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
    
    StrSql = "UPDATE sim_traza_gan SET "
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
    If AuxDedEspecial > 0 Then
        StrSql = StrSql & ", dedesp =" & AuxDedEspecial
    Else
        StrSql = StrSql & ", dedesp =" & (Items_TOPE(16))
    End If
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
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
    
                
    ' Calculo las retenciones ya realizadas
    Ret_Ant = 0
        
    'Armo Fecha hasta como el ultimo dia del mes
    If (Ret_Mes = 12) Then
        fechaFichaH = CDate("31/12/" & Ret_Ano)
    Else
        fechaFichaH = CDate("01/" & Ret_Mes + 1 & "/" & Ret_Ano) - 1
    End If
    
    fechaFichaD = CDate("01/01/" & Ret_Ano)
    
    StrSql = "SELECT SUM(importe) monto FROM sim_ficharet " & _
             " WHERE empleado =" & buliq_empleado!Ternro & _
             " AND fecha <= " & ConvFecha(fechaFichaH) & _
             " AND fecha >= " & ConvFecha(fechaFichaD)
    OpenRecordset StrSql, rs_Ficharet
    If Not rs_Ficharet.EOF Then
        If Not IsNull(rs_Ficharet!Monto) Then
            Ret_Ant = Ret_Ant + rs_Ficharet!Monto
        End If
    End If
    
    'FGZ - 17/10/2013 ---------------------------------------------
    'Calcular la retencion
    Retencion = Impuesto_Escala - Ret_Ant
    
    'Calculo de Impuesto y Debitos Bancarios, solo aplica si el impuesto retiene, si devuelve para el otro año lo declarado para este item
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc", val_impdebitos)
    If val_impdebitos > Impuesto_Escala Then
        val_impdebitos = Impuesto_Escala
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Imp Debitos Banc con Tope", val_impdebitos)
    End If
    
    Retencion = Retencion - val_impdebitos
    'FGZ - 17/10/2013 ---------------------------------------------
    
    
    ' Para el F649 va en el 9b
    StrSql = "UPDATE sim_traza_gan SET "
    StrSql = StrSql & " promo =" & val_impdebitos
    StrSql = StrSql & " WHERE "
    StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    objConn.Execute StrSql, , adExecuteNoRecords
            
    
    'FGZ - 30/12/2004
    'Determinar el saldo
    'FGZ - 08/06/2012 ------------------
    StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
    StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
    StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
    StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
    StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
    'StrSql = StrSql & " AND empresa =" & NroEmp
    StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
    OpenRecordset StrSql, rs_Traza_gan
    
    If Not rs_Traza_gan.EOF Then
        StrSql = "UPDATE sim_traza_gan SET "
        StrSql = StrSql & "  saldo =" & Retencion
        StrSql = StrSql & "  ,retenciones =" & Ret_Ant
        StrSql = StrSql & " WHERE "
        StrSql = StrSql & " pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'FGZ - 22/12/2004
        'Leo la tabla
        'FGZ - 08/06/2012 ------------------
        StrSql = "SELECT cuit_entidad11,cuit_entidad12,cuit_entidad13,cuit_entidad14 FROM sim_traza_gan WHERE "
        StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
        StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
        StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
        'StrSql = StrSql & " AND empresa =" & NroEmp
        StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
        'If rs_Traza_gan.State = adStateOpen Then rs_Traza_gan.Close
        OpenRecordset StrSql, rs_Traza_gan
    End If
    ' FGZ - 19/04/2004
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 3) & "Retenciones anteriores " & Ret_Ant
        If Gan_Imponible > 0 Then
            'FGZ - 05/01/2015 --------------------
            If rs_escala.State = adStateOpen Then
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
            'FGZ - 05/01/2015 --------------------
        End If
    End If
    
    If HACE_TRAZA Then
        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Retenciones anteriores", Ret_Ant)
        If Gan_Imponible > 0 Then
            'FGZ - 05/01/2015 --------------------
            If rs_escala.State = adStateOpen Then
                If Not rs_escala.EOF Then
                    rs_escala.MoveFirst
                    If Not rs_escala.EOF Then
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", rs_escala!escporexe)
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
                    Else
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Escala Impuesto", 0)
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "Impuesto por escala", Impuesto_Escala)
                        Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver", Retencion)
                    End If
                End If
            End If
            'FGZ - 05/01/2015 --------------------
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
                Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).ConcNro, 0, "A Retener/Devolver, x Tope General", Retencion)
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
        Call InsertarFichaRet(buliq_empleado!Ternro, buliq_proceso!profecpago, Retencion, buliq_proceso!pronro)
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
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
            StrSql = StrSql & " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope
            
            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,monto,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_TOPE(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET monto =" & Items_TOPE(I)
                StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
                StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                StrSql = StrSql & " AND empresa =" & NroEmp
                StrSql = StrSql & " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 23/04/2004
        
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
        
        If Items_PRORR(I) <> 0 Then
           'Busco las liquidaciones anteriores
            StrSql = "INSERT INTO sim_desliq (empleado,DLfecha,pronro,DLmonto,DLprorratea,itenro) VALUES (" & _
                     buliq_empleado!Ternro & "," & _
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
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top "
            StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            'StrSql = StrSql & " AND empresa =" & NroEmp
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
                'StrSql = "UPDATE sim_traza_gan_item_top SET ddjj =" & Items_DDJJ(I)
                'StrSql = StrSql & " WHERE ternro =" & buliq_empleado!Ternro
                'StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                'StrSql = StrSql & " AND empresa =" & NroEmp
                'StrSql = StrSql & " AND itenro =" & I
                
                StrSql = "UPDATE sim_traza_gan_item_top SET ddjj =" & Items_DDJJ(I) & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND empresa =" & NroEmp & _
                " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_prorr para poder usarlo en el reporte de Ganancias
        If Items_PRORR_CUOTA(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top " & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,prorr,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_PRORR_CUOTA(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET prorr =" & Items_PRORR_CUOTA(I) & _
                    " WHERE ternro =" & buliq_empleado!Ternro & _
                    " AND pronro =" & buliq_proceso!pronro & _
                    " AND empresa =" & NroEmp & _
                    " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_old_liq para poder usarlo en el reporte de Ganancias
        If Items_OLD_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top " & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND itenro =" & I
            OpenRecordset StrSql, rs_Traza_gan_items_tope

            If rs_Traza_gan_items_tope.EOF Then
                StrSql = "INSERT INTO sim_traza_gan_item_top (ternro,pronro,old_liq,empresa,itenro) VALUES (" & _
                         buliq_empleado!Ternro & "," & _
                         buliq_proceso!pronro & "," & _
                         Items_OLD_LIQ(I) & "," & _
                         NroEmp & "," & _
                         I & _
                         ")"
            Else 'Actualizo
                StrSql = "UPDATE sim_traza_gan_item_top SET old_liq =" & Items_OLD_LIQ(I) & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND empresa =" & NroEmp & _
                " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If

        ' guardo los item_liq para poder usarlo en el reporte de Ganancias
        If Items_LIQ(I) <> 0 Then
            'inserto en traza_ga_Items_tope
            'si ya está actualizo y sino inserto
            'FGZ - 08/06/2012 ------------------
            StrSql = "SELECT itenro FROM sim_traza_gan_item_top " & _
                " WHERE ternro =" & buliq_empleado!Ternro & _
                " AND pronro =" & buliq_proceso!pronro & _
                " AND itenro =" & I
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
                StrSql = "UPDATE sim_traza_gan_item_top SET liq =" & Items_LIQ(I) & _
                    " WHERE ternro =" & buliq_empleado!Ternro & _
                    " AND pronro =" & buliq_proceso!pronro & _
                    " AND empresa =" & NroEmp & _
                    " AND itenro =" & I
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
        'FGZ 31/03/2005
        
        I = I + 1
    Loop

    exito = Bien
    for_Ganancias2013_OLD = Monto
    
End Function


Public Function ValorConcepto(ByVal Conccod As String, ByVal Anio As Integer, ByVal Mes As Integer, ByVal Cant As Boolean) As Double
Dim rs_Periodos As New ADODB.Recordset
Dim aux As Double

aux = 0

    'Busco todos lod detliq entre los meses
    StrSql = "SELECT sum(sim_detliq.dlicant) cantidad, sum(sim_detliq.dlimonto) monto  FROM periodo "
    StrSql = StrSql & " INNER JOIN sim_proceso ON periodo.pliqnro = sim_proceso.pliqnro "
    StrSql = StrSql & " INNER JOIN sim_cabliq ON sim_proceso.pronro = sim_cabliq.pronro AND sim_cabliq.empleado = " & NroEmple
    StrSql = StrSql & " INNER JOIN sim_detliq ON sim_cabliq.cliqnro = sim_detliq.cliqnro "
    StrSql = StrSql & " INNER JOIN concepto ON sim_detliq.concnro = concepto.concnro AND concepto.conccod = '" & Conccod & "'"
    StrSql = StrSql & " WHERE periodo.pliqmes = " & Mes & " AND periodo.pliqanio = " & Anio
    OpenRecordset StrSql, rs_Periodos

    If Cant Then
        aux = IIf(EsNulo(rs_Periodos!Cantidad), 0, rs_Periodos!Cantidad)
    Else
        aux = IIf(EsNulo(rs_Periodos!Monto), 0, rs_Periodos!Monto)
    End If

    ValorConcepto = aux
End Function


Public Function BuscarBrutoAgosto2013(ByVal BrutoMensual As Long, ByRef MantenerIncremento As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Autor      : FGZ
' Fecha      : 03/10/2013
' Ultima Mod.: FGZ - 13/05/2015
' Descripcion: Se agregan un par de modificaciones / controles en la funcion que busca el bruto (RG3770)
' 1) Si el monto máximo del acumulador definido en el parámetro 75 (Bruto) de la fórmula de ganancias, entre Enero y Agosto de 2013, superó $25.000.- (3ra. Franja),
'    deberá compararse con el valor máximo del mismo acumulador desde Enero de 2015.
'    Si éste no supera $25.000.-, deberá considerarse la reubicación del empleado, en la 1ra. Franja,
'    si no supera $15.000.- o en alguna de las subdivisiones de la 2da. Franja, si ese valor se encuentra entre $15.000,01 y $25.000.-
'
'2) 2)  Para aquellos empleados cuyo inicio de actividades fue a partir del 1/09/2013, la forma de determinar en que franja se ubica,
'       será considerando la mayor remuneración percibida a partir del 01/01/2015 y no la inicial solamente
'
' Ultima Mod.: FGZ - 21/05/2015
' Descripcion: Se agreguna variable para indicar si se debe mantener los % anteriores
' ---------------------------------------------------------------------------------------------
Dim rs_itemacum                 As New ADODB.Recordset
Dim Acum_Maximo                 As Double
Dim Acum_Maximo_HastaAgosto     As Double
Dim Acum_Maximo_2015            As Double
Dim IngresoDespuesAgosto2013    As Boolean
Dim PosteriorMayo2015           As Boolean

    MantenerIncremento = False
    PosteriorMayo2015 = False
    
    'Busco el mayor valor del acum bruto de enero a Julio inclusive
    StrSql = "SELECT max (ammonto) monto FROM sim_acu_mes " & _
            " WHERE acunro =" & BrutoMensual & _
            " AND ternro =  " & buliq_empleado!Ternro & _
            " AND amanio =  2013 and ammes >= 1 and ammes <= 8"
    OpenRecordset StrSql, rs_itemacum
    If Not IsNull(rs_itemacum!Monto) Then
        Acum_Maximo = rs_itemacum!Monto
        
        'FGZ - 13/05/2015 ----------------------
        Acum_Maximo_HastaAgosto = Acum_Maximo
        IngresoDespuesAgosto2013 = False
        MantenerIncremento = False
        
        If Acum_Maximo_HastaAgosto > 25000 Then
            StrSql = "SELECT ammonto monto FROM sim_acu_mes " & _
                    " WHERE acunro =" & BrutoMensual & _
                    " AND ternro =  " & buliq_empleado!Ternro & _
                    " AND amanio >=  2015 " & _
                    " ORDER BY ammonto desc, amanio, ammes"
            OpenRecordset StrSql, rs_itemacum
            If Not rs_itemacum.EOF Then
                If Not IsNull(rs_itemacum!Monto) Then
                    Acum_Maximo_2015 = rs_itemacum!Monto
                End If
            End If
        
            'FGZ - 18/05/2015 -------------------------------------------------------------------
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(BrutoMensual)) Then
                If objCache_Acu_Liq_Monto.Valor(CStr(BrutoMensual)) > Acum_Maximo_2015 Then
                    Acum_Maximo_2015 = objCache_Acu_Liq_Monto.Valor(CStr(BrutoMensual))
                End If
            End If
            
            'FGZ - 19/05/2015 --------------------------------------------------------------------
            If Acum_Maximo_2015 < Acum_Maximo Then
                Acum_Maximo = Acum_Maximo_2015
                MantenerIncremento = False
            Else
                MantenerIncremento = True
            End If
            'FGZ - 19/05/2015 --------------------------------------------------------------------
        End If
        'FGZ - 13/05/2015 ----------------------
    Else
        IngresoDespuesAgosto2013 = True
        'Busco el primer valor del acum bruto de Agosto a Diciembre inclusive
        StrSql = "SELECT ammonto monto FROM sim_acu_mes "
        StrSql = StrSql & " WHERE acunro =" & BrutoMensual
        StrSql = StrSql & " AND ternro =  " & buliq_empleado!Ternro
        StrSql = StrSql & " AND amanio =  2013 "
        StrSql = StrSql & "  AND ( ammes >= 8 AND ammes <= 12) "
        StrSql = StrSql & "  ORDER BY ammes "
        OpenRecordset StrSql, rs_itemacum
        If Not rs_itemacum.EOF Then
            If Not IsNull(rs_itemacum!Monto) Then
                Acum_Maximo = rs_itemacum!Monto
            Else
                Acum_Maximo = 0
            End If
        Else
            Acum_Maximo = 0
        End If
    End If

    'FGZ - 12/05/2014 ---------------------------------------------------
    'La normativa indica que si el empleado no ha tenido un bruto habitual para ganancias desde enero a agosto de 2013 el bruto habitual que lo segmentará será el primero que cobre.
    'Por lo que si en enero no se le retuvo nada, este comportamiento debe continuar así hasta tanto este cambio legal no se modifique.
    If Acum_Maximo = 0 Then
        StrSql = "SELECT sim_acu_mes.ammonto monto FROM sim_acu_mes " & _
                " WHERE acunro =" & BrutoMensual & _
                " AND ternro =  " & buliq_empleado!Ternro & _
                " AND amanio >=  2014 " & _
                " ORDER BY amanio, ammes"
        OpenRecordset StrSql, rs_itemacum
        If Not rs_itemacum.EOF Then
            If Not IsNull(rs_itemacum!Monto) Then
                Acum_Maximo = rs_itemacum!Monto
            End If
        End If
    End If
    'FGZ - 12/05/2014 ---------------------------------------------------

    'FGZ - 21/05/2015 --------------------------------------------------------
    ''Si no tiene acum porque ingreso este mes debe tomar el sueldo actual
    'If Acum_Maximo = 0 Then
    '    'busco los acu_liq del periodo actual del acumulador brutomensual
    '    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(BrutoMensual)) Then
    '        Acum_Maximo = objCache_Acu_Liq_Monto.Valor(CStr(BrutoMensual))
    '    End If
    'End If
    'FGZ - 21/05/2015 --------------------------------------------------------
    
    'FGZ - 13/05/2015 -------------------------------------
    If IngresoDespuesAgosto2013 Then
        If Acum_Maximo <= 0 Or Acum_Maximo > 15000 Then
            StrSql = "SELECT ammonto monto, ammes, amanio FROM sim_acu_mes " & _
                    " WHERE acunro =" & BrutoMensual & _
                    " AND ternro =  " & buliq_empleado!Ternro & _
                    " AND amanio >=  2015 " & _
                    " ORDER BY ammonto desc, amanio, ammes"
            OpenRecordset StrSql, rs_itemacum
            If Not rs_itemacum.EOF Then
                If Not IsNull(rs_itemacum!Monto) Then
                    Acum_Maximo_2015 = rs_itemacum!Monto
                    If Acum_Maximo = 0 Then
                        If rs_itemacum!amanio > 2015 Then
                            PosteriorMayo2015 = True
                        Else
                            If rs_itemacum!ammes >= 5 Then
                                PosteriorMayo2015 = True
                            Else
                                PosteriorMayo2015 = False
                            End If
                        End If
                    Else
                        PosteriorMayo2015 = False
                    End If
                End If
            End If
        
            'FGZ - 18/05/2015 -------------------------------------------------------------------
            If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(BrutoMensual)) Then
                If Acum_Maximo_2015 = 0 Then
                    PosteriorMayo2015 = True
                    MantenerIncremento = True
                End If
                If objCache_Acu_Liq_Monto.Valor(CStr(BrutoMensual)) > Acum_Maximo_2015 Then
                    Acum_Maximo_2015 = objCache_Acu_Liq_Monto.Valor(CStr(BrutoMensual))
                End If
            End If
            ''Si no tiene acum porque ingreso este mes debe tomar el sueldo actual
            'If Acum_Maximo_2015 = 0 Then
            '    'busco los acu_liq del periodo actual del acumulador brutomensual
            '    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(BrutoMensual)) Then
            '        Acum_Maximo_2015 = objCache_Acu_Liq_Monto.Valor(CStr(BrutoMensual))
            '    End If
            'End If
            'FGZ - 18/05/2015 -------------------------------------------------------------------
            If Acum_Maximo_2015 >= 25000 And PosteriorMayo2015 Then
                Acum_Maximo = Acum_Maximo_2015
                MantenerIncremento = False
            Else
                If Acum_Maximo_2015 >= 25000 Then
                    MantenerIncremento = True
                    'Le considero el bruto anterior
                Else
                    If Acum_Maximo_2015 > Acum_Maximo Then
                        Acum_Maximo = Acum_Maximo_2015
                        MantenerIncremento = False
                    Else
                        If Acum_Maximo_2015 <= 15000 Then
                            Acum_Maximo = Acum_Maximo_2015
                            MantenerIncremento = True
                        Else
                            If Acum_Maximo_2015 < 25000 Then
                                Acum_Maximo = Acum_Maximo_2015
                                MantenerIncremento = False
                            Else
                                MantenerIncremento = False
                                'Le considero el bruto anterior
                            End If
                        End If
                    End If
                End If
            End If
        Else
            'Le dejo lo que tenia
            MantenerIncremento = True
        End If
    End If
    'FGZ - 13/05/2015 -------------------------------------

    BuscarBrutoAgosto2013 = Acum_Maximo

End Function


Public Function BuscarBrutoAgosto2013_old(ByVal BrutoMensual As Long) As Double
Dim rs_itemacum As New ADODB.Recordset
Dim Acum_Maximo As Double

    'FGZ - 03/10/2013 ------------------------------------------------------------------------------
    'Busco el mayor valor del acum bruto de enero a Julio inclusive
    StrSql = "SELECT max (sim_acu_mes.ammonto) monto FROM sim_acu_mes " & _
            " WHERE acunro =" & BrutoMensual & _
            " AND ternro =  " & buliq_empleado!Ternro & _
            " AND amanio =  2013 and ammes >= 1 and ammes <= 8"
    OpenRecordset StrSql, rs_itemacum
    If Not IsNull(rs_itemacum!Monto) Then
        Acum_Maximo = rs_itemacum!Monto
    Else
        'Busco el primer valor del acum bruto de Agosto a Diciembre inclusive
        StrSql = "SELECT sim_acu_mes.ammonto monto FROM sim_acu_mes "
        StrSql = StrSql & " WHERE acunro =" & BrutoMensual
        StrSql = StrSql & " AND ternro =  " & buliq_empleado!Ternro
        StrSql = StrSql & " AND amanio =  2013 "
        StrSql = StrSql & "  AND ( ammes >= 8 AND ammes <= 12) "
        StrSql = StrSql & "  ORDER BY ammes "
        OpenRecordset StrSql, rs_itemacum
        If Not rs_itemacum.EOF Then
            If Not IsNull(rs_itemacum!Monto) Then
                Acum_Maximo = rs_itemacum!Monto
            Else
                Acum_Maximo = 0
            End If
        Else
            Acum_Maximo = 0
        End If
    End If

    'FGZ - 12/05/2014 ---------------------------------------------------
    'La normativa indica que si el empleado no ha tenido un bruto habitual para ganancias desde enero a agosto de 2013 el bruto habitual que lo segmentará será el primero que cobre.
    'Por lo que si en enero no se le retuvo nada, este comportamiento debe continuar así hasta tanto este cambio legal no se modifique.
    If Acum_Maximo = 0 Then
        'FGZ - 29/05/2014
        'StrSql = "SELECT acu_mes.ammonto monto FROM sim_acu_mes "
        StrSql = "SELECT sim_acu_mes.ammonto monto FROM sim_acu_mes " & _
                " WHERE acunro =" & BrutoMensual & _
                " AND ternro =  " & buliq_empleado!Ternro & _
                " AND amanio >=  2014 " & _
                " ORDER BY amanio, ammes"
        OpenRecordset StrSql, rs_itemacum
        If Not rs_itemacum.EOF Then
            If Not IsNull(rs_itemacum!Monto) Then
                Acum_Maximo = rs_itemacum!Monto
            End If
        End If
    End If
    'FGZ - 12/05/2014 ---------------------------------------------------


    'Si no tiene acum porque ingreso este mes debe tomar el sueldo actual
    If Acum_Maximo = 0 Then
        'busco los acu_liq del periodo actual del acumulador brutomensual
        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(BrutoMensual)) Then
            Acum_Maximo = objCache_Acu_Liq_Monto.Valor(CStr(BrutoMensual))
        End If
    End If

    BuscarBrutoAgosto2013_old = Acum_Maximo

End Function



Public Function ValorEscala(ByVal Itenro, ByVal Ret_Ano As Long, ByVal Ret_Mes As Long) As Double
Dim rs_valitem As New ADODB.Recordset
Dim valor_ant As Double

    valor_ant = 0
    StrSql = "SELECT itenro,vimonto FROM valitem WHERE viano =" & Ret_Ano & _
            " AND itenro=" & Itenro & " AND vimes = " & Ret_Mes
    OpenRecordset StrSql, rs_valitem
    If Not rs_valitem.EOF Then
        valor_ant = rs_valitem!vimonto
    End If
    
    ValorEscala = valor_ant
End Function



Public Sub LimpiaTraza_Gan()

StrSql = "DELETE FROM sim_traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).ConcNro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
StrSql = StrSql & " AND ternro =" & buliq_empleado!Ternro
objConn.Execute StrSql, , adExecuteNoRecords

' Guardo la traza de Ganancia en traza_gan para utilizarla luego cuando se genere el reporte
StrSql = "INSERT INTO sim_traza_gan (pliqnro,pronro,concnro,fecha_pago,empresa,ternro,empleg) VALUES (" & _
         buliq_periodo!PliqNro & "," & _
         buliq_proceso!pronro & "," & _
         Buliq_Concepto(Concepto_Actual).ConcNro & "," & _
         ConvFecha(buliq_proceso!profecpago) & "," & _
         NroEmp & "," & _
         buliq_empleado!Ternro & "," & _
         buliq_empleado!Empleg & _
         ")"
objConn.Execute StrSql, , adExecuteNoRecords



End Sub



Public Function for_Comisiones() As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Formula para el calculo de comisiones de Raffo
' ---------------------------------------------------------------------------------------------
'La idea se resume como sigue:
'   1. Buscar todos los conceptos asociados a la escala configurada en el parametro
'   2. Por cada concepto
'      buscar la novedad individual del empleado asociada al concepto - parametro (el parametro es uno solo,
'           y el nro es uno de los parametros de la formula). Esta novedad tiene el % de venta alcanzado por el empleado
'      Se debe encontrar que estructura de tipo (Configurada como parametro de la formula) tiene el empleado
'           y entrar a escala junto con el porcentaje de venta para obtener el valor de comision
'   3. El resultado de la formula es la sumatoria de todas las comisiones encontradas en el punto 2
' ---------------------------------------------------------------------------------------------
'Tablas
        'CREATE TABLE [dbo].[escala_comision](
        '    [esccomnro] [int] IDENTITY(1,1) NOT NULL,
        '    [esccomdesabr] [varchar](100) NOT NULL,     -- Descripcion Abreviada
        '    [esccomdesext] [varchar](500) NULL,         -- Descripcion Extendida
        '    [version] [varchar](10) NULL,               -- Version
        '    [activa] [smallint] NOT NULL default(-1),   -- Activa (True / False)
        '    [fecdesde] [datetime] NOT NULL,             -- Fecha desde de vigencia
        '    [fechasta] [datetime] NULL                  -- Fecha hasta de vigencia
        ') ON [PRIMARY]
        'GO
        
        'productos(conceptos) asociados a la escala de comisiones
        'CREATE TABLE [dbo].[escala_comision_conc](
        '    [esccomnro] [int] NOT NULL,                         -- FK a tabla escala_comision
        '    [concnro] [int] NOT NULL                            -- Concepto (Producto). FK a tabla concepto
        ') ON [PRIMARY]
        'GO
        
        'Lineas de productos(Estructuras) asociados a la escala de comisiones
        'CREATE TABLE [dbo].[escala_comision_estr](
        '    [esccomnro] [int] NOT NULL,                         -- FK a tabla escala_comision
        '    [tenro] [int] NOT NULL,                             -- Tipo de estructura. FK a tabla tipo_estructura
        '    [estrnro] [int] NOT NULL                            -- Estructura. FK a tabla estructura
        ') ON [PRIMARY]
        'GO
        
        'CREATE TABLE [dbo].[escala_comision_det](
        '    [esccomnro] [int] NOT NULL,                         -- FK a tabla escala_comision
        '    [esccomdetnro] [int] IDENTITY(1,1) NOT NULL,        -- identidad del detalle
        '    [tenro] [int] NOT NULL,                             -- Tipo de estructura. FK a tabla tipo_estructura
        '    [estrnro] [int] NOT NULL,                           -- Estructura. FK a tabla estructura
        '    [concnro] [int] NOT NULL,                           -- Concepto (Producto). FK a tabla concepto
        '    [pordesde] [decimal](19, 4) NOT NULL,               -- Porcentaje desde. Rango
        '    [porhasta] [decimal](19, 4) NOT NULL,               -- Porcentaje Hasta. Rango
        '    [comision] [decimal](19, 4) NOT NULL,               -- Valor de Comision
        '    [comision2] [decimal](19, 4) NOT NULL default(0),   -- Valor de Comision 2
        '    [comision3] [decimal](19, 4) NOT NULL default(0),   -- Valor de Comision 3
        '    [comision4] [decimal](19, 4) NOT NULL default(0)    -- Valor de Comision 4
        ') ON [PRIMARY]
        'GO
' ---------------------------------------------------------------------------------------------
' Autor      : FGZ
' Fecha      : 04/04/2014
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_comision As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Productos As New ADODB.Recordset

Dim c_concepto As Long
Dim c_Escala As Long
Dim c_Tipo As Long
Dim c_parametro As Long

Dim v_Escala     As Long
Dim v_Tipo       As Long
Dim v_Par  As Long

Dim Linea       As Long
Dim venta       As Double

Dim Fecha_Desde_P  As Date
Dim Fecha_Hasta_P  As Date
Dim Fecha_Desde_E  As Date
Dim Fecha_Hasta_E  As Date
Dim Comision            As Double
Dim I                   As Long
Dim OK                  As Boolean
Dim EscalaValida    As Boolean

   
    'inicializacion de variables
    c_Escala = 51
    c_Tipo = 1025
    c_parametro = 52
    
    Bien = False
    exito = False
    EscalaValida = False
    Comision = 0
    
    '---------------------------------------------------------------
    For I = LI_WF_Tpa To LS_WF_Tpa
        Select Case Arr_WF_TPA(I).tipoparam
        Case c_Escala:
            v_Escala = Arr_WF_TPA(I).Valor
        Case c_Tipo:
            v_Tipo = Arr_WF_TPA(I).Valor
        Case c_parametro:
            v_Par = Arr_WF_TPA(I).Valor
        Case Else
        End Select
    Next I

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Escala: " & c_Escala
    End If
    '---------------------------------------------------------------
    
    
    'Establezco fechas de analisis
    Fecha_Desde_P = buliq_proceso!profecini
    Fecha_Hasta_P = buliq_proceso!profecfin
    
    Fecha_Desde_E = Empleado_Fecha_Inicio 'buliq_proceso!profecini
    Fecha_Hasta_E = Empleado_Fecha_Fin    'buliq_proceso!profecfin o fecha de baja del empleado

    'FGZ - 30/04/2014 --------------------------------------------------
    'Validacion de la escala
    StrSql = " SELECT esccomdesabr, version, fecdesde, fechasta FROM escala_comision "
    StrSql = StrSql & " WHERE esccomnro = " & v_Escala
    StrSql = StrSql & " AND activa = -1"
    StrSql = StrSql & " AND (fecdesde <= " & ConvFecha(Fecha_Desde_P) & ") AND "
    StrSql = StrSql & " ((" & ConvFecha(Fecha_Desde_P) & " <= fechasta) or (fechasta is null))"
    OpenRecordset StrSql, rs_comision
    If rs_comision.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 5) & "La Escala " & v_Escala & " no está activa y/o con vigencia a la fecha " & Fecha_Desde_P
        End If
    Else
        EscalaValida = True
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 5) & "Escala Nro    : " & v_Escala
            Flog.writeline Espacios(Tabulador * 5) & "Descripcion   : " & rs_comision!esccomdesabr
            Flog.writeline Espacios(Tabulador * 5) & "Versión       : " & rs_comision!Version
            Flog.writeline Espacios(Tabulador * 5) & "Vigencia Desde: " & rs_comision!fecdesde
            Flog.writeline Espacios(Tabulador * 5) & "Vigencia Hasta: " & IIf(EsNulo(rs_comision!fechasta), "#", rs_comision!fechasta)
        End If
    End If
    'FGZ - 30/04/2014 --------------------------------------------------
    If EscalaValida Then
        'Busco la estructura del empleado para el tipo pasado por parametro
        StrSql = " SELECT estrnro FROM sim_his_estructura " & _
                 " WHERE ternro = " & buliq_empleado!Ternro & " AND " & _
                 " tenro =" & v_Tipo & " AND " & _
                 " (htetdesde <= " & ConvFecha(Fecha_Hasta_E) & ") AND " & _
                 " ((" & ConvFecha(Fecha_Hasta_E) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            Linea = rs_Estructura!Estrnro
        Else
            Linea = 0
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "El empleado no tiene linea asociada "
            End If
        End If
    
    
        'Busco todos los conceptos asociados a la escala
        StrSql = " SELECT concnro FROM escala_comision_conc WHERE esccomnro = " & v_Escala
        OpenRecordset StrSql, rs_Productos
        If rs_Productos.EOF Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "La Escala no tiene Productos asociados"
            End If
        End If
        Do While Not rs_Productos.EOF
            'Buscar novedad asociada al producto (concepto)
            venta = 0
            Call Bus_NovGegi(3, rs_Productos!ConcNro, v_Par, Empleado_Fecha_Inicio, Empleado_Fecha_Fin, Linea, OK, venta)
            If Not OK Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontró volumne venta para el empleado"
                End If
            End If
            
            'Buscar en escala para Linea, producto, Porcentaje de venta
            StrSql = " SELECT comision, comision2, comision3, comision4 FROM escala_comision_det "
            StrSql = StrSql & " WHERE esccomnro = " & v_Escala
            StrSql = StrSql & " AND estrnro = " & Linea
            StrSql = StrSql & " AND concnro = " & rs_Productos!ConcNro
            StrSql = StrSql & " AND ( pordesde <= " & venta & " AND porhasta >= " & venta & ")"
            OpenRecordset StrSql, rs_comision
            Do While Not rs_comision.EOF
                Comision = Comision + rs_comision!Comision
                rs_comision.MoveNext
            Loop
        
            rs_Productos.MoveNext
        Loop
End If

Monto = Comision
for_Comisiones = Comision
Bien = True
exito = True
End Function


Public Function PorcentajeRG3770(ByVal Zona As Double, ByVal Valor As Double) As Double
' ---------------------------------------------------------------------------------------------
' Autor      : FGZ
' Fecha      : 12/05/2015
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Escala(1 To 2, 1 To 6) As Single
'Zona normal
Escala(1, 1) = 50
Escala(1, 2) = 44
Escala(1, 3) = 38
Escala(1, 4) = 32
Escala(1, 5) = 29
Escala(1, 6) = 26
'Zona desfavorable
'FGZ - 02/06/2015 --------------
'Escala(2, 1) = 63
'Escala(2, 2) = 56
'Escala(2, 3) = 50
'Escala(2, 4) = 43
'Escala(2, 5) = 40
'Escala(2, 6) = 37
Escala(2, 1) = 62.5
Escala(2, 2) = 56
Escala(2, 3) = 49.5
Escala(2, 4) = 43
Escala(2, 5) = 39.75
Escala(2, 6) = 36.5
'FGZ - 02/06/2015 --------------

Dim Z As Integer 'Zona
Dim R As Integer 'Rango
   
   
    'Determino Zona
    Select Case Zona
    Case 20:    'Zona Normal
        Z = 1
    Case 30:    'Zona Desfavorable
        Z = 2
    Case Else
        Z = 1
    End Select
   
    'Determino Rango
    If Valor > 15000 And Valor <= 18000 Then
        R = 1
    Else
        If Valor > 18000 And Valor <= 21000 Then
            R = 2
        Else
            If Valor > 21000 And Valor <= 22000 Then
                R = 3
            Else
                If Valor > 22000 And Valor <= 23000 Then
                    R = 4
                Else
                    If Valor > 23000 And Valor <= 24000 Then
                        R = 5
                    Else
                        If Valor > 24000 And Valor <= 25000 Then
                            R = 6
                        Else
                            R = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
   
    PorcentajeRG3770 = Escala(Z, R)
End Function

