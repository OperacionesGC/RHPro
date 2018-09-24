Attribute VB_Name = "MdlFormulasPy"
Option Explicit

' ---------------------------------------------------------
' Modulo de fórmulas conocidas para Paraguay
' ---------------------------------------------------------

Public Type T_HorasT
    THT As Variant
    MPT As Variant
End Type




Public Function for_Comision() As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Formula para el calculo de comisiones
' ---------------------------------------------------------------------------------------------
'La idea se resume como sigue:
'   1.  El cálculo de comisiones se podrá ejecutar por un periodo según conveniencia, donde las fechas del proceso indicarán el inicio y fin del período. Se deben definir procesos diferentes por línea, pero se puede realizar 1 solo proceso por mes x línea. En principio se definió un proceso por línea para determinar fácilmente el total de empleados por línea.
'   2.  Definir 1 Concepto por cada línea de servicio que contenga un parámetro definido como novedad global.
'   3.  Implementar una interfaz que permita cargar novedades globales con vigencia diaria. En este punto ellos deberían armar un archivo Excel con un formato similar al siguiente: Concepto, parámetro, monto, fecha. Luego eso se importa con el valor total diario para la línea.
'   4.  Implementar una fórmula interna en RH Pro que resuelva el monto total de la comisión. Internamente se guardarán los valores diarios del importe de comisión por hora, cantidad de horas de la persona para el día y valor total de comisión por el día. El cálculo se realiza a mes vencido.
'   Las variables a resolver son:
'   Para cada uno de los días del periodo:
'       a.  calcular el total de horas trabajadas por día por todos los empleados = THT
'       b.  Tomar el valor de la novedad global a pagar para ese dia = MPT
'       c.  Calcular las horas trabajadas para el dia por el empleado = HT
'       d.Hacer MPT / THT * HT
'       e.  Guardar valores de las variables a,b,c,d, concepto, proceso, fecha.
' ---------------------------------------------------------------------------------------------
' Autor      : FGZ
' Fecha      : 27/09/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_comision As New ADODB.Recordset

Dim c_concepto As Long
Dim c_parametro As Long
Dim c_tipohora As Long
Dim Encontro1 As Boolean
Dim Encontro2 As Boolean

Dim v_Hora_Produccion     As Long
Dim v_Monto     As Long

Dim ControlEstricto As Boolean
Dim Fecha_Desde_P  As Date
Dim Fecha_Hasta_P  As Date
Dim Fecha_Desde_E  As Date
Dim Fecha_Hasta_E  As Date
Dim Fecha_Aux  As Date
Dim Comision            As Double
Dim I                   As Long
Dim HorasT(1 To 31)     As T_HorasT
Dim HorasE(1 To 31)     As Double
Dim ComisionE(1 To 31)  As Double
    
    
    
    'inicializacion de variables
    c_concepto = Arr_conceptos(Concepto_Actual).ConcNro
    c_parametro = 51
    c_tipohora = 3
    
    
    Bien = False
    exito = False
    ControlEstricto = True  'Control estricto de vigencia
    Comision = 0
    
'    '---------------------------------------------------------------
'    Encontro1 = False
'    Encontro2 = False
'    For I = LI_WF_Tpa To LS_WF_Tpa
'        Select Case Arr_WF_TPA(I).tipoparam
'        Case c_parametro:
'            v_Monto = Arr_WF_TPA(I).Valor
'            Encontro1 = True
'        Case c_tipohora:
'            v_Hora_Produccion = Arr_WF_TPA(I).Valor
'            Encontro2 = True
'        Case Else
'        End Select
'    Next I
'
'    'si no se obtuvieron los parametros, ==> defaults.
'    If Not Encontro1 Then
'        v_Monto = 51 'Tpanro del monto
'        If CBool(USA_DEBUG) Then
'            Flog.writeline Espacios(Tabulador * 4) & "Parametro " & c_parametro & " no configurado, default " & v_Monto
'        End If
'    End If
'    If Not Encontro2 Then
'        v_Hora_Produccion = 12 'Tipo de Hora Trabajada
'        If CBool(USA_DEBUG) Then
'            Flog.writeline Espacios(Tabulador * 4) & "Parametro " & c_tipohora & " no configurado, default " & v_Hora_Produccion
'        End If
'    End If
    v_Monto = 51 'Tpanro del monto
    v_Hora_Produccion = 12 'Tipo de Hora Trabajada
'    '---------------------------------------------------------------
    
    
    'inicializo totales
    For I = 1 To 31
        HorasT(I).MPT = 0
        HorasT(I).THT = 0
    Next I
    'inicializo empleado
    For I = 1 To 31
        HorasE(I) = 0
        ComisionE(I) = 0
    Next I

    
    'Establezco fechas de analisis
    Fecha_Desde_P = buliq_proceso!profecini
    Fecha_Hasta_P = buliq_proceso!profecfin
    
    Fecha_Desde_E = Empleado_Fecha_Inicio 'buliq_proceso!profecini
    Fecha_Hasta_E = Empleado_Fecha_Fin    'buliq_proceso!profecfin o fecha de baja del empleado


    Fecha_Aux = Fecha_Desde_P
    Do While Fecha_Aux <= Fecha_Hasta_P
        I = Day(Fecha_Aux)
        HorasT(I).MPT = Nov_Global(ControlEstricto, c_concepto, c_parametro, Fecha_Aux)
        HorasT(I).THT = CalcularHoras(0, v_Hora_Produccion, Fecha_Aux)

        Fecha_Aux = DateAdd("d", 1, Fecha_Aux)
    Loop
    
    Fecha_Aux = Fecha_Desde_E
    Do While Fecha_Aux <= Fecha_Hasta_E
        I = Day(Fecha_Aux)
        HorasE(I) = CalcularHoras(buliq_empleado!Ternro, v_Hora_Produccion, Fecha_Aux)
        'Calculo de comision para el dia
        If HorasT(I).THT <> 0 Then
            ComisionE(I) = HorasT(I).MPT / HorasT(I).THT * HorasE(I)
            Comision = Comision + ComisionE(I)
        
            'Guardar valores de las variables a,b,c,d, concepto, proceso, fecha.
            'Inserto comision
            StrSql = "INSERT INTO sim_liq_comision ("
            StrSql = StrSql & "ternro,fecha,concnro,tpanro,thnro,mpt,tht,th,simpronro "
            StrSql = StrSql & ") VALUES (" & buliq_empleado!Ternro
            StrSql = StrSql & "," & ConvFecha(Fecha_Aux)
            StrSql = StrSql & "," & c_concepto
            StrSql = StrSql & "," & c_tipohora
            StrSql = StrSql & "," & v_Hora_Produccion
            StrSql = StrSql & "," & HorasT(I).MPT
            StrSql = StrSql & "," & HorasT(I).THT
            StrSql = StrSql & "," & HorasE(I)
            StrSql = StrSql & "," & buliq_proceso!pronro
            StrSql = StrSql & " )"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Horas Trabajadas totales = 0. No Hay comision "
            End If
        
        End If
        
        Fecha_Aux = DateAdd("d", 1, Fecha_Aux)
    Loop

'--------------------------------------------------------------
Monto = Comision
for_Comision = Comision
Bien = True
exito = True
End Function


Public Function CalcularHoras(ByVal Ternro As Long, ByVal Thnro As Long, ByVal Fecha As Date) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna las horas del acumulado diario
' ---------------------------------------------------------------------------------------------
'   Parametros Empleado,Tipo de Hora y Fecha
'   si el parametro ternro = 0 ==> lo calcula para todos los empleados
' ---------------------------------------------------------------------------------------------
' Autor      : FGZ
' Fecha      : 28/09/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_gti_adiario As New ADODB.Recordset

Dim Aux_Horas As Double


    'Obtengo las horas trabajadas por dia
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtengo las horas trabajadas para el dia: " & Fecha
    End If
    Aux_Horas = 0
    
    StrSql = "SELECT sum (adcanthoras) horas FROM gti_acumdiario "
    StrSql = StrSql & " WHERE ( " & ConvFecha(Fecha) & " = adfecha "
    StrSql = StrSql & " AND adfecha = " & ConvFecha(Fecha) & ")"
    StrSql = StrSql & " AND (gti_acumdiario.thnro  = " & Thnro & ")"
    If Ternro <> 0 Then
        StrSql = StrSql & " AND ternro  = " & Ternro
    End If
    OpenRecordset StrSql, rs_gti_adiario

    Do While Not rs_gti_adiario.EOF
        If Not EsNulo(rs_gti_adiario!Horas) Then
            Aux_Horas = rs_gti_adiario!Horas
        End If
       
        rs_gti_adiario.MoveNext
    Loop
    
    'Cierro y libero
    If rs_gti_adiario.State = adStateOpen Then rs_gti_adiario.Close
    Set rs_gti_adiario = Nothing
    
    CalcularHoras = Aux_Horas

End Function


Public Function Nov_Global(ByVal Estricta As Boolean, ByVal Concepto As Long, ByVal Tpanro As Long, ByVal Fecha As Date) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que retorna la Novedad Global
' ---------------------------------------------------------------------------------------------
'   Parametros:
'               Estricta (solo con vigencia para la fecha o todas las vigentes o sin vigencia),
'               Concepto,
'               Parametro y
'               Fecha
'   Busca novedades con vigencia para la fecha
' ---------------------------------------------------------------------------------------------
' Autor      : FGZ
' Fecha      : 28/09/2012
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_NovGral As New ADODB.Recordset
Dim rs_firmas As New ADODB.Recordset
Dim Firmado As Boolean
Dim Aux_Val As Double
Dim Encontro As Boolean
Dim Aux_Encontro As Boolean
Dim Vigencia_Activa As Boolean


    Encontro = False
    Aux_Val = 0
    Firmado = True

        StrSql = "SELECT concnro,tpanro,ngranro,ngravigencia,ngradesde,ngrahasta,ngravalor FROM novgral WHERE "
        StrSql = StrSql & " concnro = " & Concepto
        StrSql = StrSql & " AND tpanro = " & Tpanro
        StrSql = StrSql & " AND ((ngravigencia = -1 "
        StrSql = StrSql & " AND ngradesde = " & ConvFecha(Fecha)
        StrSql = StrSql & " AND (ngrahasta = " & ConvFecha(Fecha)
        StrSql = StrSql & " OR ngrahasta is null)) "
        StrSql = StrSql & " OR ngravigencia = 0) "
        StrSql = StrSql & " ORDER BY ngravigencia, ngradesde, ngrahasta "
        OpenRecordset StrSql, rs_NovGral
        Do While Not rs_NovGral.EOF
            If FirmaActiva19 Then
                    'Verificar si esta en el NIVEL FINAL DE FIRMA
                    StrSql = "select cysfirfin from cysfirmas where cysfiryaaut = -1 AND cysfirfin = -1 " & _
                             " AND cysfircodext = '" & rs_NovGral!ngranro & "' and cystipnro = 19"
                    OpenRecordset StrSql, rs_firmas
                    If rs_firmas.EOF Then
                        Firmado = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                        End If
                    Else
                        Firmado = True
                    End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If
        
            If Firmado Then
                If Not Estricta Then
                    If CBool(rs_NovGral!ngravigencia) Then
                        Vigencia_Activa = True
                        If Not EsNulo(rs_NovGral!ngrahasta) Then
                            If (rs_NovGral!ngrahasta < Fecha) Or (Fecha < rs_NovGral!ngradesde) Then
                                Vigencia_Activa = False
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta " & rs_NovGral!ngrahasta & " INACTIVA con valor " & rs_NovGral!ngravalor
                                End If
                            Else
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta " & rs_NovGral!ngrahasta & " ACTIVA con valor " & rs_NovGral!ngravalor
                                End If
                            End If
                        Else
                            If (Fecha < rs_NovGral!ngradesde) Then
                                Vigencia_Activa = False
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta indeterminado INACTIVA con valor " & rs_NovGral!ngravalor
                                End If
                            Else
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta indeterminado ACTIVA con valor " & rs_NovGral!ngravalor
                                End If
                            End If
                        End If
                    End If
                Else
                    If CBool(rs_NovGral!ngravigencia) Then
                        Vigencia_Activa = True
                        If Not EsNulo(rs_NovGral!ngrahasta) Then
                            If (rs_NovGral!ngrahasta = Fecha) And (Fecha = rs_NovGral!ngradesde) Then
                                Vigencia_Activa = True
                            Else
                                Vigencia_Activa = False
                            End If
                        Else
                            Vigencia_Activa = False
                        End If
                    End If
                End If
                If Vigencia_Activa Or Not CBool(rs_NovGral!ngravigencia) Then
                    Aux_Val = Aux_Val + rs_NovGral!ngravalor
                    
                    Encontro = True
                End If
            End If
            
            rs_NovGral.MoveNext
        Loop
        
        If Encontro Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Novedad global encontrada"
            End If
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontró ninguna Novedad global"
            End If
        End If

    'Cierro y libero
    If rs_NovGral.State = adStateOpen Then rs_NovGral.Close
    Set rs_NovGral = Nothing

    Nov_Global = Aux_Val

End Function

