Attribute VB_Name = "MdlFormulasGlencore"
Option Explicit
' ---------------------------------------------------------
' Modulo de fórmulas ccustomizadas para Glencore.
' ---------------------------------------------------------


Public Function for_ProvSac(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: .
' Autor      :
' Fecha      : FGZ - 22/07/2005
' Ultima Mod.:
' Traducccion:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Par_Base_Sac    As Integer
Dim Par_Mes         As Integer

Dim Concepto_Provi  As String

Dim Base_Sac        As Double
Dim Mes_Semestre    As Integer

Dim Ya_Provisionado As Double
Dim Mes_Ini_Sem     As Integer
Dim Mes_Fin_Sem     As Integer
Dim Mes_Actual      As Integer
Dim Anio            As Integer

Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim Par1 As Boolean
Dim Par2 As Boolean

'Inicializo
Bien = False

Par_Base_Sac = 51
Par_Mes = 78

Base_Sac = 0
Mes_Semestre = 0
Concepto_Provi = "12200"
Par1 = False
Par2 = False

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case Par_Base_Sac:
        Base_Sac = rs_wf_tpa!Valor
        Par1 = True
    Case Par_Mes:
        Mes_Semestre = rs_wf_tpa!Valor
        Par2 = True
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Base SAC " & Base_Sac
    Flog.writeline Espacios(Tabulador * 1) & "Mes Semestre " & Mes_Semestre
End If

'Si no se obtuvieron todos los parametros, entonces dejo la formula.
If Not (Par1 And Par2) Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No se obtuvieron todos los parametros, exit"
        Flog.writeline Espacios(Tabulador * 1) & "Base SAC " & Par_Base_Sac
        Flog.writeline Espacios(Tabulador * 1) & "Mes Semestre " & Mes_Semestre
    End If
    Exit Function
Else
    If Mes_Semestre = 6 Or Mes_Semestre = 12 Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "Parametro mes con Valor " & Mes_Semestre & ", exit"
        End If
        Exit Function
    End If
End If

'Busco la cabecera del impuesto
StrSql = "SELECT * FROM concepto "
StrSql = StrSql & " WHERE conccod ='" & Concepto_Provi & "'"
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
OpenRecordset StrSql, rs_Concepto
If rs_Concepto.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe el concepto " & Concepto_Provi
    End If
    Exit Function
Else
    Ya_Provisionado = 0
    Mes_Ini_Sem = IIf(buliq_periodo!pliqmes >= 7, 7, 1)
    Mes_Fin_Sem = IIf(buliq_periodo!pliqmes >= 7, 11, 5)
    Mes_Actual = buliq_periodo!pliqmes
    Anio = buliq_periodo!pliqanio
    Mes_Semestre = IIf(Mes_Semestre >= 7, Mes_Semestre - 6, Mes_Semestre)
    
    'Calcular todo lo provisionado en el semestre, en meses anteriores.
    StrSql = "SELECT sum(detliq.dlimonto) as Total FROM periodo "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " WHERE periodo.pliqanio = " & Anio
    StrSql = StrSql & " AND periodo.pliqmes >= " & Mes_Ini_Sem
    StrSql = StrSql & " AND periodo.pliqmes <= " & Mes_Fin_Sem
    StrSql = StrSql & " AND detliq.concnro = " & rs_Concepto!concnro
    StrSql = StrSql & " AND cabliq.empleado = " & buliq_cabliq!Empleado
    OpenRecordset StrSql, rs_Detliq
    If Not rs_Detliq.EOF Then
        Ya_Provisionado = IIf(Not EsNulo(rs_Detliq!Total), rs_Detliq!Total, 0)
    End If
End If

Monto = (Base_Sac / 12 * Mes_Semestre) - Ya_Provisionado
Bien = True

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 10, "Ya Provisionado ", Ya_Provisionado)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Mes Semestre ", Mes_Semestre)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Base Sac ", Base_Sac)
End If

for_ProvSac = Monto
exito = Bien

'Cierro y libero todo
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close

Set rs_wf_tpa = Nothing
Set rs_Concepto = Nothing
Set rs_Detliq = Nothing
End Function


Public Function for_ProvVac(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: .
' Autor      :
' Fecha      : FGZ - 22/07/2005
' Ultima Mod.:
' Traducccion:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Par_Base_Vac    As Integer
Dim Par_Mes         As Integer
Dim Par_Dias        As Integer
Dim Par_Prv         As Integer
Dim Par_Div1        As Integer
Dim Par_Div2        As Integer
Dim Par_Mul         As Integer

Dim Concepto_Provi  As String
Dim Concepto_Vac  As String

Dim Base_Vac        As Double
Dim Prom_Vac        As Double
Dim Div_Vac1        As Double
Dim Div_Vac2        As Double
Dim Mul_Vac         As Double
Dim Mes_Anio        As Integer
Dim Dias_Vac        As Integer
Dim Mes_Actual      As Integer
Dim Ya_Provisionado As Double
Dim Anio            As Integer

Dim rs_wf_tpa As New ADODB.Recordset
Dim rs_Concepto As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim Par1 As Boolean
Dim Par2 As Boolean
Dim Par3 As Boolean
Dim Par4 As Boolean
Dim Par5 As Boolean
Dim Par6 As Boolean
Dim Par7 As Boolean

'Inicializo
Bien = False

Par_Base_Vac = 51
Par_Mes = 78
Par_Dias = 29
Par_Prv = 80
Par_Div1 = 54
Par_Div2 = 143
Par_Mul = 149

Base_Vac = 0
Prom_Vac = 0
Div_Vac1 = 0
Div_Vac2 = 0
Mul_Vac = 0
Mes_Anio = 0
Dias_Vac = 0
Concepto_Provi = "12400"
Concepto_Vac = "02100"

Par1 = False
Par2 = False
Par3 = False
Par4 = False
Par5 = False
Par6 = False
Par7 = False

If HACE_TRAZA Then
    Call LimpiarTraza(Buliq_Concepto(Concepto_Actual).concnro)
End If

'Obtencion de los parametros de WorkFile
StrSql = "SELECT * FROM " & TTempWF_tpa & " WHERE fecha=" & ConvFecha(AFecha)
OpenRecordset StrSql, rs_wf_tpa

Do While Not rs_wf_tpa.EOF
    Select Case rs_wf_tpa!tipoparam
    Case Par_Base_Vac:
        Base_Vac = rs_wf_tpa!Valor
        Par1 = True
    Case Par_Mes:
        Mes_Anio = rs_wf_tpa!Valor
        Par2 = True
    Case Par_Dias:
        Dias_Vac = rs_wf_tpa!Valor
        Par3 = True
    Case Par_Prv:
        Prom_Vac = rs_wf_tpa!Valor
        Par4 = True
    Case Par_Div1:
        Div_Vac1 = rs_wf_tpa!Valor
        Par5 = True
    Case Par_Div2:
        Div_Vac2 = rs_wf_tpa!Valor
        Par6 = True
    Case Par_Mul:
        Mul_Vac = rs_wf_tpa!Valor
        Par7 = True
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Base VAC " & Base_Vac
    Flog.writeline Espacios(Tabulador * 1) & "Mes Semestre " & Mes_Anio
    Flog.writeline Espacios(Tabulador * 1) & "Dias VAC " & Dias_Vac
    Flog.writeline Espacios(Tabulador * 1) & "Promedio VAC " & Prom_Vac
    Flog.writeline Espacios(Tabulador * 1) & "Div 1 " & Div_Vac1
    Flog.writeline Espacios(Tabulador * 1) & "Div 2 " & Div_Vac2
    Flog.writeline Espacios(Tabulador * 1) & "Mul " & Mul_Vac
End If


'Si no se obtuvieron todos los parametros, entonces dejo la formula.
'If Base_Vac = 0 Or Mes_Anio = 0 Or Dias_Vac = 0 Or Prom_Vac = 0 Then
If Not (Par1 And Par2 And Par3 And Par4) Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No se obtuvieron todos los parametros, exit"
        Flog.writeline Espacios(Tabulador * 1) & "Base VAC " & Par_Base_Vac
        Flog.writeline Espacios(Tabulador * 1) & "Mes Semestre " & Par_Mes
        Flog.writeline Espacios(Tabulador * 1) & "Dias VAC " & Par_Dias
        Flog.writeline Espacios(Tabulador * 1) & "Prom AVC " & Par_Prv
    End If
    Exit Function
End If

Anio = IIf(Not EsNulo(buliq_periodo!pliqanio), buliq_periodo!pliqanio, 0)
Mes_Actual = IIf(Not EsNulo(buliq_periodo!pliqmes), buliq_periodo!pliqmes, 0)

'Si durante el año de la provision se pago vacaciones, entonces no debo seguir
StrSql = "SELECT * FROM concepto "
StrSql = StrSql & " WHERE conccod ='" & Concepto_Vac & "'"
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
OpenRecordset StrSql, rs_Concepto
If rs_Concepto.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe el concepto " & Concepto_Vac
    End If
    Exit Function
Else
    StrSql = "SELECT * FROM periodo "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " WHERE periodo.pliqanio = " & Anio
    StrSql = StrSql & " AND periodo.pliqmes >= 9"
    StrSql = StrSql & " AND detliq.concnro = " & rs_Concepto!concnro
    StrSql = StrSql & " AND cabliq.empleado = " & buliq_cabliq!Empleado
    OpenRecordset StrSql, rs_Detliq
    If Not rs_Detliq.EOF Then
        'existe una desaprovision realizada despues del mes 9, con lo cual ya se le pagaron
        ' las vacaciones y por este año no provisiono mas
        Exit Function
    End If
End If

'Busco la cabecera del impuesto
StrSql = "SELECT * FROM concepto "
StrSql = StrSql & " WHERE conccod ='" & Concepto_Provi & "'"
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
OpenRecordset StrSql, rs_Concepto
If rs_Concepto.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe el concepto " & Concepto_Provi
    End If
    Exit Function
Else
    Ya_Provisionado = 0
    
    'Calcular todo lo provisionado en el semestre, en meses anteriores.
    StrSql = "SELECT sum(detliq.dlimonto) as Total FROM periodo "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " WHERE periodo.pliqanio = " & Anio
    StrSql = StrSql & " AND periodo.pliqmes <= " & Mes_Actual
    StrSql = StrSql & " AND detliq.concnro = " & rs_Concepto!concnro
    StrSql = StrSql & " AND cabliq.empleado = " & buliq_cabliq!Empleado
    OpenRecordset StrSql, rs_Detliq
    If Not rs_Detliq.EOF Then
        If Not EsNulo(rs_Detliq!Total) Then
            Ya_Provisionado = rs_Detliq!Total
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "Ya Provisionado = 0 "
                Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
            End If
            Ya_Provisionado = 0
        End If
        If Ya_Provisionado = 0 Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 1) & "Ya Provisionado = 0 "
                Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
            End If
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "No existe datos de lo provisionado en el semestre "
            Flog.writeline Espacios(Tabulador * 1) & "SQL: " & StrSql
        End If
    End If
End If

'Busco la forma de liquidacion
StrSql = " SELECT estructura.estrnro, estructura.estrcodext  FROM his_estructura"
StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro "
StrSql = StrSql & " WHERE ternro = " & buliq_empleado!ternro & " AND "
StrSql = StrSql & " his_estructura.tenro = 22 AND "
StrSql = StrSql & " (htetdesde <= " & ConvFecha(Empleado_Fecha_Fin) & ") AND "
StrSql = StrSql & " ((" & ConvFecha(Empleado_Fecha_Fin) & " <= htethasta) or (htethasta is null))"
OpenRecordset StrSql, rs_Estructura
If Not rs_Estructura.EOF Then
    If Not EsNulo(rs_Estructura!estrcodext) Then
        If rs_Estructura!estrcodext = "2" Then
            Monto = ((Base_Vac + Prom_Vac) / Div_Vac1 * Dias_Vac * Mul_Vac / 12 * Mes_Anio) - Ya_Provisionado
        Else
            Monto = ((((Base_Vac + Prom_Vac) / Div_Vac1 * Dias_Vac) - ((Base_Vac + Prom_Vac) / Div_Vac2 * Dias_Vac)) * Mul_Vac / 12 * Mes_Anio) - Ya_Provisionado
        End If
    Else
        Monto = ((((Base_Vac + Prom_Vac) / Div_Vac1 * Dias_Vac) - ((Base_Vac + Prom_Vac) / Div_Vac2 * Dias_Vac)) * Mul_Vac / 12 * Mes_Anio) - Ya_Provisionado
    End If
Else
    Monto = ((((Base_Vac + Prom_Vac) / Div_Vac1 * Dias_Vac) - ((Base_Vac + Prom_Vac) / Div_Vac2 * Dias_Vac)) * Mul_Vac / 12 * Mes_Anio) - Ya_Provisionado
End If
Bien = True

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 10, "Ya Provisionado ", Ya_Provisionado)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Mes Semestre ", Mes_Anio)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Base Vac ", Base_Vac)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Prom Vac ", Prom_Vac)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Mul Vac ", Mul_Vac)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Div Vac1 ", Div_Vac1)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Div Vac2 ", Div_Vac2)
End If
for_ProvVac = Monto
exito = Bien

'Cierro y libero todo
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close

Set rs_wf_tpa = Nothing
Set rs_Concepto = Nothing
Set rs_Detliq = Nothing
Set rs_Estructura = Nothing
End Function


Public Function for_DesaProvSac(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: .
' Autor      :
' Fecha      : FGZ - 22/07/2005
' Ultima Mod.:
' Traducccion:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Par_Prov_Sac    As Integer
Dim Par_Mes         As Integer
Dim Mes_Semestre    As Integer
Dim Ya_Provisionado As Double

Dim rs_wf_tpa As New ADODB.Recordset

'Inicializo
Bien = False
Par_Prov_Sac = 51
Par_Mes = 78

Ya_Provisionado = 0
Mes_Semestre = 0

' Primero limpio la traza
StrSql = "DELETE FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
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
    Case Par_Prov_Sac:
        Ya_Provisionado = rs_wf_tpa!Valor
    Case Par_Mes:
        Mes_Semestre = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Provision SAC " & Ya_Provisionado
    Flog.writeline Espacios(Tabulador * 1) & "Mes Semestre " & Mes_Semestre
End If

'Si no se obtuvieron todos los parametros, entonces dejo la formula.
If Ya_Provisionado = 0 Or Mes_Semestre = 0 Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No se obtuvieron todos los parametros, exit"
        Flog.writeline Espacios(Tabulador * 1) & "Provision SAC " & Par_Prov_Sac
        Flog.writeline Espacios(Tabulador * 1) & "Mes Semestre " & Mes_Semestre
    End If
    Exit Function
End If

Monto = 0
If Mes_Semestre = 6 Or Mes_Semestre = 12 Then
    Monto = -Ya_Provisionado
End If
Bien = True

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 10, "Ya Provisionado ", Ya_Provisionado)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Mes Semestre ", Mes_Semestre)
End If
for_DesaProvSac = Monto
exito = Bien

'Cierro y libero todo
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
Set rs_wf_tpa = Nothing
End Function


Public Function for_DesaProvVac(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: .
' Autor      :
' Fecha      : FGZ - 22/07/2005
' Ultima Mod.:
' Traducccion:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Par_Cero        As Integer
Dim Concepto_Provi  As String
Dim Concepto_Vac    As String
Dim Cero            As Double
Dim Ya_Provisionado As Double
Dim Anio            As Integer
Dim Pago_Vac        As Boolean

Dim rs_wf_tpa       As New ADODB.Recordset
Dim rs_Concepto     As New ADODB.Recordset
Dim rs_Detliq       As New ADODB.Recordset

'Inicializo
Bien = False

Par_Cero = 51

Cero = 0
Concepto_Provi = "12400" 'concepto de Provision LAR
Concepto_Vac = "02100"

' Primero limpio la traza
StrSql = "DELETE FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
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
    Case Par_Cero:
        Cero = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Cer " & Cero
End If

'Si no se obtuvieron todos los parametros, entonces dejo la formula.
If Cero = 0 Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No se obtuvieron todos los parametros, exit"
        Flog.writeline Espacios(Tabulador * 1) & "Cero " & Par_Cero
    End If
    Exit Function
End If

'verifica que en el periodo se haya pagado vacaciones para el empleado
StrSql = "SELECT * FROM concepto "
StrSql = StrSql & " WHERE conccod ='" & Concepto_Vac & "'"
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
OpenRecordset StrSql, rs_Concepto
If rs_Concepto.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe el concepto " & Concepto_Vac
    End If
    Exit Function
Else
    StrSql = "SELECT * FROM periodo "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " WHERE periodo.pliqnro = " & buliq_periodo!PliqNro
    StrSql = StrSql & " AND detliq.concnro = " & rs_Concepto!concnro
    StrSql = StrSql & " AND cabliq.empleado = " & buliq_cabliq!Empleado
    OpenRecordset StrSql, rs_Detliq
    If Not rs_Detliq.EOF Then
        Pago_Vac = True
    Else
        Pago_Vac = False
    End If
End If

If Pago_Vac Then
    'Busco la cabecera del impuesto
    StrSql = "SELECT * FROM concepto "
    StrSql = StrSql & " WHERE conccod ='" & Concepto_Provi & "'"
    If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
    OpenRecordset StrSql, rs_Concepto
    If rs_Concepto.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "No existe el concepto " & Concepto_Provi
        End If
        Exit Function
    Else
        Ya_Provisionado = 0
        Anio = IIf(buliq_periodo!pliqmes < 9, buliq_periodo!pliqanio - 1, buliq_periodo!pliqanio)
        
        'Calcular todo lo provisionado.
        StrSql = "SELECT sum(detliq.dlimonto) as Total FROM periodo "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " WHERE periodo.pliqanio = " & Anio
        StrSql = StrSql & " AND periodo.pliqmes <= 12"
        StrSql = StrSql & " AND detliq.concnro = " & rs_Concepto!concnro
        StrSql = StrSql & " AND cabliq.empleado = " & buliq_cabliq!Empleado
        OpenRecordset StrSql, rs_Detliq
        If Not rs_Detliq.EOF Then
            Ya_Provisionado = IIf(Not EsNulo(rs_Detliq!Total), rs_Detliq!Total, 0)
        End If
    End If
    Monto = -Ya_Provisionado
    Bien = True
End If

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 10, "Ya Provisionado ", Ya_Provisionado)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Añio ", Anio)
End If
for_DesaProvVac = Monto
exito = Bien

'Cierro y libero todo
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close

Set rs_wf_tpa = Nothing
Set rs_Concepto = Nothing
Set rs_Detliq = Nothing
End Function


Public Function for_PorcPres(ByVal NroCab As Long, ByVal AFecha As Date, Monto As Double, Bien As Boolean) As Double
' ---------------------------------------------------------------------------------------------
' Descripcion: .
' Autor      :
' Fecha      : FGZ - 22/07/2005
' Ultima Mod.:
' Traducccion:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Par_Cero        As Integer
Dim Concepto_Provi  As String
Dim Concepto_Vac    As String
Dim Cero            As Double
Dim Ya_Provisionado As Double
Dim Anio            As Integer
Dim Pago_Vac        As Boolean

Dim rs_wf_tpa       As New ADODB.Recordset
Dim rs_Concepto     As New ADODB.Recordset
Dim rs_Detliq       As New ADODB.Recordset

'Inicializo
Bien = False

Par_Cero = 51

Cero = 0
Concepto_Provi = "12400" 'concepto de Provision LAR
Concepto_Vac = "02100"

' Primero limpio la traza
StrSql = "DELETE FROM traza_gan WHERE "
StrSql = StrSql & "pliqnro =" & buliq_periodo!PliqNro
StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
StrSql = StrSql & " AND concnro =" & Buliq_Concepto(Concepto_Actual).concnro
StrSql = StrSql & " AND fecha_pago =" & ConvFecha(buliq_proceso!profecpago)
'StrSql = StrSql & " AND empresa =" & NroEmp
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
    Case Par_Cero:
        Cero = rs_wf_tpa!Valor
    End Select
    
    rs_wf_tpa.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 1) & "Cer " & Cero
End If

'Si no se obtuvieron todos los parametros, entonces dejo la formula.
If Cero = 0 Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No se obtuvieron todos los parametros, exit"
        Flog.writeline Espacios(Tabulador * 1) & "Cero " & Par_Cero
    End If
    Exit Function
End If

'verifica que en el periodo se haya pagado vacaciones para el empleado
StrSql = "SELECT * FROM concepto "
StrSql = StrSql & " WHERE conccod ='" & Concepto_Vac & "'"
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
OpenRecordset StrSql, rs_Concepto
If rs_Concepto.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 1) & "No existe el concepto " & Concepto_Vac
    End If
    Exit Function
Else
    StrSql = "SELECT * FROM periodo "
    StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
    StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
    StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
    StrSql = StrSql & " WHERE periodo.pliqanio = " & Anio
    StrSql = StrSql & " AND detliq.concnro = " & rs_Concepto!concnro
    StrSql = StrSql & " AND cabliq.empleado = " & buliq_cabliq!Empleado
    OpenRecordset StrSql, rs_Detliq
    If Not rs_Detliq.EOF Then
        Pago_Vac = True
    Else
        Pago_Vac = False
    End If
End If

If Pago_Vac Then
    'Busco la cabecera del impuesto
    StrSql = "SELECT * FROM concepto "
    StrSql = StrSql & " WHERE conccod ='" & Concepto_Provi & "'"
    If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
    OpenRecordset StrSql, rs_Concepto
    If rs_Concepto.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "No existe el concepto " & Concepto_Provi
        End If
        Exit Function
    Else
        Ya_Provisionado = 0
        Anio = IIf(buliq_periodo!pliqmes < 9, buliq_periodo!pliqanio - 1, buliq_periodo!pliqanio)
        
        'Calcular todo lo provisionado.
        StrSql = "SELECT sum(detliq.dlimonto) as Total FROM periodo "
        StrSql = StrSql & " INNER JOIN proceso ON proceso.pliqnro = periodo.pliqnro "
        StrSql = StrSql & " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro "
        StrSql = StrSql & " INNER JOIN detliq ON detliq.cliqnro = cabliq.cliqnro "
        StrSql = StrSql & " WHERE periodo.pliqanio = " & Anio
        StrSql = StrSql & " AND periodo.pliqmes <= 12"
        StrSql = StrSql & " AND detliq.concnro = " & rs_Concepto!concnro
        StrSql = StrSql & " AND cabliq.empleado = " & buliq_cabliq!Empleado
        OpenRecordset StrSql, rs_Detliq
        If Not rs_Detliq.EOF Then
            Ya_Provisionado = IIf(Not EsNulo(rs_Detliq!Total), rs_Detliq!Total, 0)
        End If
    End If
    Monto = -Ya_Provisionado
    Bien = True
End If

If HACE_TRAZA Then
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 10, "Ya Provisionado ", Ya_Provisionado)
    Call InsertarTraza(buliq_cabliq!cliqnro, Buliq_Concepto(Concepto_Actual).concnro, 11, "Añio ", Anio)
End If
for_PorcPres = Monto
exito = Bien

'Cierro y libero todo
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close

Set rs_wf_tpa = Nothing
Set rs_Concepto = Nothing
Set rs_Detliq = Nothing
End Function

