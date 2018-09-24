Attribute VB_Name = "MdlBuliq"
Option Explicit



Public Sub Establecer_Buliq_concepto(ByVal nroconcepto As Long, ByRef OK As Boolean)
    StrSql = "SELECT * FROM concepto WHERE concnro = " & nroconcepto
    OpenRecordset StrSql, rs_Buliq_Concepto

    OK = Not rs_Buliq_Concepto.EOF
End Sub


Public Sub Establecer_Empleado(ByVal p_ternro As Long, ByVal p_grunro As Long, ByVal p_cliqnro As Long, ByVal p_fecha_inicio As Date, ByVal p_fecha_fin As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea las variable globales del EMPLEADO con los valores pasados por parametros.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.: FGZ - 04/09/2006
' Descripcion: Cuando la fecha de baja del legajo que seguia era mayor que el anterior no seteaba la fecha de baja,
'               dejaba la fecha de baja del legajo anterior.
' ---------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset

    NroEmple = p_ternro
    NroGrupo = p_grunro
    NroCab = p_cliqnro

    ' Actualizar los buffer's Auxiliares
    ' Empleado
    'FGZ - 20/09/2011 - ahora puede haber NN con lo cual esta consulta puede fallar (igualmente simpre se debiera buscar en la tabla sim_empleado)
    'StrSql = "SELECT * FROM empleado WHERE ternro = " & CStr(p_ternro)
    StrSql = "SELECT * FROM sim_empleado WHERE ternro = " & CStr(p_ternro)
    OpenRecordset StrSql, buliq_empleado

    If buliq_empleado.EOF Then
        If CBool(USA_DEBUG) Then
            'Flog.writeline Espacios(Tabulador * 1) & "no levanto ningun empleado de empleado"
            Flog.writeline Espacios(Tabulador * 1) & "El tercero no está en sim_empleado"
        End If
    Else
        Legajo = buliq_empleado!Empleg
    End If
    ' FGZ - 18/03/2004
    ' Si el empleado no esta activo ==> seteo la fecha de baja
    
    'FGZ - 04/09/2006 - Inicializo
    Empleado_Fecha_Inicio = buliq_proceso!profecini
    Empleado_Fecha_Fin = buliq_proceso!profecfin
    
    If Not CBool(buliq_empleado!empest) Then
        StrSql = "SELECT * FROM sim_fases WHERE real = -1 AND empleado = " & buliq_empleado!Ternro
        StrSql = StrSql & " AND (altfec <= " & ConvFecha(Fecha_Fin) & ") "
        'StrSql = StrSql & " AND ((altfec >= " & ConvFecha(Fecha_Inicio) & " AND altfec <= " & ConvFecha(Fecha_Fin) & ") "
        'StrSql = StrSql & " OR (bajfec <= " & ConvFecha(Fecha_Fin) & "))"
        StrSql = StrSql & " ORDER BY altfec"
        OpenRecordset StrSql, rs_Fases
        If Not rs_Fases.EOF Then rs_Fases.MoveLast
        If Not rs_Fases.EOF Then
            If Not CBool(rs_Fases!Estado) Then
                'FGZ - 04/09/2006 - Estaba mal el if
                'If rs_Fases!bajfec < Empleado_Fecha_Fin And rs_Fases!bajfec > Empleado_Fecha_Inicio Then
                If rs_Fases!bajfec < Empleado_Fecha_Fin And rs_Fases!bajfec >= Empleado_Fecha_Inicio Then
                    Empleado_Fecha_Fin = rs_Fases!bajfec
                End If
                
                Flog.writeline " El empleado no esta activo. fecha de baja : " & Empleado_Fecha_Fin
            End If
        End If
    End If
    ' FGZ - 18/03/2004
    If rs_Fases.State = adStateOpen Then rs_Fases.Close
    Set rs_Fases = Nothing

    ' Tercero
    StrSql = "SELECT * FROM tercero WHERE ternro = " & CStr(p_ternro)
    OpenRecordset StrSql, buliq_tercero_emp

    ' cabliq
    StrSql = "SELECT * FROM sim_cabliq WHERE cliqnro = " & CStr(p_cliqnro)
    OpenRecordset StrSql, buliq_cabliq
    
    ' Iniciar el cache para el empleado
    Call objCache.Limpiar
    Call objCache_detliq_Monto.Limpiar
    Call objCache_detliq_Cantidad.Limpiar
End Sub

Public Sub Establecer_Proceso(ByVal p_pronro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea las variable globales del PROCESO con los valores pasados por parametros.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    ' Actualizar el buffer Auxiliares
    ' Proceso
    StrSql = "SELECT * FROM sim_proceso WHERE pronro = " & CStr(p_pronro)
    OpenRecordset StrSql, buliq_proceso

    'FGZ - 18/03/2004
    Empleado_Fecha_Inicio = buliq_proceso!profecini
    Empleado_Fecha_Fin = buliq_proceso!profecfin
    'FGZ - 18/03/2004
    
    ' Periodo
    StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(buliq_proceso!PliqNro)
    OpenRecordset StrSql, buliq_periodo

    ' impgralarg
    StrSql = "SELECT * FROM sim_impgralarg WHERE pronro = " & CStr(p_pronro)
    OpenRecordset StrSql, buliq_impgralarg
    
End Sub

Public Sub Establecer_Impgralarg(ByVal pronro As Long, ByVal tconnro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: setea la variable global de los imponibles del Proceso.
' Autor      : FGZ
' Fecha      : 25/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    StrSql = "SELECT * FROM sim_impgralarg WHERE pronro = " & CStr(pronro) & _
             " AND tconnro =" & tconnro
    If buliq_impgralarg.State = adStateOpen Then buliq_impgralarg.Close
    OpenRecordset StrSql, buliq_impgralarg
    
End Sub

Public Sub Establecer_Empresa(ByVal FDesde As Date, ByVal FHasta As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: setea la variable global con el nro de empresa por cada empleado.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Empresa As New ADODB.Recordset

StrSql = "SELECT empresa.empnro FROM sim_his_estructura "
'StrSql = StrSql & " INNER JOIN estructura ON estructura.estrnro = his_estructura.estrnro "
StrSql = StrSql & " INNER JOIN empresa ON empresa.estrnro = sim_his_estructura.estrnro "
StrSql = StrSql & " WHERE (sim_his_estructura.tenro = 10 and sim_his_estructura.ternro =" & buliq_empleado!Ternro & ")"
StrSql = StrSql & " AND sim_his_estructura.htetdesde <=" & ConvFecha(FHasta)
StrSql = StrSql & " AND (sim_his_estructura.htethasta >= " & ConvFecha(FDesde)
StrSql = StrSql & " OR sim_his_estructura.htethasta IS NULL)"
OpenRecordset StrSql, rs_Empresa
If Not rs_Empresa.EOF Then
    NroEmp = rs_Empresa!Empnro
Else
    Flog.writeline "El empleado " & buliq_empleado!Empleg & " no tiene empresa asignada"
    NroEmp = 0
End If
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
Set rs_Empresa = Nothing

End Sub

Public Sub Establecer_Empresa_old(ByVal FDesde As Date, ByVal FHasta As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: setea la variable global con el nro de empresa por cada empleado.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Empresa As New ADODB.Recordset

StrSql = "SELECT emp.empnro FROM his_estructura "
StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE (his_estructura.tenro = 10 and his_estructura.ternro =" & buliq_empleado!Ternro & ")"
StrSql = StrSql & " AND his_estructura.htetdesde <=" & ConvFecha(FHasta)
StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(FDesde)
StrSql = StrSql & " OR his_estructura.htethasta IS NULL)"
OpenRecordset StrSql, rs_Empresa
If Not rs_Empresa.EOF Then
    NroEmp = rs_Empresa!Empnro
Else
    Flog.writeline "El empleado " & buliq_empleado!Empleg & " no tiene empresa asignada"
    NroEmp = 0
End If
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
Set rs_Empresa = Nothing

End Sub



Public Sub Establecer_Parametro(ByVal p_concnro As Long, ByVal p_tpanro As Long, ByVal p_prognro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea las variable globales del PARAMETRO con los valores pasados por parametros.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

    NroConce = p_concnro
    NroTpa = p_tpanro
    NroProg = p_prognro

    ' Actualizar el buffer de concepto
    StrSql = "SELECT * FROM concepto WHERE concnro = " & CStr(p_concnro)
    OpenRecordset StrSql, rs_Buliq_Concepto
    
End Sub

Public Sub EstablecerFirmas()
Dim rs_cystipo As New ADODB.Recordset

    
    FirmaActiva5 = False    'Novedades de liquidacion individuales
    FirmaActiva15 = False   'Novedades de liquidacion por estructura
    FirmaActiva19 = False   'Novedades de liquidacion globales
    FirmaActiva20 = False   'Novedades de Ajuste
    FirmaActiva165 = False  'Gastos
    
    'FGZ - 05/06/2012 -------------------------
    'StrSql = "select * from cystipo where (cystipnro = 5 or cystipnro = 15 OR cystipnro = 19 or cystipnro = 20) AND cystipact = -1"
    StrSql = "select cystipnro from cystipo where (cystipnro = 5 or cystipnro = 15 OR cystipnro = 19 or cystipnro = 20 or cystipnro = 165) AND cystipact = -1"
    
    OpenRecordset StrSql, rs_cystipo
    
    Do While Not rs_cystipo.EOF
    Select Case rs_cystipo!cystipnro
    Case 5:
        FirmaActiva5 = True
    Case 15:
        FirmaActiva15 = True
    Case 19:
        FirmaActiva19 = True
    Case 20:
        FirmaActiva20 = True
    Case 165:
        FirmaActiva165 = True
    Case Else
    End Select
        
        rs_cystipo.MoveNext
    Loop
    
    'FGZ - 27/06/2011 ---------------------------------
    ' se desactivo eol control de firmas del simulador
    FirmaActiva5 = False
    FirmaActiva15 = False
    FirmaActiva19 = False
    FirmaActiva20 = False
    FirmaActiva165 = False
    'FGZ - 27/06/2011 ---------------------------------
    ' se desactivo eol control de firmas del simulador
    
    
If rs_cystipo.State = adStateOpen Then rs_cystipo.Close
Set rs_cystipo = Nothing

End Sub


Public Sub CargarConceptos(ByVal Nrotipo As Long, ByVal Fecha_Inicio As Date, ByVal Fecha_Fin As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga todos los conceptos a liquidar.
' Autor      : FGZ
' Fecha      : 11/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_Conceptos As New ADODB.Recordset

    'CONCEPTO:
    StrSql = "SELECT * FROM concepto " & _
             " INNER JOIN con_tp ON con_tp.concnro = concepto.concnro " & _
             " INNER JOIN formula ON formula.fornro = concepto.fornro " & _
             " WHERE con_tp.tprocnro = " & Nrotipo & _
             " AND (concepto.concvalid = 0 or ( concdesde <= " & ConvFecha(Fecha_Inicio) & _
             " AND conchasta >= " & ConvFecha(Fecha_Fin) & "))" & _
             " ORDER BY concepto.tconnro, concepto.concorden"
    OpenRecordset StrSql, rs_Conceptos

    If rs_Conceptos.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 2) & "No se encontraron conceptos"
        End If
    Else
        Cantidad_de_Conceptos = rs_Conceptos.RecordCount
        rs_Conceptos.MoveFirst
    End If
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 2) & "Cantidad de conceptos " & Cantidad_de_Conceptos
    End If
    
    
    Max_Conceptos = Cantidad_de_Conceptos
    ReDim Preserve Arr_conceptos(Max_Conceptos) As TConcepto
    ReDim Preserve Buliq_Concepto(Max_Conceptos) As Tbuliq_concepto
    
    I = 1
    Do While Not rs_Conceptos.EOF
        If Not EsNulo(rs_Conceptos!ConcNro) Then Arr_conceptos(I).ConcNro = rs_Conceptos!ConcNro
        If Not EsNulo(rs_Conceptos!Conccod) Then Arr_conceptos(I).Conccod = rs_Conceptos!Conccod
        If Not EsNulo(rs_Conceptos!Concajuste) Then Arr_conceptos(I).Concajuste = rs_Conceptos!Concajuste
        If Not EsNulo(rs_Conceptos!Conccantdec) Then Arr_conceptos(I).Conccantdec = rs_Conceptos!Conccantdec
        If Not EsNulo(rs_Conceptos!concabr) Then Arr_conceptos(I).concabr = rs_Conceptos!concabr
        If Not EsNulo(rs_Conceptos!concretro) Then Arr_conceptos(I).concretro = rs_Conceptos!concretro
        If Not EsNulo(rs_Conceptos!tconnro) Then Arr_conceptos(I).tconnro = rs_Conceptos!tconnro
        If Not EsNulo(rs_Conceptos!Conctexto) Then Arr_conceptos(I).Conctexto = rs_Conceptos!Conctexto
        If Not EsNulo(rs_Conceptos!fornro) Then Arr_conceptos(I).fornro = rs_Conceptos!fornro
        If Not EsNulo(rs_Conceptos!Fortipo) Then Arr_conceptos(I).Fortipo = rs_Conceptos!Fortipo
        If Not EsNulo(rs_Conceptos!Forexpresion) Then Arr_conceptos(I).Forexpresion = rs_Conceptos!Forexpresion
        If Not EsNulo(rs_Conceptos!Fordabr) Then Arr_conceptos(I).Fordabr = rs_Conceptos!Fordabr
        If Not EsNulo(rs_Conceptos!Forprog) Then Arr_conceptos(I).Forprog = rs_Conceptos!Forprog
        Arr_conceptos(I).Seguir = True
        Arr_conceptos(I).NetoFijo = 0
        
        
        'buliq_concepto
        If Not EsNulo(rs_Conceptos!ConcNro) Then Buliq_Concepto(I).ConcNro = rs_Conceptos!ConcNro
        If Not EsNulo(rs_Conceptos!Conccod) Then Buliq_Concepto(I).Conccod = rs_Conceptos!Conccod
        If Not EsNulo(rs_Conceptos!Concajuste) Then Buliq_Concepto(I).Concajuste = rs_Conceptos!Concajuste
        If Not EsNulo(rs_Conceptos!Conccantdec) Then Buliq_Concepto(I).Conccantdec = rs_Conceptos!Conccantdec
        If Not EsNulo(rs_Conceptos!concabr) Then Buliq_Concepto(I).concabr = rs_Conceptos!concabr
        If Not EsNulo(rs_Conceptos!concretro) Then Buliq_Concepto(I).concretro = rs_Conceptos!concretro
        If Not EsNulo(rs_Conceptos!tconnro) Then Buliq_Concepto(I).tconnro = rs_Conceptos!tconnro
        If Not EsNulo(rs_Conceptos!Conctexto) Then Buliq_Concepto(I).Conctexto = rs_Conceptos!Conctexto
        If Not EsNulo(rs_Conceptos!concorden) Then Buliq_Concepto(I).concorden = rs_Conceptos!concorden
        If Not EsNulo(rs_Conceptos!concext) Then Buliq_Concepto(I).concext = rs_Conceptos!concext
        If Not EsNulo(rs_Conceptos!concvalid) Then Buliq_Concepto(I).concvalid = rs_Conceptos!concvalid
        If Not EsNulo(rs_Conceptos!concdesde) Then Buliq_Concepto(I).concdesde = rs_Conceptos!concdesde
        If Not EsNulo(rs_Conceptos!conchasta) Then Buliq_Concepto(I).conchasta = rs_Conceptos!conchasta
        If Not EsNulo(rs_Conceptos!concrepet) Then Buliq_Concepto(I).concrepet = rs_Conceptos!concrepet
        If Not EsNulo(rs_Conceptos!concniv) Then Buliq_Concepto(I).concniv = rs_Conceptos!concniv
        If Not EsNulo(rs_Conceptos!fornro) Then Buliq_Concepto(I).fornro = rs_Conceptos!fornro
        If Not EsNulo(rs_Conceptos!concimp) Then Buliq_Concepto(I).concimp = rs_Conceptos!concimp
        If Not EsNulo(rs_Conceptos!codseguridad) Then Buliq_Concepto(I).codseguridad = rs_Conceptos!codseguridad
        If Not EsNulo(rs_Conceptos!concusado) Then Buliq_Concepto(I).concusado = rs_Conceptos!concusado
        If Not EsNulo(rs_Conceptos!concpuente) Then Buliq_Concepto(I).concpuente = rs_Conceptos!concpuente
        If Not EsNulo(rs_Conceptos!Empnro) Then Buliq_Concepto(I).Empnro = rs_Conceptos!Empnro
        If Not EsNulo(rs_Conceptos!concautor) Then Buliq_Concepto(I).concautor = rs_Conceptos!concautor
        If Not EsNulo(rs_Conceptos!concfecmodi) Then Buliq_Concepto(I).concfecmodi = rs_Conceptos!concfecmodi
        If Not EsNulo(rs_Conceptos!Concajuste) Then Buliq_Concepto(I).Concajuste = rs_Conceptos!Concajuste

        rs_Conceptos.MoveNext
        I = I + 1
    Loop
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 2) & "I =  " & I
    End If
    
    If rs_Conceptos.State = adStateOpen Then rs_Conceptos.Close
    Set rs_Conceptos = Nothing

End Sub


Public Sub CargarCabecerasLiq(ByVal Todos As Boolean, ByVal NroProc As String, ByVal Bpronro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga todos los empleados a liquidar.
' Autor      : FGZ
' Fecha      : 11/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_Empleados As New ADODB.Recordset

    ' Liquido los empleados
    If Todos Then
        StrSql = "SELECT * FROM sim_cabliq WHERE pronro =" & NroProc
    Else
        StrSql = "SELECT * FROM sim_cabliq " & _
        " INNER JOIN batch_empleado ON batch_empleado.ternro = sim_cabliq.empleado " & _
        " WHERE sim_cabliq.pronro =" & NroProc & _
        " AND batch_empleado.bpronro = " & Bpronro
    End If
    OpenRecordset StrSql, rs_Empleados

    Max_Cabeceras = rs_Empleados.RecordCount
    ReDim Preserve Arr_EmpCab(Max_Cabeceras) As TEmpCabLiq
    
    I = 1
    Do While Not rs_Empleados.EOF
        Arr_EmpCab(I).cliqnro = rs_Empleados!cliqnro
        Arr_EmpCab(I).Empleado = rs_Empleados!Empleado
        'Arr_EmpCab(I).Ternro = rs_Empleados!Ternro
        
        I = I + 1
        rs_Empleados.MoveNext
        
    Loop
    
    Cantidad_de_Empleados = I - 1
    If rs_Empleados.State = adStateOpen Then rs_Empleados.Close
    Set rs_Empleados = Nothing

End Sub


Public Sub CargarBusquedas()
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga todas las Busquedas generadas.
' Autor      : FGZ
' Fecha      : 13/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Long
Dim rs_Programa As New ADODB.Recordset

    StrSql = "SELECT * FROM programa "
    StrSql = StrSql & " ORDER BY prognro "
    OpenRecordset StrSql, rs_Programa

    If Not rs_Programa.EOF Then
        'LLENA EL ARRAY CON LOS NOMBRES DE LAS TABLAS USADAS EN LA SIMULACION
        Call Carga_Tablas_SIM
        
        rs_Programa.MoveLast
        
        Max_Programas = rs_Programa!Prognro + 1
        ReDim Preserve Arr_Programa(Max_Programas) As TPrograma
        rs_Programa.MoveFirst
    End If
    
    Do While Not rs_Programa.EOF
        If Not EsNulo(rs_Programa!Prognro) Then
            I = rs_Programa!Prognro
            Arr_Programa(I).Prognro = rs_Programa!Prognro
            If Not EsNulo(rs_Programa!Prognom) Then Arr_Programa(I).Prognom = rs_Programa!Prognom
            If Not EsNulo(rs_Programa!Progdesc) Then Arr_Programa(I).Progdesc = rs_Programa!Progdesc
            If Not EsNulo(rs_Programa!Tprognro) Then Arr_Programa(I).Tprognro = rs_Programa!Tprognro
            If Not EsNulo(rs_Programa!Progarch) Then Arr_Programa(I).Progarch = rs_Programa!Progarch
            If Not EsNulo(rs_Programa!Auxint1) Then Arr_Programa(I).Auxint1 = rs_Programa!Auxint1
            If Not EsNulo(rs_Programa!Auxint2) Then Arr_Programa(I).Auxint2 = rs_Programa!Auxint2
            If Not EsNulo(rs_Programa!Auxint3) Then Arr_Programa(I).Auxint3 = rs_Programa!Auxint3
            If Not EsNulo(rs_Programa!Auxint4) Then Arr_Programa(I).Auxint4 = rs_Programa!Auxint4
            If Not EsNulo(rs_Programa!Auxint5) Then Arr_Programa(I).Auxint5 = rs_Programa!Auxint5
            If Not EsNulo(rs_Programa!Auxlog1) Then Arr_Programa(I).Auxlog1 = rs_Programa!Auxlog1
            If Not EsNulo(rs_Programa!Auxlog2) Then Arr_Programa(I).Auxlog2 = rs_Programa!Auxlog2
            If Not EsNulo(rs_Programa!Auxlog3) Then Arr_Programa(I).Auxlog3 = rs_Programa!Auxlog3
            If Not EsNulo(rs_Programa!Auxlog4) Then Arr_Programa(I).Auxlog4 = rs_Programa!Auxlog4
            If Not EsNulo(rs_Programa!Auxlog6) Then Arr_Programa(I).Auxlog5 = rs_Programa!Auxlog5
            If Not EsNulo(rs_Programa!Auxchar1) Then Arr_Programa(I).Auxchar1 = rs_Programa!Auxchar1
            If Not EsNulo(rs_Programa!Auxchar2) Then Arr_Programa(I).Auxchar2 = rs_Programa!Auxchar2
            If Not EsNulo(rs_Programa!Auxchar3) Then Arr_Programa(I).Auxchar3 = rs_Programa!Auxchar3
            If Not EsNulo(rs_Programa!Auxchar4) Then Arr_Programa(I).Auxchar4 = rs_Programa!Auxchar4
            If Not EsNulo(rs_Programa!Auxchar5) Then Arr_Programa(I).Auxchar5 = rs_Programa!Auxchar5
            If Not EsNulo(rs_Programa!Progarchest) Then Arr_Programa(I).Progarchest = rs_Programa!Progarchest
            If Not EsNulo(rs_Programa!Progcache) Then Arr_Programa(I).Progcache = rs_Programa!Progcache
            If Not EsNulo(rs_Programa!Progautor) Then Arr_Programa(I).Progautor = rs_Programa!Progautor
            If Not EsNulo(rs_Programa!Progfecmodi) Then Arr_Programa(I).Progfecmodi = rs_Programa!Progfecmodi
            If Not EsNulo(rs_Programa!Empnro) Then Arr_Programa(I).Empnro = rs_Programa!Empnro
            If Not EsNulo(rs_Programa!Auxlog6) Then Arr_Programa(I).Auxlog6 = rs_Programa!Auxlog6
            If Not EsNulo(rs_Programa!Auxlog7) Then Arr_Programa(I).Auxlog7 = rs_Programa!Auxlog7
            If Not EsNulo(rs_Programa!Auxlog8) Then Arr_Programa(I).Auxlog8 = rs_Programa!Auxlog8
            If Not EsNulo(rs_Programa!Auxlog9) Then Arr_Programa(I).Auxlog9 = rs_Programa!Auxlog9
            If Not EsNulo(rs_Programa!Auxlog10) Then Arr_Programa(I).Auxlog10 = rs_Programa!Auxlog10
            If Not EsNulo(rs_Programa!Auxlog11) Then Arr_Programa(I).Auxlog11 = rs_Programa!Auxlog11
            If Not EsNulo(rs_Programa!Auxlog12) Then Arr_Programa(I).Auxlog12 = rs_Programa!Auxlog12
            If Not EsNulo(rs_Programa!Auxchar) Then Arr_Programa(I).Auxchar = Reemplazar_SIM(rs_Programa!Auxchar)
        End If
        
        rs_Programa.MoveNext
    Loop
    
    If rs_Programa.State = adStateOpen Then rs_Programa.Close
    Set rs_Programa = Nothing
End Sub


Public Sub Cargar_Con_For_Tpa()
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga la tabla con_for_tpa. Configuracion de parametros.
' Autor      : FGZ
' Fecha      : 13/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_Con_For_Tpa As New ADODB.Recordset

    'EAM (6.67) - Se modifico el query para que traiga los conceptos del modelo de liquidación
    StrSql = "SELECT * FROM con_for_tpa "
    StrSql = StrSql & " INNER JOIN con_tp ON con_for_tpa.concnro = con_tp.concnro"
    StrSql = StrSql & " WHERE con_tp.tprocnro = " & buliq_proceso!tprocnro
    StrSql = StrSql & " Order BY con_for_tpa.concnro,fornro,tpanro,nivel,selecc"
    OpenRecordset StrSql, rs_Con_For_Tpa
            
    Max_Con_For_Tpa = rs_Con_For_Tpa.RecordCount
    ReDim Preserve Arr_con_for_tpa(Max_Con_For_Tpa) As TCon_for_tpa
            
    I = 1
    Do While Not rs_Con_For_Tpa.EOF
            Arr_con_for_tpa(I).ConcNro = rs_Con_For_Tpa!ConcNro
            Arr_con_for_tpa(I).fornro = rs_Con_For_Tpa!fornro
            Arr_con_for_tpa(I).tpanro = rs_Con_For_Tpa!tpanro
            Arr_con_for_tpa(I).Nivel = rs_Con_For_Tpa!Nivel
            Arr_con_for_tpa(I).depurable = rs_Con_For_Tpa!depurable
            Arr_con_for_tpa(I).cftauto = rs_Con_For_Tpa!cftauto
            If Not EsNulo(rs_Con_For_Tpa!Selecc) Then Arr_con_for_tpa(I).Selecc = Trim(rs_Con_For_Tpa!Selecc)
            If Not EsNulo(rs_Con_For_Tpa!Prognro) Then Arr_con_for_tpa(I).Prognro = rs_Con_For_Tpa!Prognro
        
        I = I + 1
        rs_Con_For_Tpa.MoveNext
    Loop
    
    If rs_Con_For_Tpa.State = adStateOpen Then rs_Con_For_Tpa.Close
    Set rs_Con_For_Tpa = Nothing
End Sub

Public Function Indice_Arr_con_for_tpa(ByVal Concepto As Long, ByVal Formula As Long, ByVal Parametro As Long, ByVal Nivel As Long, ByVal Selecc As String) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el indice en el arreglo de Arr_con_for_tpa.
' Autor      : FGZ
' Fecha      : 19/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    I = 1
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Arr_con_for_tpa(I).ConcNro = Concepto Then
            If Arr_con_for_tpa(I).fornro = Formula Then
                If Arr_con_for_tpa(I).tpanro = Parametro Then
                    If Arr_con_for_tpa(I).Nivel = Nivel Then
                        If Not EsNulo(Trim(Selecc)) Then
                            If Arr_con_for_tpa(I).Selecc = Trim(Selecc) Then
                                Encontro = True
                            Else
                                I = I + 1
                                If EsNulo(Arr_con_for_tpa(I).ConcNro) Then
                                    Termino = True
                                End If
                            End If
                        Else
                            Encontro = True
                        End If
                    Else
                        If Arr_con_for_tpa(I).Nivel > Nivel Then
                            Termino = True
                        Else
                            I = I + 1
                            If EsNulo(Arr_con_for_tpa(I).ConcNro) Then
                                Termino = True
                            End If
                        End If
                    End If
                Else
                    If Arr_con_for_tpa(I).tpanro > Parametro Then
                        Termino = True
                    Else
                        I = I + 1
                        If EsNulo(Arr_con_for_tpa(I).ConcNro) Then
                            Termino = True
                        End If
                    End If
                End If
            Else
                If Arr_con_for_tpa(I).fornro = Formula Then
                    Termino = True
                Else
                    I = I + 1
                    If EsNulo(Arr_con_for_tpa(I).ConcNro) Then
                        Termino = True
                    End If
                End If
            End If
        Else
            If Arr_con_for_tpa(I).ConcNro > Concepto Then
                Termino = True
            Else
                I = I + 1
                If EsNulo(Arr_con_for_tpa(I).ConcNro) Then
                    Termino = True
                End If
            End If
        End If
        
    Loop
    
    If Encontro Then
        Indice_Arr_con_for_tpa = I
    Else
        Indice_Arr_con_for_tpa = 0
    End If
End Function


Public Sub Cargar_Cge_Segun()
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga la tabla Cge_Segun. Alcence de los conceptos.
' Autor      : FGZ
' Fecha      : 19/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_Cge_Segun As New ADODB.Recordset

    'EAM (6.67) - Se modifico el query para que traiga los conceptos del modelo de liquidación
    StrSql = "SELECT Cge_Segun.ConcNro,Cge_Segun.nivel, Cge_Segun.origen, Cge_Segun.entidad FROM Cge_Segun "
    StrSql = StrSql & " INNER JOIN con_tp ON cge_segun.concnro = con_tp.concnro"
    StrSql = StrSql & " WHERE con_tp.tprocnro = " & buliq_proceso!tprocnro
    StrSql = StrSql & " Order BY Cge_Segun.concnro,nivel"
    OpenRecordset StrSql, rs_Cge_Segun
            
    Max_Cge_Segun = rs_Cge_Segun.RecordCount
    ReDim Preserve Arr_Cge_Segun(Max_Cge_Segun) As TCge_Segun
            
    I = 1
    Do While Not rs_Cge_Segun.EOF
            Arr_Cge_Segun(I).ConcNro = rs_Cge_Segun!ConcNro
            Arr_Cge_Segun(I).Nivel = rs_Cge_Segun!Nivel
            Arr_Cge_Segun(I).Origen = rs_Cge_Segun!Origen
            Arr_Cge_Segun(I).Entidad = rs_Cge_Segun!Entidad
        
        I = I + 1
        rs_Cge_Segun.MoveNext
    Loop
    
    If rs_Cge_Segun.State = adStateOpen Then rs_Cge_Segun.Close
    Set rs_Cge_Segun = Nothing
End Sub



Public Function Indice_Arr_Cge_Segun() As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el indice en el arreglo de Arr_cge_segun.
' Autor      : FGZ
' Fecha      : 19/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    I = 1
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Arr_Cge_Segun(I).ConcNro = Arr_conceptos(Concepto_Actual).ConcNro Then
            If Arr_Cge_Segun(I).Nivel = 0 Then
                If Arr_Cge_Segun(I).Origen = buliq_cabliq!Empleado Then
                    Encontro = True
                Else
                    I = I + 1
                    If I <= UBound(Arr_Cge_Segun) Then
                        If EsNulo(Arr_Cge_Segun(I).ConcNro) Then
                            Termino = True
                        End If
                    Else
                        Termino = True
                    End If
                End If
            Else
                Encontro = True
            End If
        Else
            If Arr_Cge_Segun(I).ConcNro > Arr_conceptos(Concepto_Actual).ConcNro Then
                Termino = True
            Else
                I = I + 1
                If I <= UBound(Arr_Cge_Segun) Then
                    If EsNulo(Arr_Cge_Segun(I).ConcNro) Then
                        Termino = True
                    End If
                Else
                    Termino = True
                End If
            End If
        End If
        
    Loop
    
    If Encontro Then
        Indice_Arr_Cge_Segun = I
    Else
        Indice_Arr_Cge_Segun = 0
    End If
End Function


Public Sub Cargar_Cft_Segun()
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga la tabla Cft_Segun. Alcence de los conceptos.
' Autor      : FGZ
' Fecha      : 19/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_Cft_Segun As New ADODB.Recordset

    'EAM (6.67) - Se modifico el query para que traiga los conceptos del modelo de liquidación
    StrSql = "SELECT Cft_Segun.ConcNro, fornro, tpanro, Nivel, Origen, Entidad, Selecc FROM Cft_Segun "
    StrSql = StrSql & " INNER JOIN con_tp ON Cft_Segun.concnro = con_tp.concnro"
    StrSql = StrSql & " WHERE con_tp.tprocnro = " & buliq_proceso!tprocnro
    StrSql = StrSql & " Order BY Cft_Segun.concnro,tpanro,fornro,nivel,origen"
    OpenRecordset StrSql, rs_Cft_Segun
            
    Max_Cft_Segun = rs_Cft_Segun.RecordCount
    ReDim Preserve Arr_Cft_Segun(Max_Cft_Segun) As TCft_Segun
            
    I = 1
    Do While Not rs_Cft_Segun.EOF
            Arr_Cft_Segun(I).ConcNro = rs_Cft_Segun!ConcNro
            Arr_Cft_Segun(I).fornro = rs_Cft_Segun!fornro
            Arr_Cft_Segun(I).tpanro = rs_Cft_Segun!tpanro
            If Not EsNulo(rs_Cft_Segun!Nivel) Then Arr_Cft_Segun(I).Nivel = rs_Cft_Segun!Nivel
            If Not EsNulo(rs_Cft_Segun!Origen) Then Arr_Cft_Segun(I).Origen = rs_Cft_Segun!Origen
            If Not EsNulo(rs_Cft_Segun!Entidad) Then Arr_Cft_Segun(I).Entidad = rs_Cft_Segun!Entidad
            If Not EsNulo(rs_Cft_Segun!Selecc) Then Arr_Cft_Segun(I).Selecc = rs_Cft_Segun!Selecc
        
        I = I + 1
        rs_Cft_Segun.MoveNext
    Loop
    
    If rs_Cft_Segun.State = adStateOpen Then rs_Cft_Segun.Close
    Set rs_Cft_Segun = Nothing
End Sub



Public Function Buscar_Sig_Indice_Arr_Cft_Segun(ByVal Parametro As Long, ByVal Formula As Long) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el siguiente indice al actual en el arreglo de Arr_cft_segun.
' Autor      : FGZ
' Fecha      : 19/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    I = Indice_Actual_Cft_Segun + 1
    'Controlo que no se salga de rango
    If I > Max_Cft_Segun Then
        Termino = True
    End If
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Arr_Cft_Segun(I).ConcNro = Arr_conceptos(Concepto_Actual).ConcNro Then
            If Arr_Cft_Segun(I).tpanro = Parametro Then
                If Arr_Cft_Segun(I).fornro = Formula Then
                    If Arr_Cft_Segun(I).Nivel = 0 Then
                        If Arr_Cft_Segun(I).Origen = buliq_cabliq!Empleado Then
                            Encontro = True
                        Else
                            I = I + 1
                            If I > Max_Cft_Segun Then
                                Termino = True
                            Else
                                If EsNulo(Arr_Cft_Segun(I).ConcNro) Then
                                    Termino = True
                                End If
                            End If
                        End If
                    Else
                        Encontro = True
                    End If
                Else
                    I = I + 1
                    If I > Max_Cft_Segun Then
                        Termino = True
                    Else
                        If EsNulo(Arr_Cft_Segun(I).ConcNro) Then
                            Termino = True
                        End If
                    End If
                End If
            Else
                I = I + 1
                If I > Max_Cft_Segun Then
                    Termino = True
                Else
                    If EsNulo(Arr_Cft_Segun(I).ConcNro) Then
                        Termino = True
                    End If
                End If
            End If
        Else
            If Arr_Cft_Segun(I).ConcNro > Arr_conceptos(Concepto_Actual).ConcNro Then
                Termino = True
            Else
                I = I + 1
                If I > Max_Cft_Segun Then
                    Termino = True
                Else
                    If EsNulo(Arr_Cft_Segun(I).ConcNro) Then
                        Termino = True
                    End If
                End If
            End If
        End If
        
    Loop
    
    If Encontro Then
        Buscar_Sig_Indice_Arr_Cft_Segun = I
    Else
        Buscar_Sig_Indice_Arr_Cft_Segun = 0
    End If
End Function

Public Sub Cargar_For_Tpa()
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga la tabla for_tpa. Configuracion de parametros.
' Autor      : FGZ
' Fecha      : 20/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_For_Tpa As New ADODB.Recordset

    StrSql = " SELECT DISTINCT for_tpa.fornro,for_tpa.tpanro,for_tpa.ftentrada,for_tpa.ftimprime,for_tpa.ftobligatorio,for_tpa.ftorden,for_tpa.ftinicial FROM for_tpa " & _
            " INNER JOIN con_for_tpa on for_tpa.fornro = for_tpa.fornro and con_for_tpa.tpanro = for_tpa.tpanro" & _
            " INNER JOIN con_tp on con_tp.concnro = con_for_tpa.concnro " & _
            " WHERE tprocnro= " & buliq_proceso!tprocnro & _
            " ORDER BY for_tpa.fornro, for_tpa.ftorden"
    OpenRecordset StrSql, rs_For_Tpa
            
    Max_For_Tpa = rs_For_Tpa.RecordCount
    ReDim Preserve Arr_For_Tpa(Max_For_Tpa) As TFor_tpa
    
    I = 1
    Do While Not rs_For_Tpa.EOF
            Arr_For_Tpa(I).fornro = rs_For_Tpa!fornro
            Arr_For_Tpa(I).tpanro = rs_For_Tpa!tpanro
            Arr_For_Tpa(I).ftentrada = rs_For_Tpa!ftentrada
            Arr_For_Tpa(I).ftimprime = rs_For_Tpa!ftimprime
            Arr_For_Tpa(I).ftobligatorio = rs_For_Tpa!ftobligatorio
            If Not EsNulo(rs_For_Tpa!ftorden) Then Arr_For_Tpa(I).ftorden = rs_For_Tpa!ftorden
            If Not EsNulo(rs_For_Tpa!ftinicial) Then Arr_For_Tpa(I).ftinicial = rs_For_Tpa!ftinicial
        
        I = I + 1
        rs_For_Tpa.MoveNext
    Loop
    
    If rs_For_Tpa.State = adStateOpen Then rs_For_Tpa.Close
    Set rs_For_Tpa = Nothing
End Sub



Public Sub Cargar_FunFormulas()
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga la tabla FunFormula en el recordset Global.
' Autor      : FGZ
' Fecha      : 20/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    
    StrSql = "SELECT * FROM funformula "
    OpenRecordset StrSql, rs_FunFormulas
End Sub


Public Sub Cargar_Acumuladores()
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga la tabla de Acumuladores en el arreglo Global.
' Autor      : FGZ
' Fecha      : 20/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_acumulador As New ADODB.Recordset

    StrSql = "SELECT * FROM acumulador "
    StrSql = StrSql & " ORDER BY acunro "
    OpenRecordset StrSql, rs_acumulador
    
    If Not rs_acumulador.EOF Then
        rs_acumulador.MoveLast
        Max_Acumuladores = rs_acumulador!acuNro + 1
        
        ReDim Preserve Arr_Acumulador(Max_Acumuladores) As TAcumulador
        rs_acumulador.MoveFirst
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 2) & "No se encontraron Acumuladores"
        End If
    End If
    
    Do While Not rs_acumulador.EOF
            I = rs_acumulador!acuNro
            Arr_Acumulador(I).acuNro = rs_acumulador!acuNro
        
            If Not EsNulo(rs_acumulador!acudesabr) Then
                Arr_Acumulador(I).acudesabr = rs_acumulador!acudesabr
            End If
            
            If Not EsNulo(rs_acumulador!acusist) Then
                Arr_Acumulador(I).acusist = rs_acumulador!acusist
            End If
            
            If Not EsNulo(rs_acumulador!acudesext) Then
                Arr_Acumulador(I).acudesext = rs_acumulador!acudesext
            End If
            
            If Not EsNulo(rs_acumulador!acumes) Then
                Arr_Acumulador(I).acumes = rs_acumulador!acumes
            End If
            
            If Not EsNulo(rs_acumulador!acutopea) Then
                Arr_Acumulador(I).acutopea = rs_acumulador!acutopea
            End If
            
            If Not EsNulo(rs_acumulador!acudesborde) Then
                Arr_Acumulador(I).acudesborde = rs_acumulador!acudesborde
            End If
                
            If Not EsNulo(rs_acumulador!acurecalculo) Then
                Arr_Acumulador(I).acurecalculo = rs_acumulador!acurecalculo
            End If
                
            If Not EsNulo(rs_acumulador!acuimponible) Then
                Arr_Acumulador(I).acuimponible = rs_acumulador!acuimponible
            End If
                
            If Not EsNulo(rs_acumulador!acuimpcont) Then
                Arr_Acumulador(I).acuimpcont = rs_acumulador!acuimpcont
            End If
                
            If Not EsNulo(rs_acumulador!acusel1) Then
                Arr_Acumulador(I).acusel1 = rs_acumulador!acusel1
            End If
                
            If Not EsNulo(rs_acumulador!acusel2) Then
                Arr_Acumulador(I).acusel2 = rs_acumulador!acusel2
            End If
                
            If Not EsNulo(rs_acumulador!acusel3) Then
                Arr_Acumulador(I).acusel3 = rs_acumulador!acusel3
            End If
                
            If Not EsNulo(rs_acumulador!acuppag) Then
                Arr_Acumulador(I).acuppag = rs_acumulador!acuppag
            End If
                
            If Not EsNulo(rs_acumulador!acudepu) Then
                Arr_Acumulador(I).acudepu = rs_acumulador!acudepu
            End If
                
            If Not EsNulo(rs_acumulador!acuhist) Then
                Arr_Acumulador(I).acuhist = rs_acumulador!acuhist
            End If
                
            If Not EsNulo(rs_acumulador!acumanual) Then
                Arr_Acumulador(I).acumanual = rs_acumulador!acumanual
            End If
                
            If Not EsNulo(rs_acumulador!acuimpri) Then
                Arr_Acumulador(I).acuimpri = rs_acumulador!acuimpri
            End If
                
            If Not EsNulo(rs_acumulador!tacunro) Then
                Arr_Acumulador(I).tacunro = rs_acumulador!tacunro
            End If
            
            If Not EsNulo(rs_acumulador!Empnro) Then
                Arr_Acumulador(I).Empnro = rs_acumulador!Empnro
            End If
                
            If Not EsNulo(rs_acumulador!acuorden) Then
                Arr_Acumulador(I).acuorden = rs_acumulador!acuorden
            End If
            
            If Not EsNulo(rs_acumulador!acuretro) Then
                Arr_Acumulador(I).acuretro = rs_acumulador!acuretro
            End If
                
            If Not EsNulo(rs_acumulador!acunoneg) Then
                Arr_Acumulador(I).acunoneg = rs_acumulador!acunoneg
            End If
        rs_acumulador.MoveNext
    Loop
    
    If rs_acumulador.State = adStateOpen Then rs_acumulador.Close
    Set rs_acumulador = Nothing
End Sub


Public Sub Cargar_Acumuladores_Log_Detallado()
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga la tabla de Acumuladores en el arreglo Global.
' Autor      : FGZ
' Fecha      : 20/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim rs_acumulador As New ADODB.Recordset

    StrSql = "SELECT * FROM acumulador "
    StrSql = StrSql & " ORDER BY acunro "
    OpenRecordset StrSql, rs_acumulador
    
    If Not rs_acumulador.EOF Then
        rs_acumulador.MoveLast
        Max_Acumuladores = rs_acumulador!acuNro + 1
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 2) & "Ultimo acumulador " & Max_Acumuladores
        End If
        
        ReDim Preserve Arr_Acumulador(Max_Acumuladores) As TAcumulador
        rs_acumulador.MoveFirst
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 2) & "No se encontraron Acumuladores"
        End If
    End If
    
    Do While Not rs_acumulador.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 2) & "Acumulador " & rs_acumulador!acuNro
            End If
    
            I = rs_acumulador!acuNro
            Arr_Acumulador(I).acuNro = rs_acumulador!acuNro
        
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acudesabr"
            End If
            If Not EsNulo(rs_acumulador!acudesabr) Then
                Arr_Acumulador(I).acudesabr = rs_acumulador!acudesabr
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acusist"
            End If
            If Not EsNulo(rs_acumulador!acusist) Then
                Arr_Acumulador(I).acusist = rs_acumulador!acusist
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acudesext"
            End If
            If Not EsNulo(rs_acumulador!acudesext) Then
                Arr_Acumulador(I).acudesext = rs_acumulador!acudesext
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acumes"
            End If
            If Not EsNulo(rs_acumulador!acumes) Then
                Arr_Acumulador(I).acumes = rs_acumulador!acumes
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acutopea"
            End If
            If Not EsNulo(rs_acumulador!acutopea) Then
                Arr_Acumulador(I).acutopea = rs_acumulador!acutopea
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acudesborde"
            End If
            If Not EsNulo(rs_acumulador!acudesborde) Then
                Arr_Acumulador(I).acudesborde = rs_acumulador!acudesborde
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acurecalculo"
            End If
            If Not EsNulo(rs_acumulador!acurecalculo) Then
                Arr_Acumulador(I).acurecalculo = rs_acumulador!acurecalculo
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuimponible"
            End If
            If Not EsNulo(rs_acumulador!acuimponible) Then
                Arr_Acumulador(I).acuimponible = rs_acumulador!acuimponible
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuimpcont"
            End If
            If Not EsNulo(rs_acumulador!acuimpcont) Then
                Arr_Acumulador(I).acuimpcont = rs_acumulador!acuimpcont
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acusel1"
            End If
            If Not EsNulo(rs_acumulador!acusel1) Then
                Arr_Acumulador(I).acusel1 = rs_acumulador!acusel1
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acusel2"
            End If
            If Not EsNulo(rs_acumulador!acusel2) Then
                Arr_Acumulador(I).acusel2 = rs_acumulador!acusel2
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acusel3"
            End If
            If Not EsNulo(rs_acumulador!acusel3) Then
                Arr_Acumulador(I).acusel3 = rs_acumulador!acusel3
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuppag"
            End If
            If Not EsNulo(rs_acumulador!acuppag) Then
                Arr_Acumulador(I).acuppag = rs_acumulador!acuppag
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acudepu"
            End If
            If Not EsNulo(rs_acumulador!acudepu) Then
                Arr_Acumulador(I).acudepu = rs_acumulador!acudepu
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuhist"
            End If
            If Not EsNulo(rs_acumulador!acuhist) Then
                Arr_Acumulador(I).acuhist = rs_acumulador!acuhist
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acumanual"
            End If
            If Not EsNulo(rs_acumulador!acumanual) Then
                Arr_Acumulador(I).acumanual = rs_acumulador!acumanual
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuimpri"
            End If
            If Not EsNulo(rs_acumulador!acuimpri) Then
                Arr_Acumulador(I).acuimpri = rs_acumulador!acuimpri
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "tacunro"
            End If
            If Not EsNulo(rs_acumulador!tacunro) Then
                Arr_Acumulador(I).tacunro = rs_acumulador!tacunro
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Empnro"
            End If
            If Not EsNulo(rs_acumulador!Empnro) Then
                Arr_Acumulador(I).Empnro = rs_acumulador!Empnro
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuorden"
            End If
            If Not EsNulo(rs_acumulador!acuorden) Then
                Arr_Acumulador(I).acuorden = rs_acumulador!acuorden
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuretro"
            End If
            If Not EsNulo(rs_acumulador!acuretro) Then
                Arr_Acumulador(I).acuretro = rs_acumulador!acuretro
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acunoneg"
            End If
            If Not EsNulo(rs_acumulador!acunoneg) Then
                Arr_Acumulador(I).acunoneg = rs_acumulador!acunoneg
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
                Arr_Acumulador(I).acunoneg = 0
            End If
        rs_acumulador.MoveNext
    Loop
    
    If rs_acumulador.State = adStateOpen Then rs_acumulador.Close
    Set rs_acumulador = Nothing
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 2) & "Fin Carga de Acumuladores"
    End If
    
End Sub


Public Function Siguiente_Acumulador() As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el siguiente indice del arreglo de acumuladores Arr_Acumulador.
' Autor      : FGZ
' Fecha      : 20/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    I = Acumulador_Actual + 1
    'Controlo que no se salga de rango
    If I > Max_Acumuladores Then
        Termino = True
    End If
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Not EsNulo(Arr_Acumulador(I).acuNro) Then
            Encontro = True
        Else
            I = I + 1
            If I > Max_Acumuladores Then
                Termino = True
            End If
        End If
    Loop
    
    If Encontro Then
        Siguiente_Acumulador = I
    Else
        Siguiente_Acumulador = 0
    End If
End Function


Public Sub Cargar_Con_Acum(ByVal Nrotipo As Long, ByVal Fecha_Inicio As Date, ByVal Fecha_Fin As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga todos los acumuladores a los cuales suma cada conceptos a liquidar.
' Autor      : FGZ
' Fecha      : 23/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim Ultimo_Concepto As Long
Dim rs_Con_Acum As New ADODB.Recordset


    StrSql = "SELECT * FROM con_acum "
    StrSql = StrSql & " INNER JOIN concepto ON con_acum.concnro = concepto.concnro "
    StrSql = StrSql & " INNER JOIN con_tp ON con_tp.concnro = concepto.concnro "
    StrSql = StrSql & " WHERE con_tp.tprocnro = " & Nrotipo
    StrSql = StrSql & " AND (concepto.concvalid = 0 or ( concdesde <= " & ConvFecha(Fecha_Inicio)
    StrSql = StrSql & " AND conchasta >= " & ConvFecha(Fecha_Fin) & "))"
    StrSql = StrSql & " ORDER BY concepto.tconnro, concepto.concorden, acunro "
    OpenRecordset StrSql, rs_Con_Acum

    ReDim Preserve Arr_Con_Acum(Max_Conceptos, Max_Acumuladores) As Long

    I = 0
    Ultimo_Concepto = 0
    Do While Not rs_Con_Acum.EOF
        If Ultimo_Concepto <> rs_Con_Acum!ConcNro Then
            I = Buscar_Indice_Concepto(rs_Con_Acum!ConcNro, I)
        End If
        Arr_Con_Acum(I, rs_Con_Acum!acuNro) = -1
        
        Ultimo_Concepto = rs_Con_Acum!ConcNro
        
        rs_Con_Acum.MoveNext
    Loop

    If rs_Con_Acum.State = adStateOpen Then rs_Con_Acum.Close
    Set rs_Con_Acum = Nothing
End Sub


Public Function Buscar_Indice_Concepto(ByVal Concepto As Long, ByVal Posicion As Integer) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el indice del concepto dentro del arreglo de conceptos Arr_concepto.
' Autor      : FGZ
' Fecha      : 23/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    I = Posicion
    'Controlo que no se salga de rango
    If I > Max_Conceptos Then
        Termino = True
    End If
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Arr_conceptos(I).ConcNro = Concepto Then
            Encontro = True
        Else
            I = I + 1
            If I > Max_Conceptos Then
                Termino = True
            End If
        End If
    Loop
    
    If Encontro Then
        Buscar_Indice_Concepto = I
    Else
        Buscar_Indice_Concepto = 0
    End If
End Function


Public Function Siguiente_Con_Acum(ByVal Posicion As Long) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el siguiente acumulador para el concepto en con_acum.
' Autor      : FGZ
' Fecha      : 23/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim j As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    j = Posicion + 1
    'Controlo que no se salga de rango
    If j > Max_Acumuladores Then
        Termino = True
    Else
        Termino = False
    End If
    
    Encontro = False
    'Termino = False
    Do While Not Encontro And Not Termino
        If CBool(Arr_Con_Acum(Concepto_Actual, j)) Then
            Encontro = True
        Else
            j = j + 1
            If j > Max_Acumuladores Then
                Termino = True
            End If
        End If
    Loop
    
    If Encontro Then
        Siguiente_Con_Acum = j
    Else
        Siguiente_Con_Acum = 0
    End If

End Function

Public Sub Carga_Tablas_SIM()
' ---------------------------------------------------------------------------------------------
' Descripcion: Crea el array con las tablas que se usan en la simulacion
' Autor      : Diego Rosso
' Fecha      : 25/08/2008
' ---------------------------------------------------------------------------------------------

Arr_Tablas_SIM(0) = "acu_liq"
Arr_Tablas_SIM(1) = "acu_mes"
Arr_Tablas_SIM(2) = "cabliq"
Arr_Tablas_SIM(3) = "desmen"
Arr_Tablas_SIM(4) = "detliq"
Arr_Tablas_SIM(5) = "embargo"
Arr_Tablas_SIM(6) = "embcuota"
Arr_Tablas_SIM(7) = "emp_lic"
Arr_Tablas_SIM(8) = "emp_ticket"
Arr_Tablas_SIM(9) = "emp_tikdist"
Arr_Tablas_SIM(10) = "fases"
Arr_Tablas_SIM(11) = "gti_acunov"
Arr_Tablas_SIM(12) = "his_estructura"
Arr_Tablas_SIM(13) = "impgralarg"
Arr_Tablas_SIM(14) = "impmesarg"
Arr_Tablas_SIM(15) = "impproarg"
Arr_Tablas_SIM(16) = "novaju"
Arr_Tablas_SIM(17) = "vales"
Arr_Tablas_SIM(18) = "novemp"
Arr_Tablas_SIM(19) = "prestamo"
Arr_Tablas_SIM(20) = "pre_cuota"
Arr_Tablas_SIM(21) = "proceso"
Arr_Tablas_SIM(22) = "traza"
Arr_Tablas_SIM(23) = "traza_gan"
Arr_Tablas_SIM(24) = "traza_gan_item_top"
Arr_Tablas_SIM(25) = "vacpagdesc"
'Saco la tabla empleado porq existe tambien el campo empleado y no hay forma de distinguirlos.
'Arr_Tablas_SIM(26) = "empleado"
'Arr_Tablas_SIM(27) = ""



End Sub

Public Function Reemplazar_SIM(ByVal Cadena As String) As String
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca en la cadena de texto (sentencia SQL) tablas que deben ser cambiadas por las
'              de simulacion y las reemplaza, agregandole el SIM_ adelante.
'              Se usa para las busquedas Internas
' Autor      : Diego Rosso
' Fecha      : 25/08/2008
' ---------------------------------------------------------------------------------------------
Dim j As Integer
'Buscar las tablas de simulacion y si hay ponerle el sim
 For j = 0 To UBound(Arr_Tablas_SIM)
    Cadena = Replace(Cadena, Arr_Tablas_SIM(j), "SIM_" & Arr_Tablas_SIM(j))
    If j = 7 Then
        Cadena = Replace(Cadena, "SIM_emp_licnro", "emp_licnro")
    End If
    
 Next j

Reemplazar_SIM = Cadena
End Function


Public Sub InicializarWF_Tpa(ByVal Concepto As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Inicializa los parametros de cada formula.
' Autor      : FGZ
' Fecha      : 24/05/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Long
'Dim LI As Long
'Dim LS As Long


    'Busco el primer y ultimo parametro de la formula del concepto actual
    LI_WF_Tpa = BuscarPrimer_For_Tpa(Arr_conceptos(Concepto_Actual).fornro)
    LS_WF_Tpa = BuscarUltimo_For_Tpa(Arr_conceptos(Concepto_Actual).fornro)
    Max_WF_Tpa = LS_WF_Tpa + 1
    ReDim Preserve Arr_WF_TPA(Max_WF_Tpa) As TWF_Tpa
    'ReDim Preserve Arr_WF_TPA(LS + 1) As TWF_Tpa
    'ReDim Preserve Arr_WF_TPA(LI To LS) As TWF_Tpa
    
    'deberia inicializar todos los parametros del concepto hasta que cambie concepto
    For I = LI_WF_Tpa To LS_WF_Tpa
            'Se reutilizan durante el calculo de cada parametro de cada formula de cada concepto
            Arr_WF_TPA(I).tipoparam = 0
            Arr_WF_TPA(I).ftorden = 0
            Arr_WF_TPA(I).nombre = "" 'Arr_conceptos(Concepto_Actual).concabr
            Arr_WF_TPA(I).Valor = 0
            Arr_WF_TPA(I).Fecha = ""
            'FGZ - 13/07/2011 ---- le agregué este campo para controlar si el parametro ya fué calculado
            Arr_WF_TPA(I).Calculado = False
    Next I

End Sub

Public Function BuscarPrimer_For_Tpa(ByVal Concepto As Long) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el primer indice al actual en el arreglo de Arr_For_Tpa.
' Autor      : FGZ
' Fecha      : 24/05/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    'I = Indice_Actual_For_Tpa + 1
    I = Indice_Actual_For_Tpa
    
    'Controlo que no se salga de rango
    Encontro = False
    Termino = False
    If I > Max_For_Tpa Then
        Termino = True
    End If
    Do While Not Encontro And Not Termino
        If Arr_For_Tpa(I).fornro = Arr_conceptos(Concepto_Actual).fornro Then
            Encontro = True
        Else
            If Arr_For_Tpa(I).fornro > Arr_conceptos(Concepto_Actual).fornro Then
                Termino = True
            Else
                I = I + 1
                If I > Max_For_Tpa Then
                    Termino = True
                Else
                    If EsNulo(Arr_For_Tpa(I).fornro) Then
                        Termino = True
                    End If
                End If
            End If
        End If
    Loop
    
    If Encontro Then
        BuscarPrimer_For_Tpa = I
    Else
        BuscarPrimer_For_Tpa = 0
    End If

End Function

Public Function BuscarUltimo_For_Tpa(ByVal Concepto As Long) As Long
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el ultimo indice al actual en el arreglo de Arr_For_Tpa.
' Autor      : FGZ
' Fecha      : 24/05/2011
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim EncontroP As Boolean
Dim EncontroU As Boolean
Dim Termino As Boolean
    
    'I = Indice_Actual_For_Tpa + 1
    I = Indice_Actual_For_Tpa
    
    'Controlo que no se salga de rango
    EncontroP = False
    EncontroU = False
    Termino = False
    If I > Max_For_Tpa Then
        Termino = True
    End If
    Do While Not EncontroU And Not Termino
        If Not EncontroP Then
            If Arr_For_Tpa(I).fornro = Arr_conceptos(Concepto_Actual).fornro Then
                EncontroP = True
            Else
                If Arr_For_Tpa(I).fornro > Arr_conceptos(Concepto_Actual).fornro Then
                    Termino = True
                Else
                    I = I + 1
                    If I > Max_For_Tpa Then
                        Termino = True
                    Else
                        If EsNulo(Arr_For_Tpa(I).fornro) Then
                            Termino = True
                        End If
                    End If
                End If
            End If
        Else    'ahora busco el ultimo
            If Arr_For_Tpa(I).fornro <> Arr_conceptos(Concepto_Actual).fornro Then
                EncontroU = True
            Else
                I = I + 1
                If I > Max_For_Tpa Then
                    Termino = True
                Else
                    If EsNulo(Arr_For_Tpa(I).fornro) Then
                        Termino = True
                    End If
                End If
            End If
        End If
    Loop
    
    If EncontroU Then
        BuscarUltimo_For_Tpa = I - 1
    Else
        If EncontroP Then
            BuscarUltimo_For_Tpa = Max_For_Tpa
        Else
            BuscarUltimo_For_Tpa = 0
        End If
    End If
End Function



Public Function BuscarSiguiente_For_Tpa(ByVal Concepto As Long) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el siguiente indice al actual en el arreglo de Arr_For_Tpa.
' Autor      : FGZ
' Fecha      : 19/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim I As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    I = Indice_Actual_For_Tpa + 1
    'Controlo que no se salga de rango
    Encontro = False
    Termino = False
    If I > Max_For_Tpa Then
        Termino = True
    End If
    Do While Not Encontro And Not Termino
        If Arr_For_Tpa(I).fornro = Arr_conceptos(Concepto_Actual).fornro Then
            Encontro = True
        Else
            If Arr_For_Tpa(I).fornro > Arr_conceptos(Concepto_Actual).fornro Then
                Termino = True
            Else
                I = I + 1
                If I > Max_For_Tpa Then
                    Termino = True
                Else
                    If EsNulo(Arr_For_Tpa(I).fornro) Then
                        Termino = True
                    End If
                End If
            End If
        End If
    Loop
    
    If Encontro Then
        BuscarSiguiente_For_Tpa = I
    Else
        BuscarSiguiente_For_Tpa = 0
    End If
End Function

