Attribute VB_Name = "MdlBuliq"
Option Explicit



Public Sub Establecer_Buliq_concepto(ByVal nroconcepto As Long, ByRef Ok As Boolean)
    StrSql = "SELECT * FROM concepto WHERE concnro = " & nroconcepto
    OpenRecordset StrSql, rs_Buliq_Concepto

    Ok = Not rs_Buliq_Concepto.EOF
End Sub


Public Sub Establecer_Empleado(ByVal p_ternro As Long, ByVal p_grunro As Long, ByVal p_cliqnro As Long, ByVal p_fecha_inicio As Date, ByVal p_fecha_fin As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Setea las variable globales del EMPLEADO con los valores pasados por parametros.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Fases As New ADODB.Recordset

    NroEmple = p_ternro
    NroGrupo = p_grunro
    NroCab = p_cliqnro

    ' Actualizar los buffer's Auxiliares
    ' Empleado
    StrSql = "SELECT * FROM empleado WHERE ternro = " & CStr(p_ternro)
    OpenRecordset StrSql, buliq_empleado

    If buliq_empleado.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 1) & "no levanto ningun empleado de empleado"
        End If
    Else
        Legajo = buliq_empleado!empleg
    End If
    ' FGZ - 18/03/2004
    ' Si el empleado no esta activo ==> seteo la fecha de baja
    If Not CBool(buliq_empleado!empest) Then
        StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro & _
                " AND ((altfec >= " & ConvFecha(fecha_inicio) & " AND altfec <= " & ConvFecha(fecha_fin) & ") " & _
                " OR (bajfec <= " & ConvFecha(fecha_fin) & "))" & _
                " ORDER BY altfec"
        OpenRecordset StrSql, rs_Fases
        If Not rs_Fases.EOF Then rs_Fases.MoveLast
        If Not rs_Fases.EOF Then
            If Not CBool(rs_Fases!estado) Then
                Empleado_Fecha_Fin = rs_Fases!bajfec
                
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
    StrSql = "SELECT * FROM cabliq WHERE cliqnro = " & CStr(p_cliqnro)
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
    StrSql = "SELECT * FROM proceso WHERE pronro = " & CStr(p_pronro)
    OpenRecordset StrSql, buliq_proceso

    'FGZ - 18/03/2004
    Empleado_Fecha_Inicio = buliq_proceso!profecini
    Empleado_Fecha_Fin = buliq_proceso!profecfin
    'FGZ - 18/03/2004
    
    ' Periodo
    StrSql = "SELECT * FROM periodo WHERE pliqnro = " & CStr(buliq_proceso!PliqNro)
    OpenRecordset StrSql, buliq_periodo

    ' impgralarg
    StrSql = "SELECT * FROM impgralarg WHERE pronro = " & CStr(p_pronro)
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

    StrSql = "SELECT * FROM impgralarg WHERE pronro = " & CStr(pronro) & _
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

StrSql = "SELECT emp.empnro FROM his_estructura "
StrSql = StrSql & " INNER JOIN estructura emp ON emp.estrnro = his_estructura.estrnro "
StrSql = StrSql & " WHERE (his_estructura.tenro = 10 and his_estructura.ternro =" & buliq_empleado!ternro & ")"
StrSql = StrSql & " AND his_estructura.htetdesde <=" & ConvFecha(FHasta)
StrSql = StrSql & " AND (his_estructura.htethasta >= " & ConvFecha(FDesde)
StrSql = StrSql & " OR his_estructura.htethasta IS NULL)"
OpenRecordset StrSql, rs_Empresa
If Not rs_Empresa.EOF Then
    NroEmp = rs_Empresa!Empnro
Else
    Flog.writeline "El empleado " & buliq_empleado!empleg & " no tiene empresa asignada"
    NroEmp = 0
End If
If rs_Empresa.State = adStateOpen Then rs_Empresa.Close
Set rs_Empresa = Nothing

End Sub

Public Sub Establecer_Empresa_old(ByVal p_empnro As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: setea la variable global con el nro de empresa.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------


NroEmp = p_empnro

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

    
    FirmaActiva5 = False
    FirmaActiva15 = False
    FirmaActiva19 = False
    FirmaActiva20 = False
    
    StrSql = "select * from cystipo where (cystipnro = 5 or cystipnro = 15 OR cystipnro = 19 or cystipnro = 20) AND cystipact = -1"
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
    Case Else
    End Select
        
        rs_cystipo.MoveNext
    Loop
    
If rs_cystipo.State = adStateOpen Then rs_cystipo.Close
Set rs_cystipo = Nothing

End Sub


Public Sub CargarConceptos(ByVal Nrotipo As Long, ByVal fecha_inicio As Date, ByVal fecha_fin As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga todos los conceptos a liquidar.
' Autor      : FGZ
' Fecha      : 11/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim rs_Conceptos As New ADODB.Recordset

    'CONCEPTO:
    StrSql = "SELECT * FROM concepto " & _
             " INNER JOIN con_tp ON con_tp.concnro = concepto.concnro " & _
             " INNER JOIN formula ON formula.fornro = concepto.fornro " & _
             " WHERE con_tp.tprocnro = " & Nrotipo & _
             " AND (concepto.concvalid = 0 or ( concdesde <= " & ConvFecha(fecha_inicio) & _
             " AND conchasta >= " & ConvFecha(fecha_fin) & "))" & _
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
    
    i = 1
    Do While Not rs_Conceptos.EOF
        If Not EsNulo(rs_Conceptos!concnro) Then Arr_conceptos(i).concnro = rs_Conceptos!concnro
        If Not EsNulo(rs_Conceptos!Conccod) Then Arr_conceptos(i).Conccod = rs_Conceptos!Conccod
        If Not EsNulo(rs_Conceptos!Concajuste) Then Arr_conceptos(i).Concajuste = rs_Conceptos!Concajuste
        If Not EsNulo(rs_Conceptos!Conccantdec) Then Arr_conceptos(i).Conccantdec = rs_Conceptos!Conccantdec
        If Not EsNulo(rs_Conceptos!concabr) Then Arr_conceptos(i).concabr = rs_Conceptos!concabr
        If Not EsNulo(rs_Conceptos!concretro) Then Arr_conceptos(i).concretro = rs_Conceptos!concretro
        If Not EsNulo(rs_Conceptos!tconnro) Then Arr_conceptos(i).tconnro = rs_Conceptos!tconnro
        If Not EsNulo(rs_Conceptos!Conctexto) Then Arr_conceptos(i).Conctexto = rs_Conceptos!Conctexto
        If Not EsNulo(rs_Conceptos!fornro) Then Arr_conceptos(i).fornro = rs_Conceptos!fornro
        If Not EsNulo(rs_Conceptos!Fortipo) Then Arr_conceptos(i).Fortipo = rs_Conceptos!Fortipo
        If Not EsNulo(rs_Conceptos!Forexpresion) Then Arr_conceptos(i).Forexpresion = rs_Conceptos!Forexpresion
        If Not EsNulo(rs_Conceptos!Fordabr) Then Arr_conceptos(i).Fordabr = rs_Conceptos!Fordabr
        If Not EsNulo(rs_Conceptos!Forprog) Then Arr_conceptos(i).Forprog = rs_Conceptos!Forprog
        
        'buliq_concepto
        If Not EsNulo(rs_Conceptos!concnro) Then Buliq_Concepto(i).concnro = rs_Conceptos!concnro
        If Not EsNulo(rs_Conceptos!Conccod) Then Buliq_Concepto(i).Conccod = rs_Conceptos!Conccod
        If Not EsNulo(rs_Conceptos!Concajuste) Then Buliq_Concepto(i).Concajuste = rs_Conceptos!Concajuste
        If Not EsNulo(rs_Conceptos!Conccantdec) Then Buliq_Concepto(i).Conccantdec = rs_Conceptos!Conccantdec
        If Not EsNulo(rs_Conceptos!concabr) Then Buliq_Concepto(i).concabr = rs_Conceptos!concabr
        If Not EsNulo(rs_Conceptos!concretro) Then Buliq_Concepto(i).concretro = rs_Conceptos!concretro
        If Not EsNulo(rs_Conceptos!tconnro) Then Buliq_Concepto(i).tconnro = rs_Conceptos!tconnro
        If Not EsNulo(rs_Conceptos!Conctexto) Then Buliq_Concepto(i).Conctexto = rs_Conceptos!Conctexto
        If Not EsNulo(rs_Conceptos!concorden) Then Buliq_Concepto(i).concorden = rs_Conceptos!concorden
        If Not EsNulo(rs_Conceptos!concext) Then Buliq_Concepto(i).concext = rs_Conceptos!concext
        If Not EsNulo(rs_Conceptos!concvalid) Then Buliq_Concepto(i).concvalid = rs_Conceptos!concvalid
        If Not EsNulo(rs_Conceptos!concdesde) Then Buliq_Concepto(i).concdesde = rs_Conceptos!concdesde
        If Not EsNulo(rs_Conceptos!conchasta) Then Buliq_Concepto(i).conchasta = rs_Conceptos!conchasta
        If Not EsNulo(rs_Conceptos!concrepet) Then Buliq_Concepto(i).concrepet = rs_Conceptos!concrepet
        If Not EsNulo(rs_Conceptos!concniv) Then Buliq_Concepto(i).concniv = rs_Conceptos!concniv
        If Not EsNulo(rs_Conceptos!fornro) Then Buliq_Concepto(i).fornro = rs_Conceptos!fornro
        If Not EsNulo(rs_Conceptos!concimp) Then Buliq_Concepto(i).concimp = rs_Conceptos!concimp
        If Not EsNulo(rs_Conceptos!codseguridad) Then Buliq_Concepto(i).codseguridad = rs_Conceptos!codseguridad
        If Not EsNulo(rs_Conceptos!concusado) Then Buliq_Concepto(i).concusado = rs_Conceptos!concusado
        If Not EsNulo(rs_Conceptos!concpuente) Then Buliq_Concepto(i).concpuente = rs_Conceptos!concpuente
        If Not EsNulo(rs_Conceptos!Empnro) Then Buliq_Concepto(i).Empnro = rs_Conceptos!Empnro
        If Not EsNulo(rs_Conceptos!concautor) Then Buliq_Concepto(i).concautor = rs_Conceptos!concautor
        If Not EsNulo(rs_Conceptos!concfecmodi) Then Buliq_Concepto(i).concfecmodi = rs_Conceptos!concfecmodi
        If Not EsNulo(rs_Conceptos!Concajuste) Then Buliq_Concepto(i).Concajuste = rs_Conceptos!Concajuste

        rs_Conceptos.MoveNext
        i = i + 1
    Loop
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 2) & "I =  " & i
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
Dim i As Integer
Dim rs_Empleados As New ADODB.Recordset

    ' Liquido los empleados
    If Todos Then
        StrSql = "SELECT * FROM cabliq WHERE pronro =" & NroProc
    Else
        StrSql = "SELECT * FROM cabliq " & _
        " INNER JOIN batch_empleado ON batch_empleado.ternro = cabliq.empleado " & _
        " WHERE cabliq.pronro =" & NroProc & _
        " AND batch_empleado.bpronro = " & Bpronro
    End If
    OpenRecordset StrSql, rs_Empleados

    Max_Cabeceras = rs_Empleados.RecordCount
    ReDim Preserve Arr_EmpCab(Max_Cabeceras) As TEmpCabLiq
    
    i = 1
    Do While Not rs_Empleados.EOF
        Arr_EmpCab(i).cliqnro = rs_Empleados!cliqnro
        Arr_EmpCab(i).Empleado = rs_Empleados!Empleado
        Arr_EmpCab(i).ternro = rs_Empleados!ternro
        
        i = i + 1
        rs_Empleados.MoveNext
        
    Loop
    
    Cantidad_de_Empleados = i - 1
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
Dim i As Long
Dim rs_Programa As New ADODB.Recordset

    StrSql = "SELECT * FROM programa "
    StrSql = StrSql & " ORDER BY prognro "
    OpenRecordset StrSql, rs_Programa

    If Not rs_Programa.EOF Then
        rs_Programa.MoveLast
        
        Max_Programas = rs_Programa!Prognro + 1
        ReDim Preserve Arr_Programa(Max_Programas) As TPrograma
        rs_Programa.MoveFirst
    End If
    
    Do While Not rs_Programa.EOF
        If Not EsNulo(rs_Programa!Prognro) Then
            i = rs_Programa!Prognro
            Arr_Programa(i).Prognro = rs_Programa!Prognro
            If Not EsNulo(rs_Programa!Prognom) Then Arr_Programa(i).Prognom = rs_Programa!Prognom
            If Not EsNulo(rs_Programa!Progdesc) Then Arr_Programa(i).Progdesc = rs_Programa!Progdesc
            If Not EsNulo(rs_Programa!Tprognro) Then Arr_Programa(i).Tprognro = rs_Programa!Tprognro
            If Not EsNulo(rs_Programa!Progarch) Then Arr_Programa(i).Progarch = rs_Programa!Progarch
            If Not EsNulo(rs_Programa!Auxint1) Then Arr_Programa(i).Auxint1 = rs_Programa!Auxint1
            If Not EsNulo(rs_Programa!Auxint2) Then Arr_Programa(i).Auxint2 = rs_Programa!Auxint2
            If Not EsNulo(rs_Programa!Auxint3) Then Arr_Programa(i).Auxint3 = rs_Programa!Auxint3
            If Not EsNulo(rs_Programa!Auxint4) Then Arr_Programa(i).Auxint4 = rs_Programa!Auxint4
            If Not EsNulo(rs_Programa!Auxint5) Then Arr_Programa(i).Auxint5 = rs_Programa!Auxint5
            If Not EsNulo(rs_Programa!Auxlog1) Then Arr_Programa(i).Auxlog1 = rs_Programa!Auxlog1
            If Not EsNulo(rs_Programa!Auxlog2) Then Arr_Programa(i).Auxlog2 = rs_Programa!Auxlog2
            If Not EsNulo(rs_Programa!Auxlog3) Then Arr_Programa(i).Auxlog3 = rs_Programa!Auxlog3
            If Not EsNulo(rs_Programa!Auxlog4) Then Arr_Programa(i).Auxlog4 = rs_Programa!Auxlog4
            If Not EsNulo(rs_Programa!Auxlog6) Then Arr_Programa(i).Auxlog5 = rs_Programa!Auxlog5
            If Not EsNulo(rs_Programa!Auxchar1) Then Arr_Programa(i).Auxchar1 = rs_Programa!Auxchar1
            If Not EsNulo(rs_Programa!Auxchar2) Then Arr_Programa(i).Auxchar2 = rs_Programa!Auxchar2
            If Not EsNulo(rs_Programa!Auxchar3) Then Arr_Programa(i).Auxchar3 = rs_Programa!Auxchar3
            If Not EsNulo(rs_Programa!Auxchar4) Then Arr_Programa(i).Auxchar4 = rs_Programa!Auxchar4
            If Not EsNulo(rs_Programa!Auxchar5) Then Arr_Programa(i).Auxchar5 = rs_Programa!Auxchar5
            If Not EsNulo(rs_Programa!Progarchest) Then Arr_Programa(i).Progarchest = rs_Programa!Progarchest
            If Not EsNulo(rs_Programa!Progcache) Then Arr_Programa(i).Progcache = rs_Programa!Progcache
            If Not EsNulo(rs_Programa!Progautor) Then Arr_Programa(i).Progautor = rs_Programa!Progautor
            If Not EsNulo(rs_Programa!Progfecmodi) Then Arr_Programa(i).Progfecmodi = rs_Programa!Progfecmodi
            If Not EsNulo(rs_Programa!Empnro) Then Arr_Programa(i).Empnro = rs_Programa!Empnro
            If Not EsNulo(rs_Programa!Auxlog6) Then Arr_Programa(i).Auxlog6 = rs_Programa!Auxlog6
            If Not EsNulo(rs_Programa!Auxlog7) Then Arr_Programa(i).Auxlog7 = rs_Programa!Auxlog7
            If Not EsNulo(rs_Programa!Auxlog8) Then Arr_Programa(i).Auxlog8 = rs_Programa!Auxlog8
            If Not EsNulo(rs_Programa!Auxlog9) Then Arr_Programa(i).Auxlog9 = rs_Programa!Auxlog9
            If Not EsNulo(rs_Programa!Auxlog10) Then Arr_Programa(i).Auxlog10 = rs_Programa!Auxlog10
            If Not EsNulo(rs_Programa!Auxlog11) Then Arr_Programa(i).Auxlog11 = rs_Programa!Auxlog11
            If Not EsNulo(rs_Programa!Auxlog12) Then Arr_Programa(i).Auxlog12 = rs_Programa!Auxlog12
            If Not EsNulo(rs_Programa!Auxchar) Then Arr_Programa(i).Auxchar = rs_Programa!Auxchar
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
Dim i As Integer
Dim rs_Con_For_Tpa As New ADODB.Recordset

    StrSql = "SELECT * FROM con_for_tpa "
    StrSql = StrSql & " Order BY concnro,fornro,tpanro,nivel,selecc"
    OpenRecordset StrSql, rs_Con_For_Tpa
            
    Max_Con_For_Tpa = rs_Con_For_Tpa.RecordCount
    ReDim Preserve Arr_con_for_tpa(Max_Con_For_Tpa) As TCon_for_tpa
            
    i = 1
    Do While Not rs_Con_For_Tpa.EOF
            Arr_con_for_tpa(i).concnro = rs_Con_For_Tpa!concnro
            Arr_con_for_tpa(i).fornro = rs_Con_For_Tpa!fornro
            Arr_con_for_tpa(i).tpanro = rs_Con_For_Tpa!tpanro
            Arr_con_for_tpa(i).Nivel = rs_Con_For_Tpa!Nivel
            Arr_con_for_tpa(i).depurable = rs_Con_For_Tpa!depurable
            Arr_con_for_tpa(i).cftauto = rs_Con_For_Tpa!cftauto
            If Not EsNulo(rs_Con_For_Tpa!Selecc) Then Arr_con_for_tpa(i).Selecc = Trim(rs_Con_For_Tpa!Selecc)
            If Not EsNulo(rs_Con_For_Tpa!Prognro) Then Arr_con_for_tpa(i).Prognro = rs_Con_For_Tpa!Prognro
        
        i = i + 1
        rs_Con_For_Tpa.MoveNext
    Loop
    
    If rs_Con_For_Tpa.State = adStateOpen Then rs_Con_For_Tpa.Close
    Set rs_Con_For_Tpa = Nothing
End Sub

Public Function Indice_Arr_con_for_tpa(ByVal concepto As Long, ByVal Formula As Long, ByVal Parametro As Long, ByVal Nivel As Long, ByVal Selecc As String) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el indice en el arreglo de Arr_con_for_tpa.
' Autor      : FGZ
' Fecha      : 19/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    i = 1
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Arr_con_for_tpa(i).concnro = concepto Then
            If Arr_con_for_tpa(i).fornro = Formula Then
                If Arr_con_for_tpa(i).tpanro = Parametro Then
                    If Arr_con_for_tpa(i).Nivel = Nivel Then
                        If Not EsNulo(Trim(Selecc)) Then
                            If Arr_con_for_tpa(i).Selecc = Trim(Selecc) Then
                                Encontro = True
                            Else
                                i = i + 1
                                If EsNulo(Arr_con_for_tpa(i).concnro) Then
                                    Termino = True
                                End If
                            End If
                        Else
                            Encontro = True
                        End If
                    Else
                        If Arr_con_for_tpa(i).Nivel > Nivel Then
                            Termino = True
                        Else
                            i = i + 1
                            If EsNulo(Arr_con_for_tpa(i).concnro) Then
                                Termino = True
                            End If
                        End If
                    End If
                Else
                    If Arr_con_for_tpa(i).tpanro > Parametro Then
                        Termino = True
                    Else
                        i = i + 1
                        If EsNulo(Arr_con_for_tpa(i).concnro) Then
                            Termino = True
                        End If
                    End If
                End If
            Else
                If Arr_con_for_tpa(i).fornro = Formula Then
                    Termino = True
                Else
                    i = i + 1
                    If EsNulo(Arr_con_for_tpa(i).concnro) Then
                        Termino = True
                    End If
                End If
            End If
        Else
            If Arr_con_for_tpa(i).concnro > concepto Then
                Termino = True
            Else
                i = i + 1
                If EsNulo(Arr_con_for_tpa(i).concnro) Then
                    Termino = True
                End If
            End If
        End If
        
    Loop
    
    If Encontro Then
        Indice_Arr_con_for_tpa = i
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
Dim i As Integer
Dim rs_Cge_Segun As New ADODB.Recordset

    StrSql = "SELECT * FROM Cge_Segun "
    StrSql = StrSql & " Order BY concnro,nivel"
    OpenRecordset StrSql, rs_Cge_Segun
            
    Max_Cge_Segun = rs_Cge_Segun.RecordCount
    ReDim Preserve Arr_Cge_Segun(Max_Cge_Segun) As TCge_Segun
            
    i = 1
    Do While Not rs_Cge_Segun.EOF
            Arr_Cge_Segun(i).concnro = rs_Cge_Segun!concnro
            Arr_Cge_Segun(i).Nivel = rs_Cge_Segun!Nivel
            Arr_Cge_Segun(i).Origen = rs_Cge_Segun!Origen
            Arr_Cge_Segun(i).Entidad = rs_Cge_Segun!Entidad
        
        i = i + 1
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
Dim i As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    i = 1
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Arr_Cge_Segun(i).concnro = Arr_conceptos(Concepto_Actual).concnro Then
            If Arr_Cge_Segun(i).Nivel = 0 Then
                If Arr_Cge_Segun(i).Origen = buliq_cabliq!Empleado Then
                    Encontro = True
                Else
                    i = i + 1
                    If i <= UBound(Arr_Cge_Segun) Then
                        If EsNulo(Arr_Cge_Segun(i).concnro) Then
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
            If Arr_Cge_Segun(i).concnro > Arr_conceptos(Concepto_Actual).concnro Then
                Termino = True
            Else
                i = i + 1
                If i <= UBound(Arr_Cge_Segun) Then
                    If EsNulo(Arr_Cge_Segun(i).concnro) Then
                        Termino = True
                    End If
                Else
                    Termino = True
                End If
            End If
        End If
        
    Loop
    
    If Encontro Then
        Indice_Arr_Cge_Segun = i
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
Dim i As Integer
Dim rs_Cft_Segun As New ADODB.Recordset

    StrSql = "SELECT * FROM Cft_Segun "
    StrSql = StrSql & " Order BY concnro,tpanro,fornro,nivel,origen"
    OpenRecordset StrSql, rs_Cft_Segun
            
    Max_Cft_Segun = rs_Cft_Segun.RecordCount
    ReDim Preserve Arr_Cft_Segun(Max_Cft_Segun) As TCft_Segun
            
    i = 1
    Do While Not rs_Cft_Segun.EOF
            Arr_Cft_Segun(i).concnro = rs_Cft_Segun!concnro
            Arr_Cft_Segun(i).fornro = rs_Cft_Segun!fornro
            Arr_Cft_Segun(i).tpanro = rs_Cft_Segun!tpanro
            If Not EsNulo(rs_Cft_Segun!Nivel) Then Arr_Cft_Segun(i).Nivel = rs_Cft_Segun!Nivel
            If Not EsNulo(rs_Cft_Segun!Origen) Then Arr_Cft_Segun(i).Origen = rs_Cft_Segun!Origen
            If Not EsNulo(rs_Cft_Segun!Entidad) Then Arr_Cft_Segun(i).Entidad = rs_Cft_Segun!Entidad
            If Not EsNulo(rs_Cft_Segun!Selecc) Then Arr_Cft_Segun(i).Selecc = rs_Cft_Segun!Selecc
        
        i = i + 1
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
Dim i As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    i = Indice_Actual_Cft_Segun + 1
    'Controlo que no se salga de rango
    If i > Max_Cft_Segun Then
        Termino = True
    End If
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Arr_Cft_Segun(i).concnro = Arr_conceptos(Concepto_Actual).concnro Then
            If Arr_Cft_Segun(i).tpanro = Parametro Then
                If Arr_Cft_Segun(i).fornro = Formula Then
                    If Arr_Cft_Segun(i).Nivel = 0 Then
                        If Arr_Cft_Segun(i).Origen = buliq_cabliq!Empleado Then
                            Encontro = True
                        Else
                            i = i + 1
                            If i > Max_Cft_Segun Then
                                Termino = True
                            Else
                                If EsNulo(Arr_Cft_Segun(i).concnro) Then
                                    Termino = True
                                End If
                            End If
                        End If
                    Else
                        Encontro = True
                    End If
                Else
                    i = i + 1
                    If i > Max_Cft_Segun Then
                        Termino = True
                    Else
                        If EsNulo(Arr_Cft_Segun(i).concnro) Then
                            Termino = True
                        End If
                    End If
                End If
            Else
                i = i + 1
                If i > Max_Cft_Segun Then
                    Termino = True
                Else
                    If EsNulo(Arr_Cft_Segun(i).concnro) Then
                        Termino = True
                    End If
                End If
            End If
        Else
            If Arr_Cft_Segun(i).concnro > Arr_conceptos(Concepto_Actual).concnro Then
                Termino = True
            Else
                i = i + 1
                If i > Max_Cft_Segun Then
                    Termino = True
                Else
                    If EsNulo(Arr_Cft_Segun(i).concnro) Then
                        Termino = True
                    End If
                End If
            End If
        End If
        
    Loop
    
    If Encontro Then
        Buscar_Sig_Indice_Arr_Cft_Segun = i
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
Dim i As Integer
Dim rs_For_Tpa As New ADODB.Recordset

    StrSql = "SELECT * FROM for_tpa "
    StrSql = StrSql & " ORDER BY fornro,ftorden"
    OpenRecordset StrSql, rs_For_Tpa
            
    Max_For_Tpa = rs_For_Tpa.RecordCount
    ReDim Preserve Arr_For_Tpa(Max_For_Tpa) As TFor_tpa
    
    i = 1
    Do While Not rs_For_Tpa.EOF
            Arr_For_Tpa(i).fornro = rs_For_Tpa!fornro
            Arr_For_Tpa(i).tpanro = rs_For_Tpa!tpanro
            Arr_For_Tpa(i).ftentrada = rs_For_Tpa!ftentrada
            Arr_For_Tpa(i).ftimprime = rs_For_Tpa!ftimprime
            Arr_For_Tpa(i).ftobligatorio = rs_For_Tpa!ftobligatorio
            If Not EsNulo(rs_For_Tpa!ftorden) Then Arr_For_Tpa(i).ftorden = rs_For_Tpa!ftorden
            If Not EsNulo(rs_For_Tpa!ftinicial) Then Arr_For_Tpa(i).ftinicial = rs_For_Tpa!ftinicial
        
        i = i + 1
        rs_For_Tpa.MoveNext
    Loop
    
    If rs_For_Tpa.State = adStateOpen Then rs_For_Tpa.Close
    Set rs_For_Tpa = Nothing
End Sub


Public Function BuscarSiguiente_For_Tpa(ByVal concepto As Long) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el siguiente indice al actual en el arreglo de Arr_For_Tpa.
' Autor      : FGZ
' Fecha      : 19/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    i = Indice_Actual_For_Tpa + 1
    'Controlo que no se salga de rango
    Encontro = False
    Termino = False
    If i > Max_For_Tpa Then
        Termino = True
    End If
    Do While Not Encontro And Not Termino
        If Arr_For_Tpa(i).fornro = Arr_conceptos(Concepto_Actual).fornro Then
            Encontro = True
        Else
            If Arr_For_Tpa(i).fornro > Arr_conceptos(Concepto_Actual).fornro Then
                Termino = True
            Else
                i = i + 1
                If i > Max_For_Tpa Then
                    Termino = True
                Else
                    If EsNulo(Arr_For_Tpa(i).fornro) Then
                        Termino = True
                    End If
                End If
            End If
        End If
    Loop
    
    If Encontro Then
        BuscarSiguiente_For_Tpa = i
    Else
        BuscarSiguiente_For_Tpa = 0
    End If
End Function

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
Dim i As Integer
Dim rs_acumulador As New ADODB.Recordset

    StrSql = "SELECT * FROM acumulador "
    StrSql = StrSql & " ORDER BY acunro "
    OpenRecordset StrSql, rs_acumulador
    
    If Not rs_acumulador.EOF Then
        rs_acumulador.MoveLast
        Max_Acumuladores = rs_acumulador!acunro + 1
        
        ReDim Preserve Arr_Acumulador(Max_Acumuladores) As TAcumulador
        rs_acumulador.MoveFirst
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 2) & "No se encontraron Acumuladores"
        End If
    End If
    
    Do While Not rs_acumulador.EOF
            i = rs_acumulador!acunro
            Arr_Acumulador(i).acunro = rs_acumulador!acunro
        
            If Not EsNulo(rs_acumulador!acudesabr) Then
                Arr_Acumulador(i).acudesabr = rs_acumulador!acudesabr
            End If
            
            If Not EsNulo(rs_acumulador!acusist) Then
                Arr_Acumulador(i).acusist = rs_acumulador!acusist
            End If
            
            If Not EsNulo(rs_acumulador!acudesext) Then
                Arr_Acumulador(i).acudesext = rs_acumulador!acudesext
            End If
            
            If Not EsNulo(rs_acumulador!acumes) Then
                Arr_Acumulador(i).acumes = rs_acumulador!acumes
            End If
            
            If Not EsNulo(rs_acumulador!acutopea) Then
                Arr_Acumulador(i).acutopea = rs_acumulador!acutopea
            End If
            
            If Not EsNulo(rs_acumulador!acudesborde) Then
                Arr_Acumulador(i).acudesborde = rs_acumulador!acudesborde
            End If
                
            If Not EsNulo(rs_acumulador!acurecalculo) Then
                Arr_Acumulador(i).acurecalculo = rs_acumulador!acurecalculo
            End If
                
            If Not EsNulo(rs_acumulador!acuimponible) Then
                Arr_Acumulador(i).acuimponible = rs_acumulador!acuimponible
            End If
                
            If Not EsNulo(rs_acumulador!acuimpcont) Then
                Arr_Acumulador(i).acuimpcont = rs_acumulador!acuimpcont
            End If
                
            If Not EsNulo(rs_acumulador!acusel1) Then
                Arr_Acumulador(i).acusel1 = rs_acumulador!acusel1
            End If
                
            If Not EsNulo(rs_acumulador!acusel2) Then
                Arr_Acumulador(i).acusel2 = rs_acumulador!acusel2
            End If
                
            If Not EsNulo(rs_acumulador!acusel3) Then
                Arr_Acumulador(i).acusel3 = rs_acumulador!acusel3
            End If
                
            If Not EsNulo(rs_acumulador!acuppag) Then
                Arr_Acumulador(i).acuppag = rs_acumulador!acuppag
            End If
                
            If Not EsNulo(rs_acumulador!acudepu) Then
                Arr_Acumulador(i).acudepu = rs_acumulador!acudepu
            End If
                
            If Not EsNulo(rs_acumulador!acuhist) Then
                Arr_Acumulador(i).acuhist = rs_acumulador!acuhist
            End If
                
            If Not EsNulo(rs_acumulador!acumanual) Then
                Arr_Acumulador(i).acumanual = rs_acumulador!acumanual
            End If
                
            If Not EsNulo(rs_acumulador!acuimpri) Then
                Arr_Acumulador(i).acuimpri = rs_acumulador!acuimpri
            End If
                
            If Not EsNulo(rs_acumulador!tacunro) Then
                Arr_Acumulador(i).tacunro = rs_acumulador!tacunro
            End If
            
            If Not EsNulo(rs_acumulador!Empnro) Then
                Arr_Acumulador(i).Empnro = rs_acumulador!Empnro
            End If
                
            If Not EsNulo(rs_acumulador!acuorden) Then
                Arr_Acumulador(i).acuorden = rs_acumulador!acuorden
            End If
            
            If Not EsNulo(rs_acumulador!acuretro) Then
                Arr_Acumulador(i).acuretro = rs_acumulador!acuretro
            End If
                
            If Not EsNulo(rs_acumulador!acunoneg) Then
                Arr_Acumulador(i).acunoneg = rs_acumulador!acunoneg
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
Dim i As Integer
Dim rs_acumulador As New ADODB.Recordset

    StrSql = "SELECT * FROM acumulador "
    StrSql = StrSql & " ORDER BY acunro "
    OpenRecordset StrSql, rs_acumulador
    
    If Not rs_acumulador.EOF Then
        rs_acumulador.MoveLast
        Max_Acumuladores = rs_acumulador!acunro + 1
        
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
                Flog.writeline Espacios(Tabulador * 2) & "Acumulador " & rs_acumulador!acunro
            End If
    
            i = rs_acumulador!acunro
            Arr_Acumulador(i).acunro = rs_acumulador!acunro
        
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acudesabr"
            End If
            If Not EsNulo(rs_acumulador!acudesabr) Then
                Arr_Acumulador(i).acudesabr = rs_acumulador!acudesabr
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acusist"
            End If
            If Not EsNulo(rs_acumulador!acusist) Then
                Arr_Acumulador(i).acusist = rs_acumulador!acusist
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acudesext"
            End If
            If Not EsNulo(rs_acumulador!acudesext) Then
                Arr_Acumulador(i).acudesext = rs_acumulador!acudesext
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acumes"
            End If
            If Not EsNulo(rs_acumulador!acumes) Then
                Arr_Acumulador(i).acumes = rs_acumulador!acumes
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acutopea"
            End If
            If Not EsNulo(rs_acumulador!acutopea) Then
                Arr_Acumulador(i).acutopea = rs_acumulador!acutopea
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acudesborde"
            End If
            If Not EsNulo(rs_acumulador!acudesborde) Then
                Arr_Acumulador(i).acudesborde = rs_acumulador!acudesborde
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acurecalculo"
            End If
            If Not EsNulo(rs_acumulador!acurecalculo) Then
                Arr_Acumulador(i).acurecalculo = rs_acumulador!acurecalculo
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuimponible"
            End If
            If Not EsNulo(rs_acumulador!acuimponible) Then
                Arr_Acumulador(i).acuimponible = rs_acumulador!acuimponible
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuimpcont"
            End If
            If Not EsNulo(rs_acumulador!acuimpcont) Then
                Arr_Acumulador(i).acuimpcont = rs_acumulador!acuimpcont
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acusel1"
            End If
            If Not EsNulo(rs_acumulador!acusel1) Then
                Arr_Acumulador(i).acusel1 = rs_acumulador!acusel1
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acusel2"
            End If
            If Not EsNulo(rs_acumulador!acusel2) Then
                Arr_Acumulador(i).acusel2 = rs_acumulador!acusel2
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acusel3"
            End If
            If Not EsNulo(rs_acumulador!acusel3) Then
                Arr_Acumulador(i).acusel3 = rs_acumulador!acusel3
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuppag"
            End If
            If Not EsNulo(rs_acumulador!acuppag) Then
                Arr_Acumulador(i).acuppag = rs_acumulador!acuppag
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acudepu"
            End If
            If Not EsNulo(rs_acumulador!acudepu) Then
                Arr_Acumulador(i).acudepu = rs_acumulador!acudepu
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuhist"
            End If
            If Not EsNulo(rs_acumulador!acuhist) Then
                Arr_Acumulador(i).acuhist = rs_acumulador!acuhist
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acumanual"
            End If
            If Not EsNulo(rs_acumulador!acumanual) Then
                Arr_Acumulador(i).acumanual = rs_acumulador!acumanual
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuimpri"
            End If
            If Not EsNulo(rs_acumulador!acuimpri) Then
                Arr_Acumulador(i).acuimpri = rs_acumulador!acuimpri
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "tacunro"
            End If
            If Not EsNulo(rs_acumulador!tacunro) Then
                Arr_Acumulador(i).tacunro = rs_acumulador!tacunro
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "Empnro"
            End If
            If Not EsNulo(rs_acumulador!Empnro) Then
                Arr_Acumulador(i).Empnro = rs_acumulador!Empnro
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuorden"
            End If
            If Not EsNulo(rs_acumulador!acuorden) Then
                Arr_Acumulador(i).acuorden = rs_acumulador!acuorden
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acuretro"
            End If
            If Not EsNulo(rs_acumulador!acuretro) Then
                Arr_Acumulador(i).acuretro = rs_acumulador!acuretro
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
            End If
                
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 3) & "acunoneg"
            End If
            If Not EsNulo(rs_acumulador!acunoneg) Then
                Arr_Acumulador(i).acunoneg = rs_acumulador!acunoneg
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & " NULL"
                End If
                Arr_Acumulador(i).acunoneg = 0
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
Dim i As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    i = Acumulador_Actual + 1
    'Controlo que no se salga de rango
    If i > Max_Acumuladores Then
        Termino = True
    End If
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Not EsNulo(Arr_Acumulador(i).acunro) Then
            Encontro = True
        Else
            i = i + 1
            If i > Max_Acumuladores Then
                Termino = True
            End If
        End If
    Loop
    
    If Encontro Then
        Siguiente_Acumulador = i
    Else
        Siguiente_Acumulador = 0
    End If
End Function


Public Sub Cargar_Con_Acum(ByVal Nrotipo As Long, ByVal fecha_inicio As Date, ByVal fecha_fin As Date)
' ---------------------------------------------------------------------------------------------
' Descripcion: Carga todos los acumuladores a los cuales suma cada conceptos a liquidar.
' Autor      : FGZ
' Fecha      : 23/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim Ultimo_Concepto As Long
Dim rs_Con_Acum As New ADODB.Recordset


    StrSql = "SELECT * FROM con_acum "
    StrSql = StrSql & " INNER JOIN concepto ON con_acum.concnro = concepto.concnro "
    StrSql = StrSql & " INNER JOIN con_tp ON con_tp.concnro = concepto.concnro "
    StrSql = StrSql & " WHERE con_tp.tprocnro = " & Nrotipo
    StrSql = StrSql & " AND (concepto.concvalid = 0 or ( concdesde <= " & ConvFecha(fecha_inicio)
    StrSql = StrSql & " AND conchasta >= " & ConvFecha(fecha_fin) & "))"
    StrSql = StrSql & " ORDER BY concepto.tconnro, concepto.concorden, acunro "
    OpenRecordset StrSql, rs_Con_Acum

    ReDim Preserve Arr_Con_Acum(Max_Conceptos, Max_Acumuladores) As Integer

    i = 0
    Ultimo_Concepto = 0
    Do While Not rs_Con_Acum.EOF
        If Ultimo_Concepto <> rs_Con_Acum!concnro Then
            i = Buscar_Indice_Concepto(rs_Con_Acum!concnro, i)
        End If
        Arr_Con_Acum(i, rs_Con_Acum!acunro) = -1
        
        Ultimo_Concepto = rs_Con_Acum!concnro
        
        rs_Con_Acum.MoveNext
    Loop

    If rs_Con_Acum.State = adStateOpen Then rs_Con_Acum.Close
    Set rs_Con_Acum = Nothing
End Sub


Public Function Buscar_Indice_Concepto(ByVal concepto As Long, ByVal Posicion As Integer) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca el indice del concepto dentro del arreglo de conceptos Arr_concepto.
' Autor      : FGZ
' Fecha      : 23/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer
Dim Encontro As Boolean
Dim Termino As Boolean
    
    i = Posicion
    'Controlo que no se salga de rango
    If i > Max_Conceptos Then
        Termino = True
    End If
    Encontro = False
    Termino = False
    Do While Not Encontro And Not Termino
        If Arr_conceptos(i).concnro = concepto Then
            Encontro = True
        Else
            i = i + 1
            If i > Max_Conceptos Then
                Termino = True
            End If
        End If
    Loop
    
    If Encontro Then
        Buscar_Indice_Concepto = i
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
    End If
    Encontro = False
    Termino = False
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
