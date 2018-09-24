Attribute VB_Name = "mdlVac_Paraguay"

Sub actualizarDiasVacPY(ByVal auxNroVac As Long, ByVal DiasCorraGen As Double, ByVal NroTPV As Long, ByVal FechaHasta As Date)
 Dim rs As New ADODB.Recordset

    StrSql = "SELECT * FROM tipovacac WHERE tpvnrocol = " & NroTPV
    OpenRecordset StrSql, rs
        
    If Not rs.EOF Then
        NroTPV = rs!tipvacnro
    Else
        'Verifica si tiene el tipo de días de vacaciones configurado Pol(1501)
        'sino pone el Primero de la tabla por Default
        If (st_TipoDia1 > 0) Then
            NroTPV = st_TipoDia1
        Else
            NroTPV = 1 ' por default
        End If
    End If

 
    StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
    OpenRecordset StrSql, rs
        
    If Not rs.EOF Then
        If Reproceso Then
            If Not IsNull(NroTPV) And NroTPV > 0 Then
                StrSql = "UPDATE vacdiascor SET vdiascormanual = 0, vdiascorcant = " & DiasCorraGen & ", tipvacnro = " & NroTPV
                    
                ' Se agrego para CR el campo de la ultima generacion de Dias Correspondiente
                If Not IsNull(FechaHasta) Then
                    StrSql = StrSql & ",vdiasfechasta = " & ConvFecha(FechaHasta)
                End If
                StrSql = StrSql & " WHERE vacnro = " & auxNroVac & " AND Ternro = " & Ternro & " AND (venc = 0 OR venc IS NULL)"
                    
            Else
               Flog.writeline "Error al actualizar los dias correspondientes. Tipo de vacación incorrecto: " & NroTPV
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            If Not IsNull(NroTPV) And Not NroTPV > 0 Then
                StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES (" & _
                         auxNroVac & "," & DiasCorraGen & ",0," & Ternro & "," & NroTPV & "," & ConvFecha(FechaHasta) & ")"
            Else
                Flog.writeline "Error al insertar los dias correspondientes. Tipo de vacación incorrecto: " & NroTPV
            End If
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    
    Else
        If Not IsNull(NroTPV) And Not NroTPV > 0 Then
             StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES (" & _
                      auxNroVac & "," & DiasCorraGen & ",0," & Ternro & "," & NroTPV & "," & ConvFecha(FechaHasta) & ")"
         Else
            StrSql = "INSERT INTO vacdiascor(vacnro,vdiascorcant,vdiascormanual,ternro,tipvacnro,vdiasfechasta) VALUES (" & _
                      auxNroVac & "," & DiasCorraGen & ",0," & Ternro & "," & NroTPV & "," & ConvFecha(FechaHasta) & ")"
         End If
         objConn.Execute StrSql, , adExecuteNoRecords
    End If
End Sub
Public Sub bus_DiasVac_PY(ByVal Ternro As Long, ByRef NroVac As Long, ByVal fechaAlta As Date, ByRef FechaHasta As Date, ByRef cantdias As Double, ByRef Columna As Integer, ByRef Mensaje As String, ByRef Genera As Boolean, ByVal Anio As Long, ByVal bprcfechasta As Date, ByVal Reproceso As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula los días de vacaciones - PARAGUAY
' Autor      : Gonzalez Nicolás
' Fecha      : 17/09/2012
'              28/05/2013 - Gonzalez Nicolás -  Permite generar dias correspondientes anteriores.
' ---------------------------------------------------------------------------------------------
Dim totalDias As Long
Dim ultFechaProcesada As Date
Dim fechaCorte As Date

Dim rsDiasCorresp As New ADODB.Recordset
Dim rsRegHorario As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim rs_valgrilla As New ADODB.Recordset
Dim rs As New ADODB.Recordset



Dim DiasHorario As Double
Dim cantDiasProp As Double
Dim arrIntervalos()
Dim i As Integer
Dim actualizarDatos As Boolean
Dim IntervalosIncorrectos As Boolean
Dim cantDia As Integer
Dim cantMes As Integer
Dim cantAnio As Integer


Dim ExcluyeFeriados As Boolean
Dim ExcluyeFeriadosCorr  As Boolean




Dim antdia As Long
Dim antmes As Long
Dim antanio As Long
Dim q As Integer
Dim parametros(5) As Integer
Dim Encontro As Boolean



    'Activo el manejador de errores
    On Error GoTo CE
    
    'fechaBaja = Empty
    Genera = False
    'tieneFaseBaja = False
    
    'Obtiene la fecha de la ultima fecha de cálculo de dias corresp y la fase del empleado
    StrSql = "SELECT vdiascorcant,vdiascorcant,vdiasfechasta,bajfec,venc,estado FROM vacdiascor " & _
            "LEFT JOIN fases ON vacdiascor.ternro = fases.empleado" & _
            "  WHERE ternro= " & Ternro & " AND venc= 0 AND estado=-1 ORDER BY vdiasfechasta DESC"
    OpenRecordset StrSql, rsDiasCorresp
    
    If Not rsDiasCorresp.EOF Then
        ultFechaProcesada = rsDiasCorresp!vdiasfechasta
        cantdias = rsDiasCorresp!vdiascorcant
        Flog.writeline "Ultima fecha de calculo de vacaciones " & rsDiasCorresp!vdiasfechasta
    Else
        StrSql = "SELECT * FROM vacdiascor WHERE ternro=" & Ternro & " ORDER BY vdiasfechasta DESC"
        OpenRecordset StrSql, rsDiasCorresp
        
        If Not rsDiasCorresp.EOF Then
            ultFechaProcesada = rsDiasCorresp!vdiasfechasta
            Flog.writeline "Ultima fecha de calculo de vacaciones " & rsDiasCorresp!vdiasfechasta
        Else
            ultFechaProcesada = fechaAlta
            Flog.writeline "No se encontro días correspondientes para el empleado y se toma la fecha de alta." & fechaAlta
        End If
    End If
    
   
'    If (CDate(ultFechaProcesada) >= CDate(FechaHasta)) And Reproceso = False Then
'        Flog.writeline "La fecha ingresada ya fue procesada. Ultimo fecha de procesamiento (" & ultFechaProcesada & ")"
'        Exit Sub
'    Else
       
    'Guardo la fecha de procesamiento
    ultFechaProcesada = bprcfechasta
       
    Genera = False
    Encontro = False
    
    Call Politica(1502)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Sub
    End If
    

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Flog.writeline "La escala de Vacaciones no esta configurada o el nro de grilla no esta bien configurado bien en la Politica 1502. Grilla " & NroGrilla
        Exit Sub
    End If
    
    Call Politica(1505)
    If Not PoliticaOK Then
        Flog.writeline "Error cargando configuracion de la Politica 1505. Tipo Base antiguedad estandar."
        VersionBaseAntig = 0
    Else
        VersionBaseAntig = st_BaseAntiguedad
    End If
    
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop

    'setea la proporcion de dias
    Call Politica(1501)
        
    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Select Case VersionBaseAntig
            Case 0:
                Flog.writeline "Antiguedad estandar "
                Flog.writeline "Antiguedad En el ultimo año " ' Se computa al año actual
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            
            Case 1:
                Flog.writeline "Antiguedad Sin redondeo "
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                      Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 2:
                Flog.writeline "Antiguedad Uruguay " ' Se computa al año anterior
                'Call bus_Antiguedad_G("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio - 1), antdia, antmes, antanio, q)
            Case 3:
                 Flog.writeline "Antiguedad Standard " ' Se computa al año actual
                 Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 If (((antmes * 30) + antdia) >= st_Dias) Or antanio <> 0 Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            Case 4: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año siguiente"
                If Not (st_Dia = 0 Or st_Mes = 0) Then
                     Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                     If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                       Call bus_Antiguedad_G("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio + 1), antdia, antmes, antanio, q)
                    End If
                 End If
            Case 5: ' Anguedad a una fecha dada por dia y mes del año
                Flog.writeline "Antiguedad a una fecha dada año actual"
                Call bus_Antiguedad_R("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
                If (((antmes * 30) + antdia >= st_Dias) Or antanio <> 0) Then
                       antdia = 0
                       antmes = 0
                       antanio = 0
                 Call bus_Antiguedad("VACACIONES", CDate(st_Dia & "/" & st_Mes & "/" & Periodo_Anio), antdia, antmes, antanio, q)
                 End If
            
            
            Case 6:
                'Calcula Antiguedad a una fecha
                Flog.writeline "Antiguedad a una fecha - PARAGUAY "
                Call bus_Antiguedad("VACACIONES", CDate(FechaHasta), antdia, antmes, antanio, q)
                
            Case Else
                Flog.writeline "Antiguedad Mal configurada. Estandar "
                'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
                Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Periodo_Anio), antdia, antmes, antanio, q)
            End Select
            parametros(j) = (antanio * 12) + antmes
            'Call bus_Antiguedad("VACACIONES", CDate("31/12/" & Year(fecha_desde)), antdia, antmes, antanio, q)
            Flog.writeline "Años " & antanio
            Flog.writeline "Meses " & antmes
            Flog.writeline "Dias " & antdia
            
        Case Else:
            Select Case j
            Case 1:
                Call bus_Estructura(rs_cabgrilla!grparnro_1)
            Case 2:
                Call bus_Estructura(rs_cabgrilla!grparnro_2)
            Case 3:
                Call bus_Estructura(rs_cabgrilla!grparnro_3)
            Case 4:
                Call bus_Estructura(rs_cabgrilla!grparnro_4)
            Case 5:
                Call bus_Estructura(rs_cabgrilla!grparnro_5)
            End Select
            parametros(j) = valor
        End Select
    Next j
    
    
    cantdias = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcion, Encontro)
    Columna = TipoVacacionProporcion
    cantdiasCorr = buscarDiasVacEscala(ant, rs_cabgrilla!cgrdimension, parametros, TipoVacacionProporcionCorr)
    columna2 = TipoVacacionProporcionCorr

    '------------------------------
    'llamada politica 1513
    '------------------------------
    
    'EAM- Tiene en cuenta los dias trabajados en el ultimo año
    Call Politica(1513)
    
    
    If Dias_efect_trab_anio Then
        Flog.writeline "Tiene en cuenta el ultimo año. Politica 1513."
            
        'Obtiene la proporcion de dias_trabajados  -->  (dias trabajados / 7) * regimen horio
        dias_efect_trabajado = DiasHabilTrabajado(Ternro, CDate("01/01/" & Periodo_Anio), CDate("31/12/" & Periodo_Anio))
        regHorarioActual = BuscarRegHorarioActual(Ternro)
        Aux_Dias_trab = ((180 / 7) * regHorarioActual)
        Aux_Dias_trab = RedondearNumero(Int(Aux_Dias_trab), (Aux_Dias_trab - Int(Aux_Dias_trab)))
        
        If dias_efect_trabajado <= Aux_Dias_trab Then
            Encontro = True
            cantdias = CalcularProporcionDiasVac(dias_efect_trabajado)
            
            Flog.writeline "Empleado " & Ternro & " con dias trabajado menor a mitad de año: " & dias_efect_trabajado
            Flog.writeline "Días Correspondientes: " & cantdias
            Flog.writeline "Tipo de redondeo: " & st_redondeo
            Flog.writeline "Parte decimal de los días correspondientes: " & aux_redondeo
            Flog.writeline
        End If
        
    End If
            
            
                                
    If Not Encontro Then
                
        'EAM- Si la columna1 es = a vacion no tiene el tipo de vacacion sino ya se configuro por la politica 1501 (columna,columna2) 10/02/2011
        If (Columna = 0) Then
            'Busco si existe algun valor para la estructura y ...
            'si hay que carga la columna correspondiente
            StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
            StrSql = StrSql & " AND vgrvalor is not null"
            For j = 1 To rs_cabgrilla!cgrdimension
                If j <> ant Then
                    StrSql = StrSql & " AND vgrcoor_" & j & "= " & parametros(j)
                End If
            Next j
            OpenRecordset StrSql, rs_valgrilla
            If Not rs_valgrilla.EOF Then
                Columna = rs_valgrilla!vgrorden
            Else
                Columna = 1
            End If
        End If
        
        dias_trabajados = ((antanio * 365) + (antmes * 30) + antdia)
        Flog.writeline "Dias trabajados " & dias_trabajados
        
        Flog.writeline "ANT " & ant
        
        If parametros(ant) <= BaseAntiguedad Then
            
            habiles = cantDiasLaborable(TipoVacacionProporcion, ExcluyeFeriados)
            habilesCorr = cantDiasLaborable(TipoVacacionProporcionCorr, ExcluyeFeriadosCorr)
            
            If ExcluyeFeriados Then
                'deberia revisar dia por dia de los dias contemplados para la antiguedad revisando si son feriados y dia habil
            End If
           
            If FactorDivision = 0 Then
                FactorDivision = 1
            End If
           
            Flog.writeline "Empleado " & Ternro & " con menos de 12 meses de trabajo."
            Flog.writeline "Dias Proporcion " & DiasProporcion
            Flog.writeline "Factor de Division " & FactorDivision
            Flog.writeline "Tipo Base Antiguedad " & BaseAntiguedad
            Flog.writeline "Dias habiles " & habiles
            Flog.writeline "Dias habiles Corridos " & habilesCorr
                 
             '_______________________________________________
             'POR EL MOMENTO NO SE PERMITE CREAR PROPORCIONAL
             '-----------------------------------------------
             If dias_trabajados < 12 Or 1 = 1 Then
                cantdias = 0
             Else
                If DiasProporcion = 12 Then
                        cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
                    Else
                        cantdias = Fix(12 * (dias_trabajados / DiasProporcion) / FactorDivision)
                End If
                
                aux_redondeo = ((dias_trabajados / DiasProporcion) / 7 * habiles) - Fix((dias_trabajados / DiasProporcion) / 7 * habiles)
                cantdias = RedondearNumero(cantdias, aux_redondeo)
              
                
                '- Obtiene los dias corridos de vacaciones a partir de los dias correspondientes
                cantdiasCorr = (cantdias * habilesCorr) / habiles
                aux_redondeo = ((cantdias * habilesCorr) / habiles) - Fix(((cantdias * habilesCorr) / habiles))
                cantdiasCorr = RedondearNumero(cantdiasCorr, aux_redondeo)
                
            End If
            Flog.writeline "Días Correspondientes:" & cantdias
            Flog.writeline "Días Correspondientes Corridos:" & cantdiasCorr
            Flog.writeline "Tipo de redondeo:" & st_redondeo
            Flog.writeline "Parte decimal de los días correspondientes:" & aux_redondeo
            Flog.writeline
            
            '-------------- Vacaciones Acordadas ------------------------------
            PoliticaOK = False
            DiasAcordados = False
            Call Politica(1511)
            If PoliticaOK And DiasAcordados Then
                 StrSql = "SELECT tipvacnro, diasacord FROM vacdiasacord "
                 StrSql = StrSql & " WHERE ternro = " & Ternro
                 OpenRecordset StrSql, rs
                 If Not rs.EOF Then
                     If rs!diasacord > cantdias Then
                         Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
                         Flog.writeline "Se utilizará la cantidad de dias acordados"
                         cantdias = rs!diasacord
                     End If
                 End If
            End If
            '25/06/2009 ------------- Vacaciones Acordadas ------------------------------
            Flog.writeline
            
         
            PoliticaOK = False
            Call Politica(1508)
            If PoliticaOK Then
                Flog.writeline "Politica 1508 activa. Analizando Licencias por Maternidad (" & Tipo_Dia_Maternidad & ")."
                dias_maternidad = 0
                'StrSql = "SELECT * FROM emplic "
                StrSql = "SELECT SUM(elcantdias) total FROM emp_lic "
                StrSql = StrSql & " WHERE tdnro = " & Tipo_Dia_Maternidad
                StrSql = StrSql & " AND empleado = " & Ternro
                StrSql = StrSql & " AND elfechadesde >= " & ConvFecha("01/01/" & (Periodo_Anio - 1))
                StrSql = StrSql & " AND elfechahasta <= " & ConvFecha("31/12/" & (Periodo_Anio - 1))
                OpenRecordset StrSql, rs
                If Not rs.EOF And (Not IsNull(rs!total)) Then
                    dias_maternidad = rs!total
                    If dias_maternidad <> 0 Then
                        Flog.writeline "  Dias por maternidad: " & dias_maternidad
                        Flog.writeline "  Dias = " & cantdias & " - (" & dias_maternidad & " x " & Factor & ")"
                        cantdias = cantdias - CInt(dias_maternidad * Factor)
                    End If
                Else
                    Flog.writeline "  No se encontraron dias por maternidad."
                End If
                rs.Close
            End If
        Else
            Flog.writeline "No se encontro la escala para el convenio"
            Genera = False
        End If
    Else
        'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
        PoliticaOK = False
        DiasAcordados = False
        Call Politica(1511)
        If PoliticaOK And DiasAcordados Then
             StrSql = "SELECT tipvacnro, diasacord FROM vacdiasacord "
             StrSql = StrSql & " WHERE ternro = " & Ternro
             OpenRecordset StrSql, rs
             If Not rs.EOF Then
                 If rs!diasacord > cantdias Then
                     Flog.writeline "La cantidad de dias correspondientes es menor a la cantidad de dias acordados. " & rs!diasacord
                     Flog.writeline "Se utilizará la cantidad de dias acordados"
                     cantdias = rs!diasacord
                 End If
             End If
        End If
        'FGZ - 25/06/2009 ------------- Vacaciones Acordadas ------------------------------
        Flog.writeline
    End If
   
    Genera = True
    If Genera = True Then
        Call actualizarDiasVacPY(NroVac, cantdias, TipoVacacionProporcion, ultFechaProcesada)
    End If
    
        
    'End If

    GoTo finalizado
CE:
    Flog.writeline " ------------------------------------------------------------"
    Flog.writeline "Error procesando Empleado:" & Ternro & " Año " & Anio
    Flog.writeline Err.Description
    Flog.writeline "SQL: " & StrSql
    Flog.writeline " ------------------------------------------------------------"

sinDatos:
    Exit Sub
finalizado:
    Set rsDiasCorresp = Nothing
    Set rsRegHorario = Nothing
End Sub
Public Function Valida_Periodo_STD(ByVal NroVac As Long)
    ' ---------------------------------------------------------------------------------------------
' Descripcion: Valida los períodos para un empleado. Devuelve vacnro
' Autor      : Gonzalez Nicolás
' Fecha      : 07/11/2013
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
                Flog.writeline EscribeLogMI("Buscando datos del periodo")
                StrSql = " SELECT * FROM vacacion WHERE vacacion.vacnro = " & NroVac
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    'Alcances (1 Global, 2 Por Estrucutra y 3 Individual)
                    vac_alcannivel = objRs!Alcannivel ' Luego se valida de que forma se insertan los registros. en vacacion/vac_alcan
                    
                    fecha_desde = objRs!vacfecdesde
                    fecha_hasta = objRs!vacfechasta
                    Periodo_Anio = objRs!vacanio
                    
                    AnioaProc = Periodo_Anio '14/05/2012
                    auxNroVac = NroVac       '14/05/2012
                Else
                    Flog.writeline EscribeLogMI("No se encontró el periodo de vacaciones para el ternro") & ": " & Ternro
                    Exit Function
                End If
                
                'Devuelve 0 en caso que no cumpla con los alcances definidos por politica y por vacación.
                Valida_Periodo_STD = PeriodoCorrespondienteAlcance(Ternro, AnioaProc, vac_alcannivel)

                
End Function
Sub generarPeriodoVacacionAlance(Ternro As Long, Anio As Integer, Optional modeloPais As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Se encarga de generar el periodo de vacacion. (Para los que tienen 1 periodo por empleado)
' Autor      : Gonzalez Nicolás
' Fecha      : 07/11/2013
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Dim rs_vacacion As New ADODB.Recordset
    Dim fechaAlta As Date
    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim vacdesc As String
    Dim vacestado As Integer
    Dim Alcannivel As Integer
  
    Alcannivel = 1 'Periodos individuales

    fechaAlta = FechaAltaEmpleado(Ternro)
    
    If Anio < Year(fechaAlta) Then
        Flog.writeline "Error al querer generar un periodo anterior a la fecha de alta del empleado."
        Exit Sub
    End If
    fechaInicio = formatFecha(CStr(Day(fechaAlta)), CStr(Month(fechaAlta)), CStr(Anio))
    fechaFin = DateAdd("d", -1, DateAdd("yyyy", 1, fechaInicio))
    vacdesc = CStr(Anio) & " - " & CStr(Anio + 1)
    
    vacestado = -1 ' NG
    
    If modeloPais = 3 Then
        If Anio < Year(Date) - 4 Then
            vacestado = 0
        Else
            vacestado = -1
        End If
    End If
    
    
    'Si Nrovac es 0, creo nuevo período en la tabla vacacion
    If NroVac = 0 Then
        'Creo el período en la tabla Vacacion
        StrSql = " INSERT INTO vacacion (vacdesc, vacfecdesde, vacfechasta, vacanio, empnro, vacestado,alcannivel) "
        StrSql = StrSql & " VALUES ('" & vacdesc & "'," & ConvFecha(fechaInicio) & "," & ConvFecha(fechaFin) & "," & Anio & ",0," & vacestado & "," & Alcannivel & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        'Recuero el vacnro insertado
        NroVac = getLastIdentity(objConn, vacacion)
    End If
    
    Flog.writeline "Busco si tiene período ya generado en vac_alcan"
    StrSql = "SELECT * FROM vac_alcan WHERE vacnro=" & NroVac
    StrSql = StrSql & " AND origen=" & Ternro
    StrSql = StrSql & " AND alcannivel=" & Alcannivel
    OpenRecordset StrSql, rs_vacacion
    If Not rs_vacacion.EOF Then
        'Guardo los valores del periodo existente.
        fecha_desde = rs_vacacion!vacfecdesde
        fecha_hasta = rs_vacacion!vacfechasta
        
    Else
        'Creo el período individual en vac_alcan
        StrSql = " INSERT INTO vac_alcan (vacnro,vacfecdesde,vacfechasta,alcannivel,origen, vacestado) "
        StrSql = StrSql & " VALUES (" & NroVac & "," & ConvFecha(fechaInicio) & "," & ConvFecha(fechaFin) & "," & Alcannivel & "," & Ternro & "," & vacestado & ")"
        objConn.Execute StrSql, , adExecuteNoRecords
        
        'Guardo los valores del periodo insertado.
        fecha_desde = fechaInicio
        fecha_hasta = fechaFin
    
    End If
    Periodo_Anio = Anio
End Sub
