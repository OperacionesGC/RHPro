Attribute VB_Name = "MdlTiposDeDiasRHPro"
Option Explicit

Public Sub Licencias(Nro_Analisis As Long, Filtrar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Analisis para tipo de factor 6 (Licencias de Rhpro)
' Autor      : FGZ
' Fecha      : 15/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

'Variables locales
Dim Desde As Date
Dim Hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim tercero As Long

Dim cantdiasper As Integer
Dim cantdiasran As Integer
Dim porcentaje As Single
Dim monto_saldo As Single
Dim monto_total As Single
Dim cant_saldo As Single
Dim cant_total As Single
Dim cubvalor1 As Single
Dim cubvalor2 As Single
Dim cubvalor3 As Single
Dim cubvalor4 As Single

'Variables para los first y last
Dim PrimerFactOri As Boolean
Dim TipoEstr As Long
Dim EstrAct As Long
Dim FactOri As Long
Dim Ok As Boolean
Dim cant_flt As Long
Dim CantRangos As Integer

' Fechas desde y hasta a analizar por el proceso entero
Dim Inicio_Per_Analizado As Date
Dim Fin_Per_Analizado As Date

' Fechas parciales que se estan analizando
Dim Dia_Inicio_Per_Analizado As Date
Dim Dia_Fin_Per_Analizado As Date

Dim rs As New ADODB.Recordset
Dim rsLicencias As New ADODB.Recordset
Dim rsAnrCab As New ADODB.Recordset
Dim rsFiltro As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset

Dim Dias As Long
Dim prod

'Código -------------------------------------------------------------------

'Abro la conexion para Nexus
'OpenConnection strConexionNexus, objConnNexus

On Error GoTo CE:

'Obtengo la cabecera
Call ObtenerCabecerayFiltro(Nro_Analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If

'Estoy reprocesando
Call PurgarCubo(Nro_Analisis, 6)


' Obtengo los rangos del analisis
StrSql = "SELECT * FROM anrrangofec" & _
         " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro & _
         " AND anrrangorepro = -1 "
OpenRecordset StrSql, rsRango
                    
Progreso = SumPorcTiempo
CantRango = 0
If Not rsRango.EOF Then
    CantRango = rsRango.RecordCount
Else
    CantRango = 1
End If

                    
Do While Not rsRango.EOF
    Fin_Per_Analizado = rsRango!anrrangfechasta
    Inicio_Per_Analizado = rsRango!anrrangfecdesde

    Dia_Inicio_Per_Analizado = Inicio_Per_Analizado
    Dia_Fin_Per_Analizado = Fin_Per_Analizado

    ' -----------------------------------------------------------------
    '/* Recorre para el analisis los l de tipos de horas configurados */
    StrSql = "SELECT * FROM anrcab_fact" & _
        " INNER JOIN anrfact_ori ON anrfact_ori.facnro = anrcab_fact.facnro" & _
        " AND anrfact_ori.tipfacnro = 6" & _
        " INNER JOIN anrfactor ON anrfactor.facnro = anrcab_fact.facnro" & _
        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
        " ORDER BY anrfact_ori.facnro"
    OpenRecordset StrSql, rsFactor
        
    If Not rsFactor.EOF Then
        'Para el simular el first_of
        PrimerFactOri = True
        'Para el simular el last_of en la tabla anrfact_ori
        FactOri = rsFactor!facnro
    End If
        
    CantFactor = 0
    If Not rsFactor.EOF Then
        CantFactor = rsFactor.RecordCount
    End If
        
    ' obtengo el conjunto de legajos a procesar
    Call ObtenerLegajos(4, Filtrar, rsAnrCab!anrcabnro, rsFiltro, rsAnrCab!anrcabfecdesde, rsAnrCab!anrcabfechasta)
    CantFiltro = 0
    If Not rsFiltro.EOF Then
        CantFiltro = rsFiltro.RecordCount
    Else
        CantFiltro = 1
    End If
        
    If (CantFactor <> 0 And CantRango <> 0 And CantFiltro <> 0) Then
        prod = CantFactor * CantFiltro
        prod = prod * CantRango
        IncPorc = PorcTiempo / prod
    Else
        IncPorc = PorcTiempo
    End If
        
    'Recorro las Licencias que entran el el rango actual analizado
    Do While Not rsFactor.EOF
        Flog.writeline Espacios(Tabulador * 2) & "Factor " & rsFactor!facnro & " Origen " & rsFactor!faccodorig
        ' voy al primer legajo nuevamente
        rsFiltro.MoveFirst
        
        Do While Not rsFiltro.EOF
            If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
                GoTo siguientelegajo
            End If
        
            StrSql = "SELECT * FROM emp_lic " & _
                " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
                " WHERE emp_lic.empleado = " & rsFiltro!Ternro & _
                " AND tipdia.tdnro = " & rsFactor!faccodorig & _
                " AND emp_lic.elfechadesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
                " AND emp_lic.elfechahasta >= " & ConvFecha(Dia_Inicio_Per_Analizado)
            OpenRecordset StrSql, rsLicencias
            
            If Not rsLicencias.EOF Then
                'IncPorc = CantRangos * (((100 / CantFactor) * (100 / rsLicencias.RecordCount)) / 100) / 100
            Else
                Flog.writeline Espacios(Tabulador * 3) & "Licencia de tipo " & rsFactor!faccodorig & " sin liquidar para el legajo " & rsFiltro!empleg & " entre el " & Inicio_Per_Analizado & " y el " & Fin_Per_Analizado
            End If
        
            Do While Not rsLicencias.EOF
    
                            ' Obtengo las estructuras
                            Call ObtenerEstructuras(Filtrar, rsLicencias!empleado, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado, rsEstructura)
                            
                            If Not rsEstructura.EOF Then
                                TipoEstr = rsEstructura!tenro
                                EstrAct = rsEstructura!estrnro
                            End If
                        
                            Do While Not rsEstructura.EOF
        
                                If PrimerFactOri Then
                                    cantdiasper = DateDiff("d", Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado) + 1
                                    monto_total = 0
                                    cant_total = 0
                                    cant_saldo = 0
                                    PrimerFactOri = False
                                End If
                            
                                Dias = CantidadDeDias(Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado, rsLicencias!elfechadesde, rsLicencias!elfechahasta)
                                'monto_total = monto_total + rsLicencias!elcantdias
                                monto_total = monto_total + Dias
                                cant_total = cant_total
                                
                                '/* Calculo los dias de rango entre las fechas del rango y
                                ' el his_estruct para proporcionar*/
                                If rsFactor!facpropor = -1 Then
                                    If rsEstructura!htetdesde < Dia_Inicio_Per_Analizado Then
                                            If rsEstructura!htethasta < Dia_Fin_Per_Analizado And (Not IsNull(rsEstructura!htethasta)) Then
                                                cantdiasran = DateDiff("d", Dia_Inicio_Per_Analizado, rsEstructura!htethasta) + 1
                                            Else
                                                cantdiasran = DateDiff("d", Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado) + 1
                                            End If
                                    Else
                                        If (rsEstructura!htethasta < Dia_Fin_Per_Analizado) And (Not IsNull(rsEstructura!htethasta)) Then
                                            cantdiasran = DateDiff("d", rsEstructura!htetdesde, rsEstructura!htethasta) + 1
                                        Else
                                            cantdiasran = DateDiff("d", rsEstructura!htetdesde, Dia_Fin_Per_Analizado) + 1
                                        End If
                                    End If
                                    
                                    '/* Porcentaje segun la cant. de dias en la his_estrutura */
                                    porcentaje = cantdiasran * 100 / cantdiasper
                                    
                                    'If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    If Ultimo(rsLicencias) Or Last_OF_estrnro() Then
                                        ' se supone que no se proporciona
                                        cubvalor1 = monto_total
                                        cubvalor2 = cant_total
                                        
                                        StrSql = "SELECT * FROM anrcubo" & _
                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                            " AND facnro = " & rsFactor!facnro & _
                                            " AND tenro = " & rsEstructura!tenro & _
                                            " AND estrnro = " & rsEstructura!estrnro & _
                                            " AND ternro = " & rsAcumDiario!Ternro & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                        OpenRecordset StrSql, rs
            
                                        'Si el cubo no existe lo creo
                                        If rs.EOF Then
                                        '/* Creo el cubo */
                                            StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                ",anrcubvalor1,anrcubvalor2"
                                                '---------------------------
                                                ' FAF 14-02-2005
                                                If CInt(rsFactor!facpresup) = -1 Then
                                                    StrSql = StrSql & ",anrcubvalor3,anrcubvalor4"
                                                End If
                                                '---------------------------
                                            StrSql = StrSql & ") VALUES (" & _
                                                rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                rsEstructura!estrnro & "," & rsFactor!facnro & "," & _
                                                rsEstructura!tenro & "," & rsLicencias!empleado & ",1"
                                        End If
                                        
                                        If Not rs.EOF Then
                                            monto_saldo = (monto_total - cubvalor1 - rs!anrcubvalor1)
                                            cant_saldo = (cant_total - cubvalor2 - cant_saldo)
                                        End If
                                        'monto_saldo = (monto_total - cubvalor1 - monto_saldo)
                                        monto_total = 0
                                        cant_total = 0
                                        
                                        '* Para que no quede saldo cuando proporciona */
                                        If monto_saldo <= 1 And monto_saldo > 0 Then
                                            'cubvalor1 = cubvalor1 + monto_saldo
                                        End If
                                        
                                        If cant_saldo <= 1 And cant_saldo > 0 Then
                                            'cubvalor2 = cubvalor2 + cant_saldo
                                        End If
                                           
                                        '---------------------------
                                        ' FAF 14-02-2005
                                        ' Si presupuesta, realizo el calculo
                                        'If CInt(rsFactor!facpresup) = -1 Then
                                        '    cubvalor3 = CalcularPresupuestado(cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                        '    cubvalor4 = CalcularPresupuestado(cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                        'End If
                                        '---------------------------
                                        
                                        'Si existe el cubo entonces actualizo
                                        If Not rs.EOF Then
                                            StrSql = "UPDATE anrcubo SET" & _
                                                " anrcubvalor1 = " & Round(rs!anrcubvalor1 + cubvalor1, 2) & _
                                                " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2
                                                '---------------------------
                                                ' FAF 14-02-2005
                                                If CInt(rsFactor!facpresup) = -1 Then
                                                    cubvalor3 = CalcularPresupuestado(Round(rs!anrcubvalor1 + cubvalor1, 2), rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 + cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                                    StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                                End If
                                                '---------------------------
                                            StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & rsFactor!facnro & _
                                                " AND tenro = " & rsEstructura!tenro & _
                                                " AND estrnro = " & rsEstructura!estrnro & _
                                                " AND ternro = " & rsLicencias!empleado & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                        Else
                                            StrSql = StrSql & "," & cubvalor1 & "," & cubvalor2
                                                '---------------------------
                                                ' FAF 14-02-2005
                                                If CInt(rsFactor!facpresup) = -1 Then
                                                    cubvalor3 = CalcularPresupuestado(cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    cubvalor4 = CalcularPresupuestado(cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                    StrSql = StrSql & "," & cubvalor3 & "," & cubvalor4
                                                End If
                                                '---------------------------
                                            StrSql = StrSql & ")"
                                        End If
                                        
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                    End If
                                Else
                                    '/* Si no proporciona tomo al 100% y la ultima his_estruc del rango*/
                                    porcentaje = 100
                                    'If Last_OF(rsFactor!facnro) Or Last_OF(rsEstructura!estrnro) Then
                                    'If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    If Ultimo(rsLicencias) Or Last_OF_estrnro() Then
                                        If Not Last_OF_tenro() Then
                                            monto_total = 0
                                            cant_total = 0
                                        Else
                                        '/*Busco la ultima his_estr dentro del rango*/
                                            StrSql = "SELECT * FROM his_estructura " & _
                                                " WHERE his_estructura.ternro = " & rsLicencias!empleado & _
                                                " AND his_estructura.tenro = " & rsEstructura!tenro & _
                                                " AND his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
                                                " AND (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & _
                                                " OR his_estructura.htethasta IS NULL) "
                                            OpenRecordset StrSql, objRs
                                            objRs.MoveLast
                                            
                                            If Not objRs.EOF Then
                                            
                                                StrSql = "SELECT * FROM anrcubo" & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & rsFactor!facnro & _
                                                    " AND tenro = " & objRs!tenro & _
                                                    " AND estrnro = " & objRs!estrnro & _
                                                    " AND ternro = " & rsLicencias!empleado & _
                                                    " AND anrrangnro = " & rsRango!anrrangnro
                                                OpenRecordset StrSql, rs
                                                
                                                cubvalor1 = monto_total * porcentaje / 100
                                                cubvalor2 = cant_total * porcentaje / 100
                                                
                                                If rs.EOF Then
                                                    '/* Creo el cubo */
                                                    StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                        ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                        ",anrcubvalor1,anrcubvalor2"
                                                        '---------------------------
                                                        ' FAF 14-02-2005
                                                        If CInt(rsFactor!facpresup) = -1 Then
                                                            StrSql = StrSql & ",anrcubvalor3,anrcubvalor4"
                                                        End If
                                                        '---------------------------
                                                    StrSql = StrSql & ") VALUES (" & _
                                                        rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                        objRs!estrnro & "," & rsFactor!facnro & "," & _
                                                        objRs!tenro & "," & rsLicencias!empleado & ",1" & _
                                                        "," & cubvalor1 & "," & cubvalor2
                                                        '---------------------------
                                                        ' FAF 14-02-2005
                                                        If CInt(rsFactor!facpresup) = -1 Then
                                                            cubvalor3 = CalcularPresupuestado(cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                            cubvalor4 = CalcularPresupuestado(cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                            StrSql = StrSql & "," & cubvalor3 & "," & cubvalor4
                                                        End If
                                                        '---------------------------
                                                    StrSql = StrSql & ")"
                                                Else
                                                    StrSql = "UPDATE anrcubo SET" & _
                                                        " anrcubvalor1 = " & rs!anrcubvalor1 + cubvalor1 & _
                                                        " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2
                                                        '---------------------------
                                                        ' FAF 14-02-2005
                                                        If CInt(rsFactor!facpresup) = -1 Then
                                                            cubvalor3 = CalcularPresupuestado(rs!anrcubvalor1 + cubvalor1, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                            cubvalor4 = CalcularPresupuestado(rs!anrcubvalor2 + cubvalor2, rsFactor!facpresup, rsFactor!facopfijo, rsFactor!facopsuma, rsFactor!facpresupmonto)
                                                            StrSql = StrSql & " ,anrcubvalor3 = " & cubvalor3
                                                            StrSql = StrSql & " ,anrcubvalor4 = " & cubvalor4
                                                        End If
                                                        '---------------------------
                                                    StrSql = StrSql & " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                        " AND facnro = " & rsFactor!facnro & _
                                                        " AND tenro = " & objRs!tenro & _
                                                        " AND estrnro = " & objRs!estrnro & _
                                                        " AND ternro = " & rsLicencias!empleado & _
                                                        " AND anrrangnro = " & rsRango!anrrangnro
                                                End If
                                                objConn.Execute StrSql, , adExecuteNoRecords
                                                
                                                monto_total = 0
                                                cant_total = 0
                                            End If
                                            objRs.Close
                                            
                                        End If
                                    End If
                                    
                                End If
                                
                                rsEstructura.MoveNext
                            Loop
                    
                rsLicencias.MoveNext
            Loop
                    
            ' Actualizo el progreso
            Progreso = Progreso + IncPorc
            StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
            objconnProgreso.Execute StrSql, , adExecuteNoRecords
                    
siguientelegajo:
            rsFiltro.MoveNext
        Loop
        
        rsFactor.MoveNext
    Loop
           
    rsRango.MoveNext
Loop
SumPorcTiempo = CDbl(SumPorcTiempo) + CDbl(PorcTiempo)

Exit Sub

CE:
    HuboErrorTipo = True
    HuboError = True
    Flog.writeline Espacios(Tabulador * 1) & "Error " & Err.Description
End Sub


Public Sub Licencias_old(Nro_Analisis As Long, Filtrar As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Analisis para tipo de factor 6 (Licencias de Rhpro)
' Autor      : FGZ
' Fecha      : 15/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

'Variables locales
Dim Desde As Date
Dim Hasta As Date
Dim horas As Single
Dim NroCab As Long
Dim tercero As Long

Dim cantdiasper As Integer
Dim cantdiasran As Integer
Dim porcentaje As Single
Dim monto_saldo As Single
Dim monto_total As Single
Dim cant_saldo As Single
Dim cant_total As Single
Dim cubvalor1 As Single
Dim cubvalor2 As Single
Dim cubvalor3 As Single
Dim cubvalor4 As Single

'Variables para los first y last
Dim PrimerFactOri As Boolean
Dim TipoEstr As Long
Dim EstrAct As Long
Dim FactOri As Long
Dim Ok As Boolean
Dim cant_flt As Long
Dim CantRangos As Integer

' Fechas desde y hasta a analizar por el proceso entero
Dim Inicio_Per_Analizado As Date
Dim Fin_Per_Analizado As Date

' Fechas parciales que se estan analizando
Dim Dia_Inicio_Per_Analizado As Date
Dim Dia_Fin_Per_Analizado As Date

Dim rs As New ADODB.Recordset
Dim rsLicencias As New ADODB.Recordset
Dim rsAnrCab As New ADODB.Recordset
Dim rsFiltro As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset

'Código -------------------------------------------------------------------

'Abro la conexion para Nexus
'OpenConnection strConexionNexus, objConnNexus

On Error GoTo CE:

'Obtengo la cabecera
Call ObtenerCabecerayFiltro(Nro_Analisis, rsAnrCab, Filtrar, rsFiltro, cant_flt, Ok)

If Not Ok Then
    Exit Sub
End If

'Estoy reprocesando
Call PurgarCubo(Nro_Analisis, 6)


' Obtengo los rangos del analisis
StrSql = "SELECT * FROM anrrangofec" & _
         " WHERE anrrangofec.anrcabnro = " & rsAnrCab!anrcabnro & _
         " AND anrrangorepro = -1 "
OpenRecordset StrSql, rsRango
                    
Progreso = 0

CantRangos = 0
If Not rsRango.EOF Then
    CantRangos = rsRango.RecordCount
End If

                    
Do While Not rsRango.EOF
    Fin_Per_Analizado = rsRango!anrrangfechasta
    Inicio_Per_Analizado = rsRango!anrrangfecdesde

    Dia_Inicio_Per_Analizado = Inicio_Per_Analizado
    Dia_Fin_Per_Analizado = Fin_Per_Analizado

    ' -----------------------------------------------------------------
    '/* Recorre para el analisis los l de tipos de horas configurados */
    StrSql = "SELECT * FROM anrcab_fact" & _
        " INNER JOIN anrfact_ori ON anrfact_ori.facnro = anrcab_fact.facnro" & _
        " AND anrfact_ori.tipfacnro = 6" & _
        " INNER JOIN anrfactor ON anrfactor.facnro = anrcab_fact.facnro" & _
        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
        " ORDER BY anrfact_ori.facnro"
    OpenRecordset StrSql, rsFactor
        
    If Not rsFactor.EOF Then
        'Para el simular el first_of
        PrimerFactOri = True
        'Para el simular el last_of en la tabla anrfact_ori
        FactOri = rsFactor!facnro
    End If
        
    CantFactor = 0
    If Not rsFactor.EOF Then
        CantFactor = rsFactor.RecordCount
    End If
        
    ' obtengo el conjunto de legajos a procesar
    Call ObtenerLegajos(4, Filtrar, rsAnrCab!anrcabnro, rsFiltro, rsAnrCab!anrcabfecdesde, rsAnrCab!anrcabfechasta)
        
    'Recorro las Licencias que entran el el rango actual analizado
    Do While Not rsFactor.EOF
        
'        'Busco si es factor totalizador
'        StrSql = "SELECT * " & _
'                " FROM   anrfact_tot, anrcab_fact" & _
'                " WHERE  anrfact_tot.facnro = " & rsFactor!facnro & _
'                " AND    anrcab_fact.facnro   = anrfact_tot.facnro " & _
'                " AND    anrcab_fact.anrcabnro = " & rsAnrCab!anrcabnro
'        OpenRecordset StrSql, rsFactorTotalizador
'
'
'        'Busco si es factor totalizador
'        'StrSql = "SELECT * FROM anrfact_tot" & _
'        '        " WHERE facnro = " & rsFactor!facnro
'        'OpenRecordset StrSql, rsFactorTotalizador
'
'        If Not rsFactorTotalizador.EOF Then
'            Totaliza = True
'            ' codigo de factor con el cual se inserta en el cubo
'            FactorTotalizador = rsFactorTotalizador!facnrotot
'        Else
'            Totaliza = False
'            FactorTotalizador = 0
'        End If
            
        ' voy al primer legajo nuevamente
        rsFiltro.MoveFirst
        
        Do While Not rsFiltro.EOF
            If (cant_flt > 0) And (rsFiltro!cant_te < cant_flt) Then
                GoTo siguientelegajo
            End If
        
            StrSql = "SELECT * FROM emp_lic " & _
                " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
                " INNER JOIN gti_acumdiario ON gti_acumdiario.thnro = tipdia.thnro" & _
                " WHERE gti_acumdiario.ternro = " & rsFiltro!Ternro & _
                " AND tipdia.tdnro = " & rsFactor!faccodorig & _
                " AND emp_lic.elfechadesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
                " AND emp_lic.elfechahasta >= " & ConvFecha(Dia_Inicio_Per_Analizado)
        
            'StrSql = "SELECT * FROM gti_acumdiario " & _
            '         " WHERE adfecha <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
            '         " AND adfecha >= " & ConvFecha(Dia_Inicio_Per_Analizado) & _
            '         " AND thnro = " & rsFactor!faccodorig & _
            '         " AND ternro = " & rsFiltro!Ternro & _
            '         " ORDER BY ternro"

            OpenRecordset StrSql, rsLicencias
            
            If Not rsLicencias.EOF Then
                IncPorc = CantRangos * (((100 / CantFactor) * (100 / rsLicencias.RecordCount)) / 100) / 100
            End If
        
            Do While Not rsLicencias.EOF
    
                            ' Obtengo las estructuras
                            Call ObtenerEstructuras(Filtrar, rsLicencias!Ternro, Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado, rsEstructura)
                            
                            If Not rsEstructura.EOF Then
                                TipoEstr = rsEstructura!tenro
                                EstrAct = rsEstructura!estrnro
                            End If
                        
                            Do While Not rsEstructura.EOF
        
                                If PrimerFactOri Then
                                    cantdiasper = DateDiff("d", Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado) + 1
                                    monto_total = 0
                                    cant_total = 0
                                    cant_saldo = 0
                                    PrimerFactOri = False
                                End If
                                               
                                '/* Acumulo por Factor */
                                ' FGZ 21/07/2003
                                ' busco la cantidad de horas justificadas en el gti_acumdiario
                                 
                                 'If IsNull(rsEstructura!htethasta) Then
                                 '   StrSql = "SELECT sum(adcanthoras) as suma FROM gti_acumdiario " & _
                                 '       " WHERE ternro = " & rsLicencias!Ternro & _
                                 '       " AND thnro = " & rsLicencias!thnro & _
                                 '       " AND adfecha >= " & ConvFecha(rsEstructura!htetdesde) & _
                                 '       " AND adfecha <= " & ConvFecha(Dia_Fin_Per_Analizado)
                                 'Else
                                 '   StrSql = "SELECT sum(adcanthoras) as suma FROM gti_acumdiario " & _
                                 '       " WHERE ternro = " & rsLicencias!Ternro & _
                                 '       " AND thnro = " & rsLicencias!thnro & _
                                 '       " AND adfecha >= " & ConvFecha(rsEstructura!htetdesde) & _
                                 '       " AND adfecha <= " & ConvFecha(rsEstructura!htethasta)
                                 'End If
                                'OpenRecordset StrSql, rsAD
                                'If Not rsAD.EOF Then
                                '    If Not IsNull(rsAD!suma) Then
                                '        monto_total = monto_total + rsAD!suma
                                '    End If
                                'End If
                                
                                
                                monto_total = monto_total + rsLicencias!adcanthoras
                                cant_total = cant_total
                                
                                '/* Calculo los dias de rango entre las fechas del rango y
                                ' el his_estruct para proporcionar*/
                                If rsFactor!facpropor = -1 Then
                                    If rsEstructura!htetdesde < Dia_Inicio_Per_Analizado Then
                                            If rsEstructura!htethasta < Dia_Fin_Per_Analizado And (Not IsNull(rsEstructura!htethasta)) Then
                                                cantdiasran = DateDiff("d", Dia_Inicio_Per_Analizado, rsEstructura!htethasta) + 1
                                            Else
                                                cantdiasran = DateDiff("d", Dia_Inicio_Per_Analizado, Dia_Fin_Per_Analizado) + 1
                                            End If
                                    Else
                                        If (rsEstructura!htethasta < Dia_Fin_Per_Analizado) And (Not IsNull(rsEstructura!htethasta)) Then
                                            cantdiasran = DateDiff("d", rsEstructura!htetdesde, rsEstructura!htethasta) + 1
                                        Else
                                            cantdiasran = DateDiff("d", rsEstructura!htetdesde, Dia_Fin_Per_Analizado) + 1
                                        End If
                                    End If
                                    
                                    '/* Porcentaje segun la cant. de dias en la his_estrutura */
                                    porcentaje = cantdiasran * 100 / cantdiasper
                                    
                                    'If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    If Ultimo(rsLicencias) Or Last_OF_estrnro() Then
                                        ' se supone que no se proporciona
                                        cubvalor1 = monto_total
                                        cubvalor2 = cant_total
                                        
                                        StrSql = "SELECT * FROM anrcubo" & _
                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                            " AND facnro = " & rsFactor!facnro & _
                                            " AND tenro = " & rsEstructura!tenro & _
                                            " AND estrnro = " & rsEstructura!estrnro & _
                                            " AND ternro = " & rsAcumDiario!Ternro & _
                                            " AND anrrangnro = " & rsRango!anrrangnro
                                        OpenRecordset StrSql, rs
            
                                        'Si el cubo no existe lo creo
                                        If rs.EOF Then
                                        '/* Creo el cubo */
                                            StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                rsEstructura!estrnro & "," & rsFactor!facnro & "," & _
                                                rsEstructura!tenro & "," & rsLicencias!Ternro & ",1"
                                        End If
                                        
                                        If Not rs.EOF Then
                                            monto_saldo = (monto_total - cubvalor1 - rs!anrcubvalor1)
                                            cant_saldo = (cant_total - cubvalor2 - cant_saldo)
                                        End If
                                        'monto_saldo = (monto_total - cubvalor1 - monto_saldo)
                                        monto_total = 0
                                        cant_total = 0
                                        
                                        '* Para que no quede saldo cuando proporciona */
                                        If monto_saldo <= 1 And monto_saldo > 0 Then
                                            'cubvalor1 = cubvalor1 + monto_saldo
                                        End If
                                        
                                        If cant_saldo <= 1 And cant_saldo > 0 Then
                                            'cubvalor2 = cubvalor2 + cant_saldo
                                        End If
                                           
                                        'Si existe el cubo entonces actualizo
                                        If Not rs.EOF Then
                                            StrSql = "UPDATE anrcubo SET" & _
                                                " anrcubvalor1 = " & Round(rs!anrcubvalor1 + cubvalor1, 2) & _
                                                " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                " AND facnro = " & rsFactor!facnro & _
                                                " AND tenro = " & rsEstructura!tenro & _
                                                " AND estrnro = " & rsEstructura!estrnro & _
                                                " AND ternro = " & rsLicencias!Ternro & _
                                                " AND anrrangnro = " & rsRango!anrrangnro
                                        Else
                                            StrSql = StrSql & "," & cubvalor1 & "," & cubvalor2 & ")"
                                        End If
                                        
                                        objConn.Execute StrSql, , adExecuteNoRecords
                                        
'                                        'FZG 25/07/2003
'                                        'Actualizo Totalizador
'                                        If Totaliza Then
'                                            StrSql = "SELECT * FROM anrcubo" & _
'                                                " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'                                                " AND facnro = " & FactorTotalizador & _
'                                                " AND tenro = " & rsEstructura!tenro & _
'                                                " AND estrnro = " & rsEstructura!estrnro & _
'                                                " AND ternro = " & rsLicencias!Ternro & _
'                                                " AND anrrangnro = " & rsRango!anrrangnro
'                                            OpenRecordset StrSql, rsTot
'
'                                            If rsTot.EOF Then
'                                                ' Creo el cubo
'                                                StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
'                                                    ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
'                                                    ",anrcubvalor1,anrcubvalor2) VALUES (" & _
'                                                    rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
'                                                    rsEstructura!estrnro & "," & FactorTotalizador & "," & _
'                                                    rsEstructura!tenro & "," & rsLicencias!Ternro & ",1," & _
'                                                    cubvalor1 & "," & cubvalor2 & ")"
'                                            Else
'                                                StrSql = "UPDATE anrcubo SET" & _
'                                                    " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
'                                                    " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
'                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'                                                    " AND facnro = " & FactorTotalizador & _
'                                                    " AND tenro = " & rsEstructura!tenro & _
'                                                    " AND estrnro = " & rsEstructura!estrnro & _
'                                                    " AND ternro = " & rsLicencias!Ternro & _
'                                                    " AND anrrangnro = " & rsRango!anrrangnro
'                                            End If
'                                            objConn.Execute StrSql, , adExecuteNoRecords
'                                        End If
                                        
                                    End If
                                Else
                                    '/* Si no proporciona tomo al 100% y la ultima his_estruc del rango*/
                                    porcentaje = 100
                                    'If Last_OF(rsFactor!facnro) Or Last_OF(rsEstructura!estrnro) Then
                                    'If Last_OF_Factor() Or Last_OF_estrnro() Then
                                    If Ultimo(rsLicencias) Or Last_OF_estrnro() Then
                                        If Not Last_OF_tenro() Then
                                            monto_total = 0
                                            cant_total = 0
                                        Else
                                        '/*Busco la ultima his_estr dentro del rango*/
                                            StrSql = "SELECT * FROM his_estructura " & _
                                                " WHERE his_estructura.ternro = " & rsLicencias!Ternro & _
                                                " AND his_estructura.tenro = " & rsEstructura!tenro & _
                                                " AND his_estructura.htetdesde <= " & ConvFecha(Dia_Fin_Per_Analizado) & _
                                                " AND (his_estructura.htethasta >= " & ConvFecha(Dia_Inicio_Per_Analizado) & _
                                                " OR his_estructura.htethasta IS NULL) "
                                            OpenRecordset StrSql, objRs
                                            objRs.MoveLast
                                            
                                            If Not objRs.EOF Then
                                            
                                                StrSql = "SELECT * FROM anrcubo" & _
                                                    " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                    " AND facnro = " & rsFactor!facnro & _
                                                    " AND tenro = " & objRs!tenro & _
                                                    " AND estrnro = " & objRs!estrnro & _
                                                    " AND ternro = " & rsLicencias!Ternro & _
                                                    " AND anrrangnro = " & rsRango!anrrangnro
                                                OpenRecordset StrSql, rs
                                                
                                                cubvalor1 = monto_total * porcentaje / 100
                                                cubvalor2 = cant_total * porcentaje / 100
                                                
                                                If rs.EOF Then
                                                    '/* Creo el cubo */
                                                    StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
                                                        ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
                                                        ",anrcubvalor1,anrcubvalor2) VALUES (" & _
                                                        rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
                                                        objRs!estrnro & "," & rsFactor!facnro & "," & _
                                                        objRs!tenro & "," & rsLicencias!Ternro & ",1" & _
                                                        "," & cubvalor1 & "," & cubvalor2 & ")"
                                                Else
                                                    StrSql = "UPDATE anrcubo SET" & _
                                                        " anrcubvalor1 = " & rs!anrcubvalor1 + cubvalor1 & _
                                                        " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
                                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
                                                        " AND facnro = " & rsFactor!facnro & _
                                                        " AND tenro = " & objRs!tenro & _
                                                        " AND estrnro = " & objRs!estrnro & _
                                                        " AND ternro = " & rsLicencias!Ternro & _
                                                        " AND anrrangnro = " & rsRango!anrrangnro
                                                End If
                                                objConn.Execute StrSql, , adExecuteNoRecords
                                                
'                                                'FZG 25/07/2003
'                                                'Actualizo Totalizador
'                                                If Totaliza Then
'                                                    StrSql = "SELECT * FROM anrcubo" & _
'                                                        " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'                                                        " AND facnro = " & FactorTotalizador & _
'                                                        " AND tenro = " & rsEstructura!tenro & _
'                                                        " AND estrnro = " & rsEstructura!estrnro & _
'                                                        " AND ternro = " & rsLicencias!Ternro & _
'                                                        " AND anrrangnro = " & rsRango!anrrangnro
'                                                    OpenRecordset StrSql, rsTot
'
'                                                    If rsTot.EOF Then
'                                                        ' Creo el cubo
'                                                        StrSql = "INSERT INTO anrcubo(anrcabnro,anrcubmanual" & _
'                                                            ",anrrangnro,estrnro,facnro,tenro,Ternro,tipnro" & _
'                                                            ",anrcubvalor1,anrcubvalor2) VALUES (" & _
'                                                            rsAnrCab!anrcabnro & ",0," & rsRango!anrrangnro & "," & _
'                                                            rsEstructura!estrnro & "," & FactorTotalizador & "," & _
'                                                            rsEstructura!tenro & "," & rsLicencias!Ternro & ",1," & _
'                                                            cubvalor1 & "," & cubvalor2 & ")"
'                                                    Else
'                                                        StrSql = "UPDATE anrcubo SET" & _
'                                                            " anrcubvalor1 = " & Round(rsTot!anrcubvalor1 + cubvalor1, 2) & _
'                                                            " ,anrcubvalor2 = " & rs!anrcubvalor2 + cubvalor2 & _
'                                                            " WHERE anrcabnro = " & rsAnrCab!anrcabnro & _
'                                                            " AND facnro = " & FactorTotalizador & _
'                                                            " AND tenro = " & rsEstructura!tenro & _
'                                                            " AND estrnro = " & rsEstructura!estrnro & _
'                                                            " AND ternro = " & rsLicencias!Ternro & _
'                                                            " AND anrrangnro = " & rsRango!anrrangnro
'                                                    End If
'                                                    objConn.Execute StrSql, , adExecuteNoRecords
'                                                End If
                                                
                                                monto_total = 0
                                                cant_total = 0
                                                
                                                
                                            End If
                                            objRs.Close
                                            
                                        End If
                                    End If
                                    
                                End If
                                
                                rsEstructura.MoveNext
                            Loop
                    
                rsLicencias.MoveNext
            Loop
                    
siguientelegajo:
            rsFiltro.MoveNext
        Loop
        
        rsFactor.MoveNext
    Loop
           
           
    ' Actualizo el progreso
    Progreso = Progreso + IncPorc
    StrSql = "UPDATE batch_proceso SET bprcprogreso = " & CInt(Progreso) & " WHERE bpronro = " & NroProceso
    objconnProgreso.Execute StrSql, , adExecuteNoRecords
           
    rsRango.MoveNext
Loop

Exit Sub

CE:
    HuboErrorTipo = True
    HuboError = True
    Flog.writeline Espacios(Tabulador * 1) & "Error " & Err.Description
End Sub


Public Function CantidadDeDias(ByVal PeriodoDesde As Date, ByVal PeriodoHasta As Date, ByVal Desde As Date, ByVal Hasta As Date) As Integer
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias que caen dentro de un periodo (especificado como un
'              rango de fechas) .
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim FechaInicioAuxiliar As Date
Dim FechaFinAuxiliar As Date
    
    FechaInicioAuxiliar = IIf(Desde > PeriodoDesde, Desde, PeriodoDesde)
    If Not IsNull(Hasta) Then
        FechaFinAuxiliar = Hasta
        FechaFinAuxiliar = IIf(Hasta > PeriodoHasta, PeriodoHasta, Hasta)
    Else
        'FechaFinAuxiliar = PeriodoHasta
        FechaFinAuxiliar = IIf(Hasta > PeriodoHasta, PeriodoHasta, Hasta)
    End If
    
    CantidadDeDias = DateDiff("d", FechaInicioAuxiliar, FechaFinAuxiliar) + 1

End Function

